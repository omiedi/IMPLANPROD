VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FORPRE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestación de Cámaras"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FORPRE04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
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
      Left            =   1155
      TabIndex        =   60
      Text            =   " "
      Top             =   210
      Width           =   2325
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
      Height          =   600
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Top             =   525
      Width           =   2925
   End
   Begin VB.TextBox Text8 
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
      Height          =   285
      Left            =   6090
      TabIndex        =   58
      Top             =   840
      Width           =   2340
   End
   Begin VB.TextBox Text9 
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
      Height          =   285
      Left            =   9660
      TabIndex        =   57
      Top             =   210
      Width           =   1125
   End
   Begin VB.CommandButton Command16 
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
      Left            =   10755
      Picture         =   "FORPRE04.frx":0442
      TabIndex        =   56
      Top             =   525
      Width           =   175
   End
   Begin VB.TextBox Text11 
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
      Height          =   285
      Left            =   9120
      TabIndex        =   55
      Top             =   840
      Width           =   1650
   End
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
      Left            =   9660
      TabIndex        =   54
      Top             =   525
      Width           =   1125
   End
   Begin VB.TextBox Text5 
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
      Height          =   285
      Left            =   5145
      TabIndex        =   53
      Top             =   525
      Width           =   3300
   End
   Begin VB.CommandButton Command9 
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
      Left            =   3470
      Picture         =   "FORPRE04.frx":074C
      TabIndex        =   52
      Top             =   200
      Width           =   175
   End
   Begin VB.TextBox Text18 
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
      Height          =   285
      Left            =   5145
      TabIndex        =   51
      Top             =   210
      Width           =   3300
   End
   Begin VB.TextBox Text7 
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
      Height          =   285
      Left            =   5145
      TabIndex        =   50
      Top             =   840
      Width           =   870
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8D8C0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   8160
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "FORPRE04.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   155
         Picture         =   "FORPRE04.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   725
         Picture         =   "FORPRE04.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   0
      Width           =   1275
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
         Height          =   650
         Left            =   105
         Picture         =   "FORPRE04.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Acceso a Base de Datos de Clientes"
         Top             =   1575
         Width           =   640
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Save"
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
         Picture         =   "FORPRE04.frx":2812
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Guardar "
         Top             =   2310
         Width           =   640
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "FORPRE04.frx":2B1C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Alta de un Nuevo Presupuesto por Copia de uno Existente"
         Top             =   3690
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPRE04.frx":2C66
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Ver y Re-Imprimir"
         Top             =   4380
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   7245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
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
         Picture         =   "FORPRE04.frx":2F70
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprime Presupuesto"
         Top             =   5835
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "List"
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
         Picture         =   "FORPRE04.frx":35DA
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Listados Clasificados"
         Top             =   5145
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
         Picture         =   "FORPRE04.frx":38E4
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6525
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
         Height          =   640
         Left            =   105
         Picture         =   "FORPRE04.frx":3D26
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPRE04.frx":4030
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPRE04.frx":417A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Picture         =   "FORPRE04.frx":4484
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Alta de un Nuevo Presupuesto"
         Top             =   3000
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FORPRE04.frx":45CE
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10920
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11774
      _Version        =   327681
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Presupuesto"
      TabPicture(0)   =   "FORPRE04.frx":64F0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Rubros / Cotización"
      TabPicture(1)   =   "FORPRE04.frx":650C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame20"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Aprobación / Observaciones"
      TabPicture(2)   =   "FORPRE04.frx":6528
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame13"
      Tab(2).Control(1)=   "Frame12"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame13 
         Caption         =   "Aprobación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   103
         Top             =   600
         Width           =   9975
         Begin VB.TextBox Text15 
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
            Left            =   6120
            TabIndex        =   106
            Text            =   " "
            Top             =   280
            Width           =   3045
         End
         Begin VB.TextBox Text6 
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
            Left            =   1680
            TabIndex        =   104
            Text            =   " "
            Top             =   280
            Width           =   3045
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Autorizó:"
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
            Left            =   4080
            TabIndex        =   107
            Top             =   405
            Width           =   1905
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Controló:"
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
            Left            =   -360
            TabIndex        =   105
            Top             =   405
            Width           =   1905
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   -74640
         TabIndex        =   101
         Top             =   1560
         Width           =   9975
         Begin VB.TextBox Text29 
            Height          =   4095
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   102
            Top             =   360
            Width           =   9495
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -70560
         TabIndex        =   100
         Top             =   600
         Width           =   6135
         Begin VB.Frame Frame23 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            TabIndex        =   149
            Top             =   4150
            Width           =   5895
            Begin VB.Frame Frame15 
               Caption         =   "Valorización en:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               TabIndex        =   164
               Top             =   550
               Width           =   1935
               Begin VB.TextBox TICAMBIS 
                  Alignment       =   1  'Right Justify
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
                  Height          =   300
                  Left            =   465
                  MaxLength       =   14
                  TabIndex        =   165
                  Top             =   540
                  Width           =   1065
               End
               Begin VB.Label Label81 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pesos"
                  BeginProperty Font 
                     Name            =   "Arial Black"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   195
                  TabIndex        =   167
                  Top             =   150
                  Width           =   1575
               End
               Begin VB.Label Label80 
                  BackStyle       =   0  'Transparent
                  Caption         =   "T-C:"
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
                  Left            =   45
                  TabIndex        =   166
                  Top             =   615
                  Width           =   435
               End
            End
            Begin VB.TextBox Text27 
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
               Left            =   3240
               TabIndex        =   152
               Top             =   620
               Width           =   1185
            End
            Begin VB.TextBox Text23 
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
               Left            =   3240
               TabIndex        =   151
               Top             =   320
               Width           =   1185
            End
            Begin VB.Label PORIVA 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2480
               TabIndex        =   171
               Top             =   680
               Width           =   495
            End
            Begin VB.Label PORIVA 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   2480
               TabIndex        =   170
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label82 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3000
               TabIndex        =   169
               Top             =   680
               Width           =   255
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3000
               TabIndex        =   168
               Top             =   360
               Width           =   255
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total General:"
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
               Left            =   2550
               TabIndex        =   163
               Top             =   1450
               Width           =   1905
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   162
               Top             =   1380
               Width           =   1185
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   161
               Top             =   920
               Width           =   1185
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   160
               Top             =   920
               Width           =   1185
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I.B.:"
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
               Left            =   1200
               TabIndex        =   159
               Top             =   1020
               Width           =   1905
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   520
               TabIndex        =   158
               Top             =   680
               Width           =   1905
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   157
               Top             =   620
               Width           =   1185
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   156
               Top             =   320
               Width           =   1185
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   520
               TabIndex        =   155
               Top             =   360
               Width           =   1905
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   154
               Top             =   0
               Width           =   1185
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Neto:"
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
               Left            =   2520
               TabIndex        =   153
               Top             =   80
               Width           =   1905
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen          de Cotización  "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   0
               TabIndex        =   150
               Top             =   0
               Width           =   1545
            End
         End
         Begin VB.Frame Frame21 
            BorderStyle     =   0  'None
            Caption         =   " "
            Height          =   400
            Left            =   120
            TabIndex        =   144
            Top             =   3615
            Width           =   5895
            Begin VB.TextBox Text26 
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
               Left            =   3240
               TabIndex        =   148
               Top             =   80
               Width           =   1185
            End
            Begin VB.TextBox Text20 
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
               Left            =   2280
               TabIndex        =   146
               Top             =   80
               Width           =   660
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   172
               Top             =   80
               Width           =   1185
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   2880
               TabIndex        =   147
               Top             =   180
               Width           =   255
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Descuentos -    Bonificaciones:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   0
               TabIndex        =   145
               Top             =   45
               Width           =   1545
            End
         End
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
            Caption         =   "Frame18"
            Height          =   765
            Left            =   120
            TabIndex        =   133
            Top             =   2760
            Width           =   5895
            Begin VB.TextBox Text16 
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
               Left            =   2280
               TabIndex        =   142
               Top             =   435
               Width           =   660
            End
            Begin VB.TextBox Text44 
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
               Left            =   2280
               TabIndex        =   140
               Top             =   120
               Width           =   660
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   5880
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   2880
               TabIndex        =   143
               Top             =   540
               Width           =   255
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   2880
               TabIndex        =   141
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   139
               Top             =   435
               Width           =   1185
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   138
               Top             =   435
               Width           =   1185
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   137
               Top             =   120
               Width           =   1185
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "M.Obra:"
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
               Left            =   375
               TabIndex        =   136
               Top             =   525
               Width           =   1905
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materiales:"
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
               Left            =   375
               TabIndex        =   135
               Top             =   210
               Width           =   1905
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gastos       Generales:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   -740
               TabIndex        =   134
               Top             =   30
               Width           =   1905
            End
         End
         Begin VB.Frame Frame14 
            BorderStyle     =   0  'None
            Height          =   990
            Left            =   120
            TabIndex        =   126
            Top             =   1785
            Width           =   5895
            Begin VB.Line Line4 
               X1              =   0
               X2              =   5880
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   132
               Top             =   555
               Width           =   1185
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Montaje -    Instalación:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   -830
               TabIndex        =   131
               Top             =   200
               Width           =   2025
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   130
               ToolTipText     =   "Doble Click para Abrir Cotización de Transporte"
               Top             =   555
               Width           =   1185
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   129
               ToolTipText     =   "Doble Click para Abrir Cotización de Mano de Obra"
               Top             =   240
               Width           =   1185
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Transporte:"
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
               Left            =   1200
               TabIndex        =   128
               Top             =   645
               Width           =   1905
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mano de Obra:"
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
               Left            =   1200
               TabIndex        =   127
               Top             =   360
               Width           =   1905
            End
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   120
            TabIndex        =   113
            Top             =   240
            Width           =   5895
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materiales    y Componentes:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   -640
               TabIndex        =   125
               Top             =   0
               Width           =   2025
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   124
               ToolTipText     =   "Doble Click para Abrir Cotización de Varios"
               Top             =   1260
               Width           =   1185
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   123
               ToolTipText     =   "Doble Click para Abrir Cotización de Equipos"
               Top             =   945
               Width           =   1185
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   122
               ToolTipText     =   "Doble Click para Abrir Cotización de Accesorios"
               Top             =   630
               Width           =   1185
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   121
               ToolTipText     =   "Doble Click para Abrir Cotización de Puertas"
               Top             =   315
               Width           =   1185
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   120
               ToolTipText     =   "Doble Click para Abrir Cotización de Cámara"
               Top             =   0
               Width           =   1185
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Varios:"
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
               Left            =   1200
               TabIndex        =   119
               Top             =   1335
               Width           =   1905
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Accesorios:"
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
               Left            =   1200
               TabIndex        =   118
               Top             =   705
               Width           =   1905
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Puertas:"
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
               Left            =   1200
               TabIndex        =   117
               Top             =   405
               Width           =   1905
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cámara:"
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
               Left            =   1200
               TabIndex        =   116
               Top             =   120
               Width           =   1905
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipos:"
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
               Left            =   1200
               TabIndex        =   115
               Top             =   1035
               Width           =   1905
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   114
               Top             =   1260
               Width           =   1185
            End
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   6000
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   6000
            Y1              =   3600
            Y2              =   3600
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Dimensiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74880
         TabIndex        =   80
         Top             =   600
         Width           =   4215
         Begin VB.Frame Frame8 
            Caption         =   "MEDIDAS CALCULADAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2355
            Left            =   120
            TabIndex        =   91
            Top             =   3330
            Width           =   3945
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1320
               TabIndex        =   111
               Top             =   1785
               Width           =   1185
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1320
               TabIndex        =   110
               Top             =   1350
               Width           =   1185
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1320
               TabIndex        =   109
               Top             =   915
               Width           =   1185
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1320
               TabIndex        =   108
               Top             =   480
               Width           =   1185
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Piso:"
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
               Index           =   0
               Left            =   -630
               TabIndex        =   99
               Top             =   1455
               Width           =   1905
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Techo:"
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
               Left            =   -645
               TabIndex        =   98
               Top             =   1050
               Width           =   1905
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Volumen(int):"
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
               Left            =   -645
               TabIndex        =   97
               Top             =   585
               Width           =   1905
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lateral:"
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
               Left            =   -630
               TabIndex        =   96
               Top             =   1905
               Width           =   1905
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "m2"
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
               Left            =   2640
               TabIndex        =   95
               Top             =   1920
               Width           =   1905
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "m3"
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
               Left            =   2640
               TabIndex        =   94
               Top             =   600
               Width           =   1905
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "m2"
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
               Left            =   2640
               TabIndex        =   93
               Top             =   1065
               Width           =   1905
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "m2"
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
               Left            =   2640
               TabIndex        =   92
               Top             =   1470
               Width           =   1905
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "MEDIDAS EXTERIORES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   3945
            Begin VB.TextBox Text22 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   840
               TabIndex        =   173
               Top             =   1830
               Width           =   1185
            End
            Begin VB.TextBox Text19 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   840
               TabIndex        =   82
               Top             =   440
               Width           =   1185
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   840
               TabIndex        =   84
               Top             =   1350
               Width           =   1185
            End
            Begin VB.TextBox Text14 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   840
               TabIndex        =   83
               Top             =   870
               Width           =   1185
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2600
               TabIndex        =   179
               Top             =   440
               Width           =   1185
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2600
               TabIndex        =   178
               Top             =   870
               Width           =   1185
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2600
               TabIndex        =   177
               Top             =   1350
               Width           =   1185
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2600
               TabIndex        =   176
               Top             =   1830
               Width           =   1185
            End
            Begin VB.Label Label85 
               BackStyle       =   0  'Transparent
               Caption         =   "mm"
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
               Left            =   2080
               TabIndex        =   175
               Top             =   1930
               Width           =   1905
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Espesor:"
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
               Left            =   -1100
               TabIndex        =   174
               Top             =   1930
               Width           =   1905
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Largo:"
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
               Left            =   -1100
               TabIndex        =   90
               Top             =   585
               Width           =   1905
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ancho:"
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
               Left            =   -1100
               TabIndex        =   89
               Top             =   1005
               Width           =   1905
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Alto:"
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
               Left            =   -1100
               TabIndex        =   88
               Top             =   1470
               Width           =   1905
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "mts."
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
               Left            =   2080
               TabIndex        =   87
               Top             =   585
               Width           =   1905
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "mts."
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
               Left            =   2080
               TabIndex        =   86
               Top             =   1470
               Width           =   1905
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "mts."
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
               Left            =   2080
               TabIndex        =   85
               Top             =   1005
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6120
         Left            =   200
         TabIndex        =   8
         Top             =   420
         Width           =   10320
         Begin VB.Frame Frame11 
            Caption         =   "Representante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1000
            Left            =   7980
            TabIndex        =   11
            Top             =   4520
            Width           =   2325
            Begin VB.TextBox Text21 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1440
               TabIndex        =   4
               Text            =   " "
               Top             =   300
               Width           =   720
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   5
               Left            =   675
               TabIndex        =   3
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   5
               Left            =   90
               TabIndex        =   2
               Text            =   " "
               Top             =   315
               Width           =   630
            End
            Begin VB.Label DETEXT 
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
               Index           =   5
               Left            =   90
               TabIndex        =   13
               Top             =   630
               Width           =   2070
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%C:"
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
               Left            =   -345
               TabIndex        =   12
               Top             =   375
               Width           =   1710
            End
         End
         Begin VB.Frame Frame17 
            Height          =   615
            Left            =   7980
            TabIndex        =   75
            Top             =   5490
            Width           =   2325
            Begin VB.CommandButton BOTCON 
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
               Left            =   2025
               TabIndex        =   77
               Top             =   240
               Width           =   175
            End
            Begin VB.TextBox CUETEXT 
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
               Left            =   840
               TabIndex        =   76
               Text            =   " "
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Destino: "
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
               Left            =   -840
               TabIndex        =   78
               Top             =   320
               Width           =   1710
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Condiciones de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   0
            TabIndex        =   36
            Top             =   4515
            Width           =   4800
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   900
               IntegralHeight  =   0   'False
               ItemData        =   "FORPRE04.frx":6544
               Left            =   120
               List            =   "FORPRE04.frx":6546
               TabIndex        =   48
               Top             =   600
               Width           =   4530
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00FFFFC0&
               Height          =   345
               Left            =   120
               ScaleHeight     =   285
               ScaleWidth      =   4470
               TabIndex        =   46
               Top             =   315
               Width           =   4530
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "     It.     Descripcion                     Importe         Fecha"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   0
                  TabIndex        =   47
                  Top             =   20
                  Width           =   4950
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "L.Entrega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   4935
            TabIndex        =   34
            Top             =   4515
            Width           =   2955
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
               Height          =   1170
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   315
               Width           =   2745
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10305
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
               Left            =   9840
               Picture         =   "FORPRE04.frx":6548
               TabIndex        =   43
               Top             =   420
               Width           =   175
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
               Left            =   4800
               Picture         =   "FORPRE04.frx":6852
               TabIndex        =   42
               Top             =   420
               Width           =   175
            End
            Begin VB.TextBox Text13 
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
               Left            =   6090
               TabIndex        =   31
               Top             =   420
               Width           =   2040
            End
            Begin VB.TextBox Text12 
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
               Left            =   8820
               TabIndex        =   30
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text24 
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
               Left            =   3780
               TabIndex        =   1
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text25 
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
               Left            =   1155
               TabIndex        =   0
               Top             =   420
               Width           =   1935
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Orden de Compra:"
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
               Left            =   5025
               TabIndex        =   33
               Top             =   315
               Width           =   1065
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   8190
               TabIndex        =   32
               Top             =   465
               Width           =   585
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   3150
               TabIndex        =   29
               Top             =   465
               Width           =   585
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido de Cotización:"
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
               Left            =   -315
               TabIndex        =   10
               Top             =   315
               Width           =   1380
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -945
            TabIndex        =   37
            Top             =   840
            Width           =   11670
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00E0E0E0&
               Height          =   2955
               Left            =   9555
               ScaleHeight     =   2895
               ScaleWidth      =   1560
               TabIndex        =   72
               Top             =   525
               Width           =   1620
               Begin VB.TextBox Text38 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   74
                  Top             =   -30
                  Width           =   1650
               End
               Begin VB.Label Label90 
                  Height          =   435
                  Left            =   315
                  TabIndex        =   180
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Pesos"
                  BeginProperty Font 
                     Name            =   "Arial Black"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   0
                  TabIndex        =   73
                  Top             =   420
                  Width           =   1575
               End
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               Height          =   2955
               Left            =   1020
               ScaleHeight     =   2895
               ScaleWidth      =   795
               TabIndex        =   70
               Top             =   525
               Width           =   855
               Begin VB.TextBox Text36 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   71
                  Top             =   -30
                  Width           =   850
               End
            End
            Begin VB.TextBox Text37 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
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
               Left            =   0
               TabIndex        =   49
               Text            =   "Text36"
               Top             =   0
               Width           =   825
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFC0&
               Height          =   330
               Left            =   1020
               ScaleHeight     =   270
               ScaleWidth      =   10095
               TabIndex        =   44
               Top             =   210
               Width           =   10155
               Begin VB.Line Line24 
                  X1              =   0
                  X2              =   10080
                  Y1              =   315
                  Y2              =   315
               End
               Begin VB.Line Line23 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line2 
                  X1              =   840
                  X2              =   840
                  Y1              =   315
                  Y2              =   0
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   $"FORPRE04.frx":6B5C
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
                  Left            =   50
                  TabIndex        =   45
                  Top             =   30
                  Width           =   10200
               End
               Begin VB.Line Line3 
                  X1              =   860
                  X2              =   860
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line22 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
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
               Height          =   2955
               Left            =   1890
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   41
               Top             =   525
               Width           =   7680
            End
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   3795
         TabIndex        =   7
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   4215
         TabIndex        =   6
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.Label Label50 
      Caption         =   "Label50"
      Height          =   495
      Left            =   5400
      TabIndex        =   79
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. de Presupuesto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   69
      Top             =   110
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   -210
      TabIndex        =   68
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   4020
      TabIndex        =   67
      Top             =   630
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   4200
      TabIndex        =   66
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
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
      Height          =   330
      Left            =   8925
      TabIndex        =   65
      Top             =   315
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Validez:"
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
      Left            =   8805
      TabIndex        =   64
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   8550
      TabIndex        =   63
      Top             =   945
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   3990
      TabIndex        =   62
      Top             =   315
      Width           =   1110
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   330
      Left            =   3360
      TabIndex        =   61
      Top             =   105
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueOT 
         Caption         =   "Nuevo Presupuesto"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Consultas y Listados"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiCli 
      Caption         =   "Edición"
      Begin VB.Menu CoInEv 
         Caption         =   "Copia Interactiva de Presupuesto"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FORPRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$
Dim CAJENTRA%

Private Sub AyuCli_Click()
  Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
End Sub

Private Sub BOTCON_Click()
   '
   NEV1$ = TRIM(Text1)
   If NEV1$ = "" Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   If ULTREG&("ECUECON") < 1 Then
      Call MENSERR(24, "Falta Definir Plan de Cuentas Contables")
      Call CARCUEINGAS
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      CUETEXT.TEXT = Left$(RG1$, 12)
   End If
   '
End Sub
Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub CoInEv_Click()
  Call Command22_Click
End Sub

Private Sub CoInGr_Click()
    If SSTab1.Tab = 2 Then
        Call Command22_Click
    Else
        Call COMUNI("Para realizar esta Copia debe Situarse en la Solapa de Grupos.")
    End If
End Sub

Private Sub Command10_Click()
 If TRIM(Text1) <> "" Then
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text6.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text6.TEXT)
 End If
End Sub

Private Sub Command11_Click()
   '
 '  Call COMUNI("No Disponible en la presente versión")
'   If Val(TRIM$(Text1)) > 0 Then
'     Call SELECHO("DEUDOR1")
'     NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
'     If NCLIE% < 1 Then Exit Sub
'     '
'     NC% = NCLIE%
'     Text2 = TRIM$(Str$(NC%))
'     Text4 = RASOCLI$(NC%)
'   End If
   '
End Sub

Private Sub Command16_Click()
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text10.TEXT = VALACT1$("SELFECHA")
     End If
     FF% = FECHANUM(Text10.TEXT)
  End If
End Sub


Private Sub Command22_Click()
  '\\\OT-05-0391-WAR-09/06/05///
  Command22.Enabled = False
  If DERACCE%("ANUPRECO", "Generar Nuevo Presupuesto por Copia") < 2 Then
    GoTo 99
  End If
  '
  SSTab1.Tab = 0
9 Screen.MousePointer = 1
  '
  VDEF$ = Space$(64) + String$(6, 0)
  '
  OBX$ = OBJPLA$("DEACOPRE", VDEF$)
  If OBX$ = "" Then GoTo 99
  VDEF$ = OBX$
  '
  PRESU1$ = TRIM$(Mid$(OBX$, 1, 16))
  PRESU2$ = TRIM$(Mid$(OBX$, 17, 48))
  '//////////////VALIDACIONES////////////
  If TRIM$(PRESU1$) = "" Then
    Call COMUNI("Número de Presupuesto no valido")
    GoTo 9
  End If
  '
  If TRIM$(PRESU2$) = "" Then
    Call COMUNI("Número de Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU1$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU2$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If TRIM$(PRESU1$) = TRIM$(PRESU2$) Then
    Call MENSERR(24, "Copia Imposible.\Coinciden Presupuesto Origen y Destino.")
    GoTo 99
  End If
  'validación de que en el presupuesto no haya detalle.
  Call ABREPRECAM
  DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
  If DetaPresu.NoMatch = True Then
    GoTo 17
  Else
    If COMALTER%("El Presupuesto de Cámara " & PRESU2$ & " \ya Contiene Detalle, ¿Desea Sobre-Escribirlo?\\Si,Sobre-escribir.\\No, Cancelar.") = 2 Then
      On Error GoTo 0
      Screen.MousePointer = 1
      Command22.Enabled = True
      Exit Sub
    Else
      DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
      If DetaPresu.NoMatch = False Then
        DetaPresu.Delete
      End If
14    DetaPresu.FindNext " Nro_Presu= '" & PRESU2$ & "' "
      If DetaPresu.NoMatch = True Then
        GoTo 17
      Else
        DetaPresu.Delete
        GoTo 14
      End If
    End If
  End If
    '
17  Call ABREPRECAM
    EncaPresu.FindFirst "Nro_Presu= '" & PRESU1$ & "' "
    With EncaPresu
     USUA$ = !uSUCONF_Presu
     USUC$ = !UsuCont_Presu
     USUAP$ = !UsuApro_Presu
     NOC$ = !NroOcom_Presu
     FPCLIP$ = !FeOcom_Presu
     NUPECOT$ = !PediCoti_Presu
     FEPECOT = !FePedCot_Presu
     CAOPE# = !CanPed_Presu
     ESPECI$ = !Especi_Presu
     MONEOTR% = !Mone_Presu
     COMOTR# = !CotiMone_Presu
     LUENOTR$ = !LENTREGA_Presu
     VEOTR% = !Vended_Presu
     COVEOTR$ = !ComiVend_Presu
     On Error Resume Next
      OBSOTR$ = !Observa_Presu
     On Error GoTo 0
     PREVEOTR# = !PreVeN_PRESU
     LARPRE = !Largo_Presu
     ANCHOPRE = !Ancho_Presu
     ALTOPRE = !Alto_Presu
     ESPEPRE = !Espesor_Presu
     On Error Resume Next
      PORGGMAT# = !GGMat_Presu '% MATERIALES
     On Error GoTo 0
     On Error Resume Next
      PORGGOB# = !GGOba_Presu '% MANO DE OBRA
     On Error GoTo 0
     On Error Resume Next
      PORDEMO# = !DesBon_Presu '% DESCUENTOS Y MODIFICACIONES
     On Error GoTo 0
    End With
    '
    Call ABREPRECAM
    EncaPresu.FindFirst "Nro_Presu= '" & PRESU2$ & "' "
    If EncaPresu.NoMatch = False Then
      With EncaPresu
       .Edit
         !uSUCONF_Presu = USUA$ 'usuario de confección
         !UsuCont_Presu = USUAC$ 'usuario que controlo
         !UsuApro_Presu = USUAP$ 'usuario que aprobara
         !NroOcom_Presu = NOC$ 'numero de orden de compra del cliente
         !FeOcom_Presu = FPCLIP$  'fecha de pedido del cliente
         !PediCoti_Presu = NUPECOT$ 'Numero de Oferta
         !FePedCot_Presu = FEPECOT 'fecha de la oferta
         !Nuoritem_Presu = 0
         !CanPed_Presu = CAOPE#
         !Especi_Presu = ESPECI$
         !PreVeN_PRESU = PREVEOTR#
         !Mone_Presu = MONEOTR%
         !CotiMone_Presu = COMOTR#
         !LENTREGA_Presu = LUENOTR$
         !Ajusta_Presu = 0
         !Vended_Presu = VEOTR%
         !ComiVend_Presu = COVEOTR$
         !Observa_Presu = OBSOTR$
         !PreVeN_PRESU = PREVEOTR#
         !Largo_Presu = LARPRE
         !Ancho_Presu = ANCHOPRE
         !Alto_Presu = ALTOPRE
         !Espesor_Presu = ESPEPRE
         !GGMat_Presu = PORGGMAT# '% MATERIALES
         !GGOba_Presu = PORGGOB# '% MANO DE OBRA
         !DesBon_Presu = PORDEMO# '% DESCUENTOS Y MODIFICAC
       .Update
      End With
    End If
    'ahora copio la lista de grupos
    SQLX$ = " SELECT * FROM detapresu "
    SQLX$ = SQLX$ + " WHERE Nro_Presu= '" & PRESU1$ & "' "
    Set DETAPRESUTMP = PresuCam.OpenRecordset(SQLX$, 4)
    '
    With DETAPRESUTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 36
     Else
       Do While .EOF = False
        NURUB% = !rubro_det
        SUBRUB% = !SUBRUB_DET
        CODET% = !CODINTE_DET
        CODEXTET$ = !CODEXTE_DET
        CANMAT# = !Cantid_Det
        PREUMAT# = !PreUni_Det
        MONEDET% = !MONE_DET
        MED1 = !MED1_DET
        MED2 = !Med2_Det
        MED3 = !Med3_Det
        IMPODEP# = !Impo_Det
        DESCRIDET$ = !DESCRI_DET
        HORADET = !Hora_Det
        OPERADET% = !Opera_Det
        VALOHORDET# = !ValoHor_Det
        OBSGRU$ = !OBSERVA_GRUP
        With DetaPresu
          .AddNew
            !NRO_PRESU = PRESU2$
            !rubro_det = NURUB%
            !SUBRUB_DET = SUBRUB%
            !CODINTE_DET = CODET%
            !CODEXTE_DET = CODEXTET$
            !Cantid_Det = CANMAT#
            !PreUni_Det = PREUMAT#
            !MONE_DET = MONEDET%
            !MED1_DET = MED1
            !Med2_Det = MED2
            !Med3_Det = MED3
            !Impo_Det = IMPODEP#
            !DESCRI_DET = DESCRIDET$
            !Hora_Det = HORADET
            !Opera_Det = OPERADET%
            !ValoHor_Det = VALOHORDET#
            !OBSERVA_GRUP = OBSGRU$
          .Update
        End With
        .MoveNext
       Loop
       End If
    End With
    '
36  Call GENEINDIPRECA
    Text1 = TRIM$(PRESU2$)
    GoTo 99
    '
99  Screen.MousePointer = 1
    Command22.Enabled = True
    '\\\OT-05-0391-WAR-FIN///
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
        Command22.ToolTipText = "Alta de Nueva Presupuesto de Cámara por Copia de una Existente"
    End If
    If SSTab1.Tab = 2 Then
        Command22.ToolTipText = "Alta de Nueva Lista de Grupos a partir de una Lista Existente"
    End If
    '
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub Command7_Click()
    '
    If TRIM$(Text1) <> "" Then
        '
        FORIPRE$ = CONTROL$("", "LISTGRUP")
        '
        If TRIM$(FORIPRE$) <> "" Then
          '
          Screen.MousePointer = 11
          FECHDOC$ = Text9
          FECHFIN$ = Text10
          NUMEDOC$ = TRIM$(Text1)
          DENOOBRA$ = Text3
          DESTIDOC% = Val(Label22) 'CLIENTE
          RPLA$ = TRIM$(Text15)
          '
          CODPARAM$(1) = "RASO-CLI" 'CLIENTE
          CODPARAM$(2) = "REF-MAT1" 'DENOMINACION DE LA ESP. DE VENTA
          CODPARAM$(3) = "NUME-COM" 'NUMERO DE ESPECIFICACION DE VENTA
          CODPARAM$(4) = "FECH-EMI" 'FECHA EMISION DEL LISTADO DE GRUPOS
          CODPARAM$(5) = "FECH-VEN" 'FECHA DE ENTREGA
          CODPARAM$(6) = "HOJA-NUM" 'NUMERO DE HOJA
          CODPARAM$(7) = "HOJA-TOT" 'TOTAL DE HOJAS
          CODPARAM$(8) = "R-PLANO"  'NUMERO DE REVISION
          CAPARDOC% = 8
          
          DOCPARAM(1) = RASOCLI(DESTIDOC%)
          DOCPARAM(2) = DENOOBRA$
          DOCPARAM(3) = NUMEDOC$
          DOCPARAM(4) = FECHDOC$
          DOCPARAM(5) = FECHFIN$
          DOCPARAM(6) = ""
          DOCPARAM(7) = ""
          DOCPARAM(8) = RPLA$
          '
          CAITAB1% = 0
          CODTABU1$(1) = "COD-MAT" ' CODIGO DE GRUPO
          CODTABU1$(2) = "DES-MAT" ' DENOMINACION DEL GRUPO
          CODTABU1$(3) = "MEDIDA-1" ' PROGRAMA
          CODTABU1$(4) = "MEDIDA-2" '
          CODTABU1$(5) = "MEDIDA-3" ' OBSERVACIONES
          CACOLTAB1% = 5
          '
          For U& = 1 To ULTREG&("!GRUORTRA")
            Z$ = REGLEIDO$("!GRUORTRA", U&)
            NROGRU$ = Mid$(Z$, 1, 8)
            PROGRU$ = Mid$(Z$, 9, 12)
            COMGRU$ = Mid$(Z$, 21, 12)
            DENGRU$ = Mid$(Z$, 33, 64)
            OBSGRU$ = Mid$(Z$, 97, 64)
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = NROGRU$
            TETABU1$(2, CAITAB1%) = DENGRU$
            TETABU1$(3, CAITAB1%) = PROGRU$
            TETABU1$(4, CAITAB1%) = COMGRU$
            TETABU1$(5, CAITAB1%) = OBSGRU$
          Next U&
          '
          NOCONFIRM% = 1
          Call PRINTDOC("LISTGRUP")                   ' LISTADO DE GRUPOS
          '
        End If
199     Call CIERRARCH("*.*")
        Call BAJALDISCO
        Screen.MousePointer = 1
        '
        Call CIERRARCH("*.*")
        '
        Exit Sub
        '
999     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        '
    Else
        Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
    End If
End Sub


Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG04/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
99 Command8.Enabled = True
   '
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(RASOCLI$(NCLIE%), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE%)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
    Call Command14_Click
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub COTEXT_LostFocus(Index As Integer)
  '
  If Val(COTEXT(5)) <= 0 Then
    TEXT21 = ""
  End If
  '
End Sub

Private Sub CUETEXT_Change()
     MODIFIC% = 1
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

Private Sub Form_Initialize()
  '
  SSTab1.Tab = 0
  '
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call GRATITFOR(Caption)
    '
    On Error Resume Next
    Kill LUCOM$ + "GRUORTRA.WKF"
    On Error GoTo 0
    '
    COLOBAK = RGB(230, 230, 230)
    TICAMBIS.BackColor = COLOBAK
    Text6.BackColor = COLOBAK
    Text15.BackColor = COLOBAK
    '
    Call LEETASIVA
    CONVERPRECIO = 1
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   If TRIM(Text1) <> "" Then
    Call SELECHO("VENDEDOR")
    VDEVU$ = VALACT1$("VENDEDOR")
    If TRIM$(VDEVU$) <> "" Then
         COTEXT(5).TEXT = VDEVU$
    End If
   Else
    Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   'Call GRAGENCLI
   'Call GRANOTACLI
   GRABATODO% = 0
End Sub

 Sub Command13_Click()
   Command13.Enabled = False
      LIPRESU04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPCFPRE04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    If TRIM$(Text1.TEXT) <> "" Then
        Call GRABOT
        Call SUMAHS
    Else
        Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
    End If
End Sub
Sub SUMAHS()
    XX1 = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
    Text34 = FORMATNUM(XX1, "F6.2")
End Sub
Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   'Call GRAGENCLI
   'Call GRANOTACLI
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   'Call CARDATCLIE
End Sub

'Private Sub Command20_Click()
'   Static ARGU$
'   Command20.Enabled = False
'10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
'   If ARGU$ <> "" Then
'     Screen.MousePointer = 11
'     JJ& = 0
'     Call COPYSTRU("DEUDOR1", "SELECLI")
'     '
'     Call ABRECLIEN
'     On Error Resume Next
'     With Deudor12
'       .MoveFirst
'       If Err < 1 Then
'         Do While Not .EOF
'           For KKI% = 1 To .Fields.Count
'             VATEX$ = .Fields(KKI% - 1).Value
'             If InStr(VATEX$, ARGU$) > 0 Then
'               NCLIE% = !Nume_Cli
'               RACLI$ = !Raso_Cli
'               XX$ = MKI$(NCLIE%) + RACLI$
'               JJ& = JJ& + 1
'               Call GRAREG("!SELECLI", XX$, JJ&)
'               Exit For
'             End If
'           Next KKI%
'         .MoveNext
'         Loop
'       End If
'     End With
'     Call SETULTREG("!SELECLI", JJ&)
'     '
'     Screen.MousePointer = 1
'     If JJ& < 1 Then
'         Call MENSERR(24, "No se Encontraron Coincidencias")
'         GoTo 10
'       ElseIf JJ& = 1 Then
'         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
'       Else
'         Call FRESHECHO("!SELECLI")
'         Call SELECHO("!SELECLI")
'         NCLIX$ = VALACT1$("!SELECLI")
'         If NCLIX$ <> "" Then
'           Text1 = NCLIX$
'         End If
'     End If
'   End If
'   Command20.Enabled = True
'End Sub

Private Sub Command21_Click()
   '\\\OT-05-0330-WAR-26/05/05///
   If DERACCE%("ALTANUPR", "Alta de Nuevo Presupuesto de Cámara") < 2 Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
   'FALTA PONER EL CODIGO PARA QUE TRAIGA LAS FECHAS POR DEFAULT
   '
   Call ABREPRECAM
   SQLX$ = "SELECT NRO_PRESU FROM ENCAPRESU "
   SQLX$ = SQLX$ + "WHERE NRO_PRESU LIKE 'PRE-*' "
   SQLX$ = SQLX$ + "ORDER BY NRO_PRESU DESC;"
   Set EncaTemp = PresuCam.OpenRecordset(SQLX$, dbOpenSnapshot)
   NROMA& = 0
   On Error Resume Next
   With EncaTemp
     .MoveFirst
     If Err < 1 Then
       NROMA& = Val(Mid$(!NRO_PRESU, 5))
     End If
     On Error GoTo 0
   End With
   '
   NPRESX$ = "PRE-" + TRIM$(Str$(NROMA& + 1))
   VDEF$ = Space$(64) + String$(6, 0)
   Call REPLA(VDEF$, NPRESX$, 1, 16)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 67, 2)
   '
9  OBj$ = OBJPLA$("NUEPREC", VDEF$)
   If OBj$ = "" Then GoTo 99
   VDEF$ = OBj$
   '
   NUPRE$ = TRIM$(Mid$(OBj$, 1, 16))
   DESCRIPRE$ = TRIM$(Mid$(OBj$, 17, 48))
   NUCLIPRE% = CVI(Mid$(OBj$, 65, 2))
   RASOCLIPRE$ = RASOCLI$(NUCLIOTR%)
   FECPRE% = CVI(Mid$(OBj$, 67, 2))
   FEVAPRE% = CVI(Mid$(OBj$, 69, 2))
   
   '//////////////VALIDACIONES////////////
   If TRIM$(NUPRE$) = "" Then
      Call COMUNI("Número de Presupuesto de Cámara no válido.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   EncaPresu.FindFirst "Nro_Presu = '" & NUPRE$ & "' "
   If EncaPresu.NoMatch = False Then
      Call COMUNI("Numero de Presupuesto de Cámara no Válido o Pre-Existente")
      GoTo 9
   End If
   '
   If TRIM$(DESCRIPRE$) = "" Then
      Call COMUNI("Debe Ingresar una Descripcion para el Presupuesto de Cámara")
      GoTo 9
   End If
   '
   If NUCLIPRE% <= 0 Then
      Call COMUNI("Cliente no Válido.")
      GoTo 9
   End If
   '
   If FEVAPRE% < FECPRE% Then
      Call COMUNI("La Fecha de Validez Debe ser Superior o Igual a la Fecha de Generación")
      GoTo 9
   End If
   '
   jjj = CVINT(Date)
   If FECPRE% > jjj Then
      Call COMUNI("La Fecha de Generación no puede ser Superior al dia de la Fecha.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   With EncaPresu
      .AddNew
        !NRO_PRESU = NUPRE$
        !Refe_Presu = DESCRIPRE$
        !Nclie_Presu = NUCLIPRE%
        !Femi_Presu = CVDAT(FECPRE%)
        !FVali_Presu = CVDAT(FEVAPRE%)
        !uSUCONF_Presu = USERNAME 'usuario de confección
        !UsuCont_Presu = "" 'usuario que controlo
        !UsuApro_Presu = "" 'usuario que aprobara
        !NroOcom_Presu = "" 'numero de orden de compra del cliente
        !FeOcom_Presu = Date 'fecha de pedido del cliente
        !PediCoti_Presu = "" 'Numero de Oferta
        !FePedCot_Presu = Date 'fecha de la oferta
        !Nuoritem_Presu = 0
        !CanPed_Presu = 0
        !Especi_Presu = ""
        !PreVeN_PRESU = 0
        !Mone_Presu = 2
        !CotiMone_Presu = 0
        !LENTREGA_Presu = " "
        !Ajusta_Presu = 0
        !Vended_Presu = 0
        !ComiVend_Presu = 0
        !Observa_Presu = " "
      .Update
   End With
   '
   Call GENEINDIPRECA
   Text1 = TRIM$(NUPRE$)
   '
99 Screen.MousePointer = 1
   Command21.Enabled = True
   '\\\OT-05-0330-WAR-FIN///
End Sub

Private Sub Command23_Click()
   SSTab1.Tab = 0
   Text20.TEXT = " Suspendido"
   Text20.Refresh
End Sub

Private Sub Command25_Click()
   '
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If Text1 <> "" Then
        Call PRIORVEN ' IMPRESION DE  PRESUPUESTO
        Call PRIDETPRE ' IMPRESION DE DETALLE PRESUPUESTO
   Else
        Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
   End If
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub

Private Sub Command5_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text24.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text24.TEXT)
 End If
End Sub

Private Sub Command6_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text12.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text12.TEXT)
 End If
End Sub

Private Sub Command9_Click()
   '
  ' Call GRABOT
   Call SELECHO("IPRESCAM")
   NUPRECAM$ = TRIM$(VALACT1$("IPRESCAM"))
   Text1.TEXT = TRIM$(NUPRECAM$)
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub



Private Sub Frame3_DBLCLICK()
    Call List1_DBLClick
End Sub

Private Sub Label2_dblClick()
   '
   Static CAMON%, DMONE$(), TCMONE()
   '
   If TRIM$(Text1) <> "" Then
     If MODIFIC% > 0 Then
       Call MENSERR(24, "Presupuesto Modificado.\Imposible Cambiar Moneda")
       Exit Sub
     End If
   End If
   '
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = DMX$
          TCMONE(U&) = VHX
         Next U&
     End If
   End If
   '
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   Label2 = DMONE$(MONEMI%)
   Label2.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label2.ForeColor = RGB(255, 0, 0)
   End If
   Label2.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   PRESU2$ = Text1
   Call ABREPRECAM
   EncaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
   If EncaPresu.NoMatch = False Then
     EncaPresu.Edit
     MONEMIA% = EncaPresu!Mone_Presu
     PRUNI# = EncaPresu!PreVeN_PRESU * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     EncaPresu!PreVeN_PRESU = PRUNI#
     EncaPresu!Mone_Presu = MONEMI%
     EncaPresu!CotiMone_Presu = TICAMONE(MONEMI%)
     EncaPresu.Update
   End If
   '
   DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
25 If DetaPresu.NoMatch = False Then
     DetaPresu.Edit
     PRUNI# = DetaPresu!PreUni_Det * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     DetaPresu!PreUni_Det = PRUNI#
     IMPOR# = DetaPresu!Impo_Det * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     DetaPresu!Impo_Det = IMPOR#
     DetaPresu.Update
     DetaPresu.FindNext " Nro_Presu= '" & PRESU2$ & "' "
     GoTo 25
   End If
   '
   Text1 = ""
   Text1 = PRESU2$
   Call CARGPRE
   
'   Call CARGPRE
'   '
End Sub

Private Sub Label25_Change()
   Label35 = Label25
   SUMA = Val(Label25) + Val(Label29)
   Label44 = FORMATNUM(SUMA, "F6.1")
End Sub

Private Sub Label29_Change()
   Label40 = Label29
   SUMA = Val(Label25) + Val(Label29)
   Label44 = FORMATNUM(SUMA, "F6.1")
End Sub
Private Sub LABEL40_CHANGE()
    '
    Call TOTMATCOM
    '
End Sub
Private Sub Label40_DblClick()
   '
   With COTCAM04
    .BLANFORMU
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTCAM04.Label31)
   NUGRU$ = TRIM$(COTCAM04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   'Call COTCAM04.CARGOLISMAT
   Call COTCAM04.CAMGRIS
   '
   COTCAM04.Show 1
End Sub
Private Sub LABEL42_CHANGE()
    Call TOTMATCOM
End Sub
Private Sub Label42_DBLClick()
    '
   With COTPUE04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTPUE04.Label31)
   NUGRU$ = TRIM$(COTPUE04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTPUE04.CAMGRIS
   '
   COTPUE04.Show 1
End Sub

Private Sub Label43_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label43_DblClick()
    '
        '
   With COTACC04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTACC04.Label31)
   NUGRU$ = TRIM$(COTACC04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTACC04.CAMGRIS
   '
   COTACC04.Show 1
    '
End Sub

Private Sub Label44_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label44_DblClick()
        '
   With COTEQU04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTEQU04.Label31)
   NUGRU$ = TRIM$(COTEQU04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTEQU04.CAMGRIS
   '
   COTEQU04.Show 1
    '
End Sub

Private Sub Label47_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label47_DblClick()
        '
   With COTVAR04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTVAR04.Label31)
   NUGRU$ = TRIM$(COTVAR04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTVAR04.CAMGRIS
   '
   COTVAR04.Show 1
   '
End Sub

Private Sub Label49_CHANGE()
    '
    Text23 = FORMATNUM(Val(Label49), "F11.2")
    '
End Sub

Private Sub Label5_dblClick()
    Call List1_DBLClick
End Sub

Private Sub Label52_CHANGE()
  '
  Call CATONET
  '
End Sub

Private Sub Label54_CHANGE()
  Call TOTGEN
End Sub

Private Sub LABEL62_CHANGE()
  '
  Call TOTMONT
  '
End Sub
Private Sub Label62_DBLClick()
  '
   With COTMOB04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
   End With
   '
   NUOTR$ = TRIM$(COTMOB04.Label31)
   NUGRU$ = TRIM$(COTMOB04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTMOB04.CAMGRIS
   '
   COTMOB04.Show 1
End Sub

Sub LABEL63_DBLCLICK()
  '
   '
   With COTRAN04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTRAN04.Label31)
   NUGRU$ = TRIM$(COTRAN04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTRAN04.CAMGRIS
   '
   COTRAN04.Show 1
  '
End Sub
Private Sub Label63_CHANGE()
  '
  Call TOTMONT
  '
End Sub

Private Sub Label65_CHANGE()
  '
  Call CATONET
  '
End Sub

Private Sub Label70_CHANGE()
  '
  Call CATONET
  '
End Sub

Private Sub Label71_CHANGE()
  Call TOTGEN
End Sub

Private Sub Label76_CHANGE()
  Call TOTGEN
End Sub

Private Sub Label83_CHANGE()
  '
  Call CATONET
  '
End Sub

Private Sub LiGeClaCli_Click()
    Call Command13_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub List1_DBLClick()
        'GENERO EL DAT
        If TRIM$(Text1.TEXT) <> "" Then
        NUOTRAB$ = TRIM$(Text1.TEXT)
        Call ABREORTRA
        SQLX$ = "SELECT * FROM PLACOB "
        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
        J& = 0
        With PLACOTMP
          '
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do Until .EOF = True
              Y$ = REGBLAN("!CAPLACO")
              '
              NUTRA$ = TRIM$(!Nro_Otrab)
              Call REPLA(Y$, NUTRA$, 1, 16)
              NUORITOT% = !NuOrItem
              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
              DESCRIOT$ = TRIM$(!Refe_Pcob)
              Call REPLA(Y$, DESCRIOT$, 19, 64)
              IMPOT# = !Impo_Pcob
              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
              FECOT% = CVINT(!FeVen_Pcob)
              Call REPLA(Y$, MKI(FECOT%), 91, 2)
              If IsNull(!NroFac_Pcob) = True Then
                NROFAOT$ = " "
              Else
                NROFAOT$ = TRIM$(!NroFac_Pcob)
              End If

              Call REPLA(Y$, NROFAOT$, 93, 16)
              MONEOT% = !Moneda
              Call REPLA(Y$, MKI(MONEOT%), 109, 1)
              COTIOT# = !Coti_Pcob
              Call REPLA(Y$, MKD(COTIOT#), 110, 8)
              '
              J& = J& + 1
              Call GRAREG("!CAPLACO", Y$, J&)
            .MoveNext
            Loop
          End If
          On Error GoTo 0
        End With
        '
        Call SETULTREG("!CAPLACO", J&)
        Call CIERRARCH("!CAPLACO")
        '
        'Call GENEPLACO 'GENERO EL DAT
        PRF$ = "CAPLACO"
        ATRI$ = "/NC"
        ATRI$ = ATRI$ + "/TITUPAN=Plan de Cobranza"
        ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left))
        ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + SSTab1.Top + Frame18.Top - 500))
        VTB% = VENTABU%(PRF$ + ATRI$)
        'NUEVA VALIDACION
        If VTB% < 0 Then Exit Sub
        '
        If VTB% >= 0 Then
          'GUARDO O EDITO LOS REGISTROS EN TABLA
          'ACA VA RUTINA DE GRABACION
          Call ABREORTRA
          J& = ULTREG("!CAPLACO")
          For KKIL& = 1 To J&
             XX$ = REGLEIDO$("!CAPLACO", KKIL&)
             NUOTRAB$ = TRIM$(Text1.TEXT)
             ITE% = CVI(Mid$(XX$, 17, 2))
             'VALIDACION SI ITEM VALE 0
             If ITE% <= 0 Then GoTo 7
             '
             DESCRI$ = Mid$(XX$, 19, 64)
             IMPOT# = CVD(Mid$(XX$, 83, 8))
             FECOTR% = CVI(Mid$(XX$, 91, 2))
             NUFAC = Mid$(XX$, 93, 16)
             '
             MONE1% = CVI(Mid$(XX$, 109, 1))
             COTI1# = CVD(Mid$(XX$, 110, 8))
             '
             With PlanCob
             '/// Aca FindFirst de Item+NUOTRAB$, Si existe
             ' se edita el registro, sino,
             'se crea registro nuevo////
             SQLX$ = " Nro_Otrab = '" & NUOTRAB$ & "' "
             SQLX$ = SQLX$ + " AND NuOrItem = " & ITE% & " "
             .FindFirst SQLX$
             If .NoMatch = True Then
               .AddNew
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
                 !Moneda = MONE1%
                 !Coti_Pcob = COTI1#
               .Update
              Else
               .Edit
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
                 !Moneda = MONE1%
                 !Coti_Pcob = COTI1#
               .Update
              End If
              End With
7         Next KKIL&
        End If
    
'    '/////////*******ACA GENERO NUEVAMENTE EL DAT*******///////////
'        NUOTRAB$ = TRIM$(Text1.TEXT)
'        Call ABREORTRA
'        SQLX$ = "SELECT * FROM PLACOB "
'        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
'        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
'        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
'        J& = 0
'        With PLACOTMP
'          '
'          On Error Resume Next
'          .MoveFirst
'          xxsd = .RecordCount
'          MsgBox xxsd
'          If Err < 1 Then
'            On Error GoTo 0
'            Do Until .EOF = True
'              Y$ = REGBLAN("!CAPLACO")
'              '
'              NUTRA$ = TRIM$(!Nro_Otrab)
'              Call REPLA(Y$, NUTRA$, 1, 16)
'              NUORITOT% = !NuOrItem
'              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
'              DESCRIOT$ = TRIM$(!Refe_Pcob)
'              Call REPLA(Y$, DESCRIOT$, 19, 64)
'              IMPOT# = !Impo_Pcob
'              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
'              FECOT% = CVINT(!FeVen_Pcob)
'              Call REPLA(Y$, MKI(FECOT%), 91, 2)
'              If IsNull(!NroFac_Pcob) = True Then
'                NROFAOT$ = " "
'              Else
'                NROFAOT$ = TRIM$(!NroFac_Pcob)
'              End If
'
'              Call REPLA(Y$, NROFAOT$, 93, 16)
'              MONEOT% = !Moneda
'              Call REPLA(Y$, MKI(MONEOT%), 109, 2)
'              COTIOT# = !Coti_Pcob
'              Call REPLA(Y$, MKD(COTIOT#), 111, 8)
'              '
'              J& = J& + 1
'              Call GRAREG("!CAPLACO", Y$, J&)
'            .MoveNext
'            Loop
'          End If
'          On Error GoTo 0
'        End With
'        '
'        Call SETULTREG("!CAPLACO", J&)
'        Call CIERRARCH("!CAPLACO")
        '
    '/////////******ACA GUARDO EL PLAN EDITADO******///////////////
     With PlanCob
20   .FindFirst "Nro_Otrab = '" & NUOTRAB$ & "' "
     If .NoMatch = False Then
       .Delete
       GoTo 20
     End If
     '
     For U& = 1 To ULTREG("!CAPLACO")
        Y$ = REGLEIDO$("!CAPLACO", U&)
        'NUOTR$ = Mid$(Y$, 1, 16)
        ITE1% = CVI(Mid$(Y$, 17, 2))
        REFE1$ = Mid$(Y$, 19, 64)
        IMP1# = CVD(Mid$(Y$, 83, 8))
        FECVEN% = CVI(Mid$(Y$, 91, 2))
        NFAC1$ = Mid$(Y$, 93, 16)
        MON1% = CVI(Mid$(Y$, 109, 2))
        COTMON1# = CVD(Mid$(Y$, 110, 8))
       '
       .AddNew
        !Nro_Otrab = NUOTRAB$
        !NuOrItem = ITE1%
        !Refe_Pcob = REFE1$
        !Impo_Pcob = IMP1#
        !FeVen_Pcob = CVDAT(FECVEN%)
        !NroFac_Pcob = NFAC1$
        !Moneda = MON1%
        !Coti_Pcob = COTMON1#
       .Update
     Next U&
   End With
   '//////*********/////////////
    List1.Clear
    Call CARGOPLACO
    '
    End If
    '
End Sub


Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuOrItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !Moneda, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
      
    
End Sub



Private Sub NueOT_Click()
    Call Command21_Click
End Sub



Private Sub SSTab1_Click(PREVIOUS As Integer)
    If SSTab1.Tab = 1 Then 'SI RUBROS/COTIZACION
        Call CALTOTRUB 'CALCULO LOS TOTALES DE LOS RUBROS
    End If
End Sub

Private Sub Text10_Change()
   MODIFIC% = 1
End Sub

Private Sub Text11_GotFocus()
    Text25.SetFocus
End Sub

Sub GRABOT() 'GRABAR PRESUPUESTO DE CAMARA
   '
   If YACAR% < 1 Or MODIFIC% < 1 Then
     Exit Sub
   End If
   '
   If DERACCE%("GRADATPR", "Grabar Datos de Presupuesto de Cámara") < 2 Then
       Exit Sub
   End If
   '
   'VALIDACIONES
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Referencia / Descripción Corta.")
   End If
   '
   If FECHANUM(Text10) < FECHANUM(Text9) Then
       Call COMUNI("Falta Fecha Validez de Oferta.")
   End If
   '
   If SSTab1.Tab = 0 Then
      If TRIM$(Text17) = "" Then
        Call COMUNI("Falta Especificación / Descripcion Larga.")
      End If
      '
      If Val(TRIM$(Text36)) < 1 Then
          Call COMUNI("Falta Cantidad.")
      End If
      '
      If CDbl(Val(TRIM$(Text38))) <= 0.005 Then
          Call COMUNI("Falta Importe.")
      End If
      '
      'If TRIM$(Text2) = "" Then
      '    Call COMUNI("Lugar de Entrega no Válido.")
      'End If
      If TRIM$(Text25) = "" Then
         Call COMUNI("Falta Pedido de Cotización.")
      End If
      '
      If FECHANUM(TRIM$(Text24)) < 1 Then
          Call COMUNI("Falta Fecha Pedido de Cotizacion.")
      End If
      '
'      If TRIM$(Text13) = "" Then
'          Call COMUNI("Orden de Compra no Válida.")
'      End If
'      '
'      If TRIM$(Text12) = "" Then
'          Call COMUNI("Fecha de la Orden de Compra no Válida.")
'      End If
      '
'      If TRIM$(Text21) = "" Then
'          Call COMUNI("Comisión no Válida.")
'      End If
'      '
'      If TRIM$(CUETEXT) = "" Then
'          Call COMUNI("Cuenta de Imputación no Válida.")
'      End If
      '
'      If TRIM$(COTEXT(5)) = "" Then
'          Call COMUNI("Representante no Válido.")
'      End If
      '
'      If List1.ListCount <= 0 Then
'          Call COMUNI("Plan de Cobranza no Válido.")
'          Exit Sub
'      End If
      '
   End If
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Then Exit Sub
   '
   Call ABREPRECAM
   With EncaPresu
     .FindFirst "  NRO_Presu= '" & NUOTR$ & "' "
     If .NoMatch = True Then
        Exit Sub
     End If
     On Error GoTo 0
     '
     .Edit
     !Especi_Presu = Text17
     !FVali_Presu = CVDAT(FECHANUM(Text10))
     !Refe_Presu = TRIM$(Text3)
     !NroOcom_Presu = TRIM$(Text13)
     !FeOcom_Presu = CVDAT(FECHANUM(Text24))
     !PediCoti_Presu = TRIM$(Text25)
     !FePedCot_Presu = CVDAT(FECHANUM(Text12))
     !CanPed_Presu = Val(Text36)
     !PreVeN_PRESU = Val(Text38)
     !LENTREGA_Presu = Text2
     'medidas
     !Alto_Presu = Val(Text4)
     !Ancho_Presu = Val(Text14)
     !Largo_Presu = Val(Text19)
     !Espesor_Presu = Val(Text22)
     '
     If Val(COTEXT(5)) >= 0 Then
         !Vended_Presu = Val(COTEXT(5))
     End If
     !GGMat_Presu = Val(Text44)
     !GGOba_Presu = Val(Text16)
     !DesBon_Presu = Val(Text20)
     !ComiVend_Presu = Val(TEXT21)
     '
     !Mone_Presu = MONEMI%
     !Observa_Presu = Text35
     !UsuCont_Presu = TRIM$(Text6)
     !UsuApro_Presu = Text15
     !ICONTAB = CUEINT%(CUETEXT) 'imputa. contab. de la EV
     !Observa_Presu = Text29
     '
     .Update
   End With
   '
   MODIFIC% = 0
   Call GENEINDIPRECA
   '
End Sub

Private Sub Text12_Change()
    MODIFIC% = 1
End Sub
Private Sub Text12_GotFocus()
  '
 ' Text17.SetFocus
  '
End Sub

Private Sub Text13_GOTFOCUS()
  '
  'Text17.SetFocus
  '
End Sub
Private Sub Text13_Change()
    MODIFIC% = 1
End Sub

Private Sub Text14_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) * Val(Text14) * Val(Text4)
   Label56 = FORMATNUM(SUMA, "F6.1") ' Volumen
   SUMA = Val(Text19) * Val(Text14)
   Label57 = FORMATNUM(SUMA, "F6.1") ' Techo
   Label58 = Label57 ' Piso
   SUMA = 2 * ((Val(Text4) * Val(Text14)) + (Val(Text4) * Val(Text19)))
   Label59 = FORMATNUM(SUMA, "F6.1") ' Lateral
      '\\\OT-05-0405-WAR-17/06/05///
   'AVERIGUO SI LA CAMARA TIENE PISO
   'SI TIENE PISO
   ESPESO = Val(Text22) / 1000 'PARA TENERLO EN METROS
   ANCHIN = Val(Text14) - 2 * ESPESO
   LARGIN = Val(Text19) - 2 * ESPESO
   ALTIN = Val(Text4)
   'SI TIENE ALGUN REGISTRO EN CAMARA/PISO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 2 " 'PISO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   'SI TIENE ALGUN REGISTRO EN CAMARA/TECHO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 3 " 'TECHO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   '
   VOLINTE = ANCHIN * LARGIN * ALTIN
   Label56 = FORMATNUM(VOLINTE, "F6.1")
   '\\\OT-05-0405-WAR-FIN///
End Sub

Private Sub Text14_LOSTFOCUS()
    Text14 = FORMATNUM(Val(Text14), "F6.3")
End Sub

Private Sub Text15_GOTFOCUS()
    Text29.SetFocus
End Sub

Private Sub Text16_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    MODIFIC% = 1
    MOBPORC# = Val(Text16)
    Text16 = FORMATNUM(MOBPORC#, "F5.2")
  End If
  '
  
End Sub
Sub TEXT16_CHANGE()
  '
  TOTOB# = Val(Label62)
  POROB# = Val(Text16)
  TOTGAMA# = 0
  If TOTOB# > 0 And POROB# > 0 Then
    TOGAMA# = TOTOB# * POROB# / 100
  End If
  Label69 = FORMATNUM(TOGAMA#, "F12.2")
  Call CATOGG
  '
End Sub
Private Sub Text17_Change()
    MODIFIC% = 1
End Sub

Private Sub Text18_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Text14.SetFocus
    End If
    '
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
    '
End Sub
Private Sub Text2_Change()
    MODIFIC% = 1
End Sub

Private Sub Text2_DblClick()
   Call Command11_Click
End Sub

Private Sub TEXT21_LOSTFOCUS()
    '
    TEXT21 = FORMATNUM$(Val(TRIM(TEXT21)), "F6.2")
    '
End Sub

Private Sub Text35_Change()
   MODIFIC% = 1
End Sub

Private Sub Text22_LOSTFOCUS()
  '
   Text22 = FORMATNUM(Val(Text22), "F6")
  '
End Sub

Private Sub Text23_GotFocus()
   CAJENTRA% = 21
End Sub

Private Sub Text27_GotFocus()
   CAJENTRA% = 10
End Sub

Private Sub Text38_lostfocus()
    '
    If TRIM$(Text1.TEXT) <> "" Then
      MODIFIC% = 1
      Text38 = FORMATNUM$(Val(Text38), "F12.2")
    End If
    '
End Sub

Private Sub Text4_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) * Val(Text14) * Val(Text4)
   Label56 = FORMATNUM(SUMA, "F6.1") ' Volumen
   SUMA = Val(Text19) * Val(Text14)
   Label57 = FORMATNUM(SUMA, "F6.1") ' Techo
   Label58 = Label57 ' Piso
   SUMA = 2 * ((Val(Text4) * Val(Text14)) + (Val(Text4) * Val(Text19)))
   Label59 = FORMATNUM(SUMA, "F6.1") ' Lateral
      '\\\OT-05-0405-WAR-17/06/05///
   'AVERIGUO SI LA CAMARA TIENE PISO
   'SI TIENE PISO
   ESPESO = Val(Text22) / 1000 'PARA TENERLO EN METROS
   ANCHIN = Val(Text14) - 2 * ESPESO
   LARGIN = Val(Text19) - 2 * ESPESO
   ALTIN = Val(Text4)
   'SI TIENE ALGUN REGISTRO EN CAMARA/PISO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 2 " 'PISO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   'SI TIENE ALGUN REGISTRO EN CAMARA/TECHO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 3 " 'TECHO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   '
   VOLINTE = ANCHIN * LARGIN * ALTIN
   Label56 = FORMATNUM(VOLINTE, "F6.1")
   '\\\OT-05-0405-WAR-FIN///
End Sub

Private Sub Text4_LOSTFOCUS()
    Text4 = FORMATNUM(Val(Text4), "F6.3")
End Sub
Private Sub Text44_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    MODIFIC% = 1
    PORCMAT# = Val(Text44)
    Text44 = FORMATNUM(PORCMAT#, "F5.2")
  End If
  '
End Sub
Sub TEXT44_CHANGE()
  '
  TOTMAT# = Val(Label52)
  PORMAT# = Val(Text44)
  TOGAMA# = 0
  If PORMAT# > 0 And TOTMAT > 0 Then
    TOGAMA# = TOTMAT# * PORMAT# / 100
  End If
  Label68 = FORMATNUM(TOGAMA#, "F12.2")
  '
  Call CATOGG
End Sub
Function PORDESBON$(NUOTR1$)
'DEVUELVE EL PORCENTAJE DE DESCUENTOS Y BONIFICACIONES
  Call ABREPRECAM
  With EncaPresu
    NUPRESU2$ = TRIM(NUOTR1$)
    SQLX$ = " Nro_Presu = '" & NUPRESU2$ & "' "
    .FindFirst SQLX$
    If .NoMatch = False Then
      If IsNull(!DesBon_Presu) Then GoTo 10
        PORDESBON$ = !DesBon_Presu
        Exit Function
    End If
  End With
10 PORDESBON$ = ""
'
End Function
Function PORGASGEN$(NUPRESU1$, TIPGAST$)
'DEVUELVE EL PORCENTAJE DE GASTOS GENERALES
  Call ABREPRECAM
  With EncaPresu
    '
    NUPRESU2$ = TRIM$(NUPRESU1$)
    SQLX$ = " Nro_Presu = '" & NUPRESU2$ & "' "
    .FindFirst SQLX$
    If .NoMatch = False Then
     Select Case UCase(TIPGAST$)
       Case "MATERIALES"
         If IsNull(!GGMat_Presu) Then GoTo 10
         PORGASGEN$ = !GGMat_Presu
         Exit Function
       Case "OBRA"
         If IsNull(!GGOba_Presu) Then GoTo 10
         PORGASGEN$ = !GGOba_Presu
         Exit Function
     End Select
    End If
10  PORGASGEN$ = ""
    '
  End With
End Function


Private Sub Text5_GOTFOCUS()
    Text25.SetFocus
End Sub
Private Sub COTEXT_Change(Index As Integer)
ECO$ = ""
If TRIM$(COTEXT(5).TEXT) <> "" Then
    TV11% = TIPFIELD("VENDEDOR", 1)
    LF11% = LENFIELD("VENDEDOR", 1)
    ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
    ECO$ = TRIM$(ECOARCH$("VENDEDOR", ARGU$))
End If
If ECO$ <> "" Then
   DETEXT(5).Caption = ECO$
Else
   DETEXT(5).Caption = "No Definido"
End If
'   '
'   ECO$ = ""
'   If TRIM$(COTEXT(Index).TEXT) <> "" Then
'      TV11% = TIPFIELD(ARCHEX$(Index), 1)
'      LF11% = LENFIELD(ARCHEX$(Index), 1)
'      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
'      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
'   End If
'   '
'   If ECO$ <> "" Then
'        DETEXT(Index).Caption = ECO$
'      Else
'        DETEXT(Index).Caption = "No Definido"
'   End If
'   '
   MODIFIC% = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
'   Call SELECHO(ARCHEX$(Index))
'   VDEVU$ = VALACT1$(ARCHEX$(Index))
'   If TRIM$(VDEVU$) <> "" Then
'      COTEXT(Index).TEXT = VDEVU$
'   End If
End Sub

Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text36.TEXT = ""
   Text38.TEXT = ""
   Text44.TEXT = ""
   Text16.TEXT = ""
   Text26.TEXT = ""
   Text20.TEXT = ""
   Text23.TEXT = ""
   Text27.TEXT = ""
   TICAMBIS.TEXT = ""
   Text4.TEXT = ""
   Text14.TEXT = ""
   Text19.TEXT = ""
   Text29.TEXT = ""
   Text19.TEXT = ""
   Text36.TEXT = ""
   TEXT21.TEXT = ""
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   Text3 = ""
   Text18 = ""
   Text5 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text25 = ""
   Text24 = ""
   Text11 = ""
   Text13 = ""
   Text12 = ""
   Text17 = ""
   List1.Clear
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
'   Call GRAGENCLI
'   Call GRANOTACLI
'   Call CARDATCLIE
'
  Call CIERRARCH("*.*")
'   If AGRECLI% > 0 Then
'      Call SORTCLI(1)
'      AGRECLI% = 0
'   End If
  
  Call BAJALDISCO
  Call FINAL("")
  ' Hide
'   If APLINVO$ <> "" Then
'       Call FINAL("")
'     Else
'       Hide
'   End If
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
    Call COMMAND15_CLICK
End Sub

Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub List_Click()
   Call Command13_Click
End Sub

Private Sub List3_dblClick()
   '
   LISMAT04.Label31 = Text1
   LISMAT04.Label7 = Text3
   LISMAT04.Label8 = Mid$(List3.TEXT, 24, 48)
   LISMAT04.Label2 = Left$(List3.TEXT, 8)
   LISMAT04.Label5 = Text18
   LISMAT04.Text9 = Text9
   LISMAT04.Text10 = Text10
   LISMAT04.Text11 = Text11
   '
   NUOTR$ = TRIM$(LISMAT04.Label31)
   NUGRU$ = TRIM$(LISMAT04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call LISMAT04.CARGOLISMAT
   Call LISMAT04.CAMGRIS
   '
   LISMAT04.Show 1
   '
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub NueCli_Click()
   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call COMMAND15_CLICK
End Sub

Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub Susp_Cli_Click()
   Call Command23_Click
End Sub

Private Sub Text1_Change()
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If TRIM$(Text1.TEXT) <> "" Then
      NUOTR$ = TRIM$(Text1.TEXT)
      Call CARGPRE
      Call CALTOTRUB
      Text44.TEXT = PORGASGEN(NUOTR$, "MATERIALES")
      PORCMAT# = Val(Text44.TEXT)
      Text44 = FORMATNUM(PORCMAT#, "F5.2")
      Text16.TEXT = PORGASGEN(NUOTR$, "OBRA")
      PORCOB# = Val(Text16.TEXT)
      Text16 = FORMATNUM(PORCOB#, "F5.2")
      Text20.TEXT = PORDESBON(NUOTR$)
      PORDEBO# = Val(Text20.TEXT)
      Text20 = FORMATNUM(PORDEBO#, "F5.2")
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRAGENCLI
   'Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
   MODIFIC% = 1
   SUMA = Val(Text19) * Val(Text14) * Val(Text4)
   Label56 = FORMATNUM(SUMA, "F6.1") ' Volumen
   SUMA = Val(Text19) * Val(Text14)
   Label57 = FORMATNUM(SUMA, "F6.1") ' Techo
   Label58 = Label57 ' Piso
   SUMA = 2 * ((Val(Text4) * Val(Text14)) + (Val(Text4) * Val(Text19)))
   Label59 = FORMATNUM(SUMA, "F6.1") ' Lateral
   '\\\OT-05-0405-WAR-17/06/05///
   'AVERIGUO SI LA CAMARA TIENE PISO
   'SI TIENE PISO
   ESPESO = Val(Text22) / 1000 'PARA TENERLO EN METROS
   ANCHIN = Val(Text14) - 2 * ESPESO
   LARGIN = Val(Text19) - 2 * ESPESO
   ALTIN = Val(Text4)
   'SI TIENE ALGUN REGISTRO EN CAMARA/PISO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 2 " 'PISO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   'SI TIENE ALGUN REGISTRO EN CAMARA/TECHO THEN ALTIN=ALTIN-ESPESO
   NUPRE$ = TRIM(Text1)
   SQLX$ = "SELECT * FROM DETAPRESU "
   SQLX$ = SQLX$ + " WHERE NRO_PRESU = '" & NUPRE$ & "' "
   SQLX$ = SQLX$ + " AND Rubro_Det = 1 " 'CAMARA
   SQLX$ = SQLX$ + " AND SubRub_Det = 3 " 'TECHO
   Call ABREPRECAM
   Set CAMTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With CAMTMP
    On Error Resume Next
      .MoveFirst
    If Not (Err) Then
      On Error GoTo 0
      ALTIN = ALTIN - ESPESO
    End If
   End With
   '
   VOLINTE = ANCHIN * LARGIN * ALTIN
   Label56 = FORMATNUM(VOLINTE, "F6.1")
   '\\\OT-05-0405-WAR-FIN///
End Sub
Sub text19_lostfocus()
   Text19 = FORMATNUM(Val(Text19), "F6.3")
End Sub
Private Sub text20_change()
  '
  TOTMAT# = Val(Label52)
  TOTMON# = Val(Label65)
  TOTGAGE# = Val(Label70)
  SUTOT# = TOTMAT# + TOTMON# + TOTGAGE#
  PORDEBO# = Val(Text20)
  IMPDEBO# = 0
  If SUTOT# > 0 And PORDEBO# > 0 Then
    IMPDEBO# = SUTOT# * PORDEBO# / 100
  End If
  Label83 = FORMATNUM(IMPDEBO#, "F11.2")
  Text26 = FORMATNUM(IMPDEBO#, "F11.2")
  '
End Sub
Private Sub Text20_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    MODIFIC% = 1
    BONIPORC# = Val(Text20)
    Text20 = FORMATNUM(BONIPORC#, "F5.2")
  End If
  '
End Sub
'

Private Sub Text21_Change()
   MODIFIC% = 1
End Sub
Private Sub text21_gotfocus()
  '
  If Val(COTEXT(5)) <= 0 Then
    COTEXT(5).SetFocus
  End If
  '
End Sub
Private Sub Text23_Change()
   MODIFIC% = 1
   If CAJENTRA% = 21 Then
     Text27 = FORMATNUM$(Val(Label49) - Val(Text23), "F12.2")
   End If
   Call CALCIVA
End Sub
Private Sub text23_lostfocus()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    PORIVA21# = Val(Text23)
    Text23 = FORMATNUM(PORIVA21#, "F12.2")
  End If
  '
End Sub


Private Sub Text24_Change()
   MODIFIC% = 1
End Sub


Private Sub Text25_Change()
   MODIFIC% = 1
End Sub

Private Sub Text26_Change()
   MODIFIC% = 1
   '
   IMPODEBO# = Val(Text26)
   Label83 = FORMATNUM(IMPODEBO#, "F11.2")
   '
End Sub
Private Sub TEXT26_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    BONIF1# = Val(Text26)
    Text26 = FORMATNUM(BONIF1#, "F12.2")
  End If
  '
End Sub
Private Sub Text27_Change()
   MODIFIC% = 1
   If CAJENTRA% = 10 Then
     Text23 = FORMATNUM$(Val(Label49) - Val(Text27), "F12.2")
   End If
'   IMPTAS1# = Val(Label49)
'   IMPTAS2# = Val(Text27)
'   If IMPTAS2# > IMPTAS1# Then Text27 = ""
'   IMPTMP# = IMPTAS1# - IMPTAS2#
'   If IMPTMP# = 0 Then Text23 = ""
'   If IMPTAS1# > IMPTAS2# Then
'     Text23 = FORMATNUM(IMPTMP#, "F11.2")
'   End If
   Call CALCIVA2
End Sub
Private Sub TEXT27_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    PORIVA10# = Val(Text27)
    Text27 = FORMATNUM(PORIVA10#, "F11.2")
  End If
  '
End Sub
Private Sub Text28_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT28_LOSTFOCUS()
    Text28 = FORMATNUM(Val(Text28), "F6.1")
End Sub
Private Sub Text29_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Private Sub Text3_Change()
   MODIFIC% = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT30_LOSTFOCUS()
    Text30 = FORMATNUM(Val(Text30), "F6.1")
End Sub
Private Sub Text31_Change()
   MODIFIC% = 1
End Sub

Private Sub Text32_Change()
   MODIFIC% = 1
   SUMA = Val(Text30) + Val(Text29) + Val(Text28) + Val(Text32)
   Label29 = FORMATNUM(SUMA, "F6.1")
End Sub
Sub TEXT32_LOSTFOCUS()
    Text32 = FORMATNUM(Val(Text32), "F6.1")
End Sub
Private Sub Text33_Change()
   MODIFIC% = 1
End Sub

Private Sub Text34_Change()
   MODIFIC% = 1
End Sub
Sub Text34_LOSTFOCUS()
    Text34 = FORMATNUM(Val(Text34), "F6.1")
End Sub
Private Sub LText35_DblClick()
  Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub


Private Sub Text36_Change()
   MODIFIC% = 1
End Sub
Sub TEXT36_LOSTFOCUS()
    Text36 = FORMATNUM(Val(Text36), "F6.1")
End Sub

Private Sub Text6_GOTFOCUS()
    Text29.SetFocus
End Sub

Private Sub Text7_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text8_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text9_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub VenCli_Click()
With OPMASCL04
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Sub CARGPRE() 'CARGA LA FICHA DE UN PRESUPUESTO DE Cámara DETERMINADO
   '
   Static CAMON%, DMONE$(), TCMONE()
   '
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = DMX$
          TCMONE(U&) = VHX
         Next U&
     End If
   End If
   '
   NUPRECA$ = TRIM$(Text1)
   If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREPRECAM
   SQLX$ = "SELECT * FROM ENCAPRESU"
   Set OTTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With OTTMP
     .FindFirst "  NRO_pRESU= '" & NUPRECA$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        'Call COMUNI("No Existe la Orden de Trabajo")
        Exit Sub
     End If
     On Error GoTo 0
     '
     Text3 = !Refe_Presu
     '
     NCLIE% = !Nclie_Presu
     Label22 = NCLIE%
     If NCLIE% > 0 Then
       Text18 = RASOCLI(NCLIE%)
       Text5 = DOMICLI$(NCLIE%)
       Text7 = CPOSCLI$(NCLIE%)
       Text8 = LOCACLI$(NCLIE%)
     End If
     '
     Text9 = FECHATEX(CVINT(!Femi_Presu))
     Text10 = FECHATEX(CVINT(!FVali_Presu))
     Text25 = !NroOcom_Presu
     Text24 = FECHATEX(CVINT(!FeOcom_Presu))
     Text11 = !uSUCONF_Presu
     Text13 = !PediCoti_Presu
     Text12 = FECHATEX(CVINT(!FePedCot_Presu))
     Text17 = !Especi_Presu
     Text36 = !CanPed_Presu
     ALTPRE = !Alto_Presu
     ANCPRE = !Ancho_Presu
     LARPRE = !Largo_Presu
     ESPRE = !Espesor_Presu
     Text4 = FORMATNUM$(ALTPRE, "F6.3")
     Text14 = FORMATNUM$(ANCPRE, "F6.3")
     Text19 = FORMATNUM$(LARPRE, "F6.3")
     Text22 = FORMATNUM$(ESPRE, "F6")
     If IsNull(!UsuCont_Presu) Then
        Text6 = ""
     Else
        Text6 = TRIM(!UsuCont_Presu)
     End If
     If IsNull(!UsuApro_Presu) Then
        Text15 = ""
     Else
        Text15 = TRIM(!UsuApro_Presu)
     End If
     '
     Text2 = !LENTREGA_Presu
     COTEXT(5) = !Vended_Presu
     COMIVE = Val(!ComiVend_Presu)
     TEXT21 = FORMATNUM(COMIVE, "F6.2")
     If IsNull(!ICONTAB) = True Then
        CUETEXT = ""
      Else
        CUETEXT = CODCUE(!ICONTAB) 'CARGO LA IMPUTACION CONTABLE
     End If
     '
     MONEMI% = !Mone_Presu
     If MONEMI% < 1 Then MONEMI% = 1
     Label2 = DMONE$(MONEMI%)
     Label2.ForeColor = RGB(0, 0, 255)
     If MONEMI% > 1 Then
        Label2.ForeColor = RGB(255, 0, 0)
     End If
     Label2.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
     '
     Label81 = Label2
     Label81.ForeColor = Label2.ForeColor
     Label81.FontUnderline = Label81.FontUnderline
     PVBAS = !PreVeN_PRESU
     Text38 = FORMATNUM$(PVBAS, "F12.2")
     If IsNull(!Observa_Presu) = True Then
       Text29 = ""
     Else
       Text29 = !Observa_Presu
     End If
     '
   End With
   '
   Call CARGOPLACO
   Call SUMAHS
   MODIFIC% = 0
   YACAR% = 1
   '
End Sub
'
Sub GENEINDIPRECA() 'GENERA EL INDICE DE ORDENES DE TRABAJO
   'If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABREPRECAM
      Set ORDETRATECO = PresuCam.OpenRecordset("SELECT NRO_Presu, Nclie_Presu, REFE_Presu,FEMI_Presu FROM encapresu  ORDER BY REFE_Presu ASC;", dbOpenDynaset, dbReadOnly)
      '
      JJ& = 0
      On Error Resume Next
      ORDETRATECO.MoveLast
      If Err Then Exit Sub
      '
      FIN& = CAREC%
      REBLB$ = REGBLAN$("IPRESCAM")
      With ORDETRATECO
        .MoveFirst
        '
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          NCLIE% = !Nclie_Presu
          EL2$ = REBLB$
          Call REPLA(EL2$, !NRO_PRESU, 1, 16)
            FEX = CVINT(!Femi_Presu)
            REFEINDI$ = FECHATEX$(FEX) + " - " + TRIM$(FANTCLI$(NCLIE%))
            REFEINDI$ = REFEINDI$ + " - " + !Refe_Presu
          Call REPLA(EL2$, REFEINDI$, 17, 64)
          Call GRAREG("IPRESCAM", EL2$, JJ&)
          '
        .MoveNext
        Loop
      End With
      '
90    ORDETRATECO.Close
      Call SETULTREG("IPRESCAM", JJ&)
      Call FILESORT("IPRESCAM", "", 1, 1, "ASC")
      Call FRESHECHO("IPRESCAM")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
End Sub
'\\\OT-05-0394-WAR-10/06/05///
Sub PRIORVEN()
    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
    'Call SETULTREG("!OBSERVOF", 0)
    '
    ' IMPRESION DEL PRESUPUESTO DE CAMARA
    '
    FORIPRE$ = CONTROL$("", "ESPVENTA")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Text1)
      TIPODOC$ = "Presupuesto de Cámara"
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "IMP-TOTA" 'IMPORTE TOTAL
      CODPARAM$(2) = "VENDEDOR" 'REPRESENTANTE
      CODPARAM$(3) = "FECH-LAN" 'FECHA DE LA OFERTA
      CODPARAM$(4) = "FECH-REC" 'FECHA DEL PEDIDO
      CODPARAM$(5) = "DOMI-TRA" 'LUGAR DE ENTREGA
      CODPARAM$(6) = "NRO-OCOM" 'NUMERO DE PEDIDO
      CODPARAM$(7) = "NRO-AENT" 'NUMERO DE OFERTA
      CODPARAM$(8) = "FECH-VEN" 'FECHA DE ENTREGA
      CODPARAM$(9) = "MEDIDA-1" 'MEDIDA1
      CODPARAM$(10) = "MEDIDA-2" 'MEDIDA2
      CODPARAM$(11) = "MEDIDA-3" 'MEDIDA3
      CODPARAM$(12) = "REF-MAT1" '
      CODPARAM$(13) = "REF-MAT2" '
      CODPARAM$(14) = "REF-MAT3" '
      CODPARAM$(15) = "R-PLANO" 'ESPESOR
      CODPARAM$(16) = "F-PLANO" '
      CODPARAM$(17) = "N-HPROF" 'VOLUMEN
      CODPARAM$(18) = "N-HINSP" 'TECHO
      CODPARAM$(19) = "N-ORIGIN" 'PISO
      CODPARAM$(20) = "N-OCOMPR" 'LATERAL
      CODPARAM$(21) = "FECH-EMI" 'FECHA EMISION
      CODPARAM$(22) = "NUME-COM" 'NUMERO PRESUPUESTO
      CODPARAM$(23) = "COT-PLCO" 'TOT.CAMARA
      CODPARAM$(24) = "REF-PLCO" ' TOT. PUERTAS
      CODPARAM$(25) = "MED-PLCO" 'TOT. ACCESORIOS
      CODPARAM$(26) = "FRE-PLCO" 'TOT. VARIOS
      CODPARAM$(27) = "REA-PLCO" 'TOT. RUBROS
      CODPARAM$(28) = "LAR-CUTE" 'TOT.MONTAJE
      CODPARAM$(29) = "ACH-CUTE" 'TOT.TRANSPORTE
      CODPARAM$(30) = "LAR-MATE" 'TOT.MONTAJE/INSTALACION
      CODPARAM$(31) = "DIB-COMP"  ' PORCENAJE MATERIALES
      CODPARAM$(32) = "ACH-MATE" 'TOT.MATERIALES GG
      CODPARAM$(33) = "MAQ-COMP" 'PORCENTAJE MDO
      CODPARAM$(34) = "LAR-CUCO" 'TOT.MDO GG
      CODPARAM$(35) = "ACH-CUCO" 'TOT. GASTOS GRALES
      CODPARAM$(36) = "LAR-COMP" 'PORCENTAJE DTOS.
      CODPARAM$(37) = "IMP-DESC" 'IMP. DTOS.
      CODPARAM$(38) = "IMPNEDES" 'TOT. DTOS.
      CODPARAM$(39) = "IMP-NETO" 'TOT.IMPORTE NETO
      CODPARAM$(40) = "LAR-MACO" 'TOT. IVA PORCENTAJE 1
      CODPARAM$(41) = "ACH-MACO" 'TOT. IVA PORCENTAJE 2
      CODPARAM$(42) = "IMP-IVA" 'IMPORTE IVA 1
      CODPARAM$(43) = "IMP-IVA1" 'IMPORTE IVA 2
      CODPARAM$(44) = "PIE-COMP" 'TOT. PERCEPCION I.B.
      CODPARAM$(45) = "PERC-BRU" 'IMPORTE PERCEP. I.B.
      CODPARAM$(46) = "CANTIDAD" 'CANTIDAD
      CODPARAM(47) = "CAN-PLCO" 'TOTAL EQUIPO
      CAPARDOC% = 47
      
      DOCPARAM(1) = Val(Label77)
      DOCPARAM(2) = DETEXT(5)
      DOCPARAM(3) = Text24
      DOCPARAM(4) = Text12
      DOCPARAM(5) = Text2
      DOCPARAM(6) = Text13
      DOCPARAM(7) = Text25
      DOCPARAM(8) = Text10
      DOCPARAM(9) = Text19
      DOCPARAM(10) = Text14
      DOCPARAM(11) = Text4
      DOCPARAM(12) = Label89
      DOCPARAM(13) = Label88
      DOCPARAM(14) = Label87
      DOCPARAM(15) = Text22
      DOCPARAM(16) = Label86
      DOCPARAM(17) = Label56
      DOCPARAM(18) = Label57
      DOCPARAM(19) = Label58
      DOCPARAM(20) = Label59
      DOCPARAM(21) = Text9
      DOCPARAM(22) = Text1
      DOCPARAM(23) = Label40
      DOCPARAM(24) = Label42
      DOCPARAM(25) = Label43
      DOCPARAM(26) = Label47
      DOCPARAM(27) = Label52
      DOCPARAM(28) = Label62
      DOCPARAM(29) = Label63
      DOCPARAM(30) = Label65
      DOCPARAM(31) = Text44
      DOCPARAM(32) = Label68
      DOCPARAM(33) = Text16
      DOCPARAM(34) = Label69
      DOCPARAM(35) = Label70
      DOCPARAM(36) = Text20
      DOCPARAM(37) = Text26
      DOCPARAM(38) = Label83
      DOCPARAM(39) = Label49
      DOCPARAM(40) = Text23
      DOCPARAM(41) = Text27
      DOCPARAM(42) = Label54
      DOCPARAM(43) = Label71
      DOCPARAM(44) = Label75
      DOCPARAM(45) = Label76
      DOCPARAM(46) = Text36
      DOCPARAM(47) = Label44
      '
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 1
      REPLAORI% = 0
      RECOPROP% = 0
      REVICODX% = 0
      '
      CODTABU1$(1) = "DES-MAT"
      CACOLTAB1% = 1
      '
ANTEX% = 52
      Call FRATEXTO(Text17, ANTEX%)
      For UI% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = RETEX$(UI%)
      Next UI%
      '
      CODTABU2$(1) = "TIPO-VAL"
      CODTABU2$(2) = "BANC-CHE"
      CODTABU2$(3) = "IMPO-CHE"
      CODTABU2$(4) = "FECH-CHE"
      CACOLTAB2% = 4
      '
      CAITAB2% = 0
      '
     'DESDE ACA NUEVO PARA IMPRIMIR PLAN DE COBRANZA
      NUOTR$ = TRIM$(Text1)
      Call ABREORTRA
      SQLX$ = "SELECT * FROM PLACOB "
      SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
      SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
      Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
      With OTTMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         GoTo 198
       End If
       On Error GoTo 0
       '
       Do Until .EOF = True
         NOITOT1$ = CStr(!NuOrItem)
         DESCRIOT1$ = CStr(!Refe_Pcob)
         IMPOT1$ = FORMATNUM(!Impo_Pcob, "F10.2")
         FEOT1$ = FECHATEX(CVINT(!FeVen_Pcob))
         CAITAB2% = CAITAB2% + 1
         TETABU2$(1, CAITAB2%) = NOITOT1$
         TETABU2$(2, CAITAB2%) = DESCRIOT1$
         TETABU2$(3, CAITAB2%) = IMPOT1$
         TETABU2$(4, CAITAB2%) = FEOT1$
        .MoveNext
       Loop
      End With
'      'HASTA ACA NUEVO PARA IMPRIMIR PLAN DE COBRANZA
'      'ACA IMPRIMO OBSERVACION
198  ' Call FRATEXTO(Text29, 70)
'      For KKU& = 1 To CARETEX%
'       Call REGAPP("!OBSERVOF", RETEX$(KKU&))
'      Next KKU&
'      Call CIERRARCH("!OBSERVOF")
'      '
      Call PRINTDOC("ESPVENTA") 'PRESUPUESTO DE CAMARA
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
    Call CIERRARCH("*.*")
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'\\\OT-05-0394-WAR-FIN///
Sub CALTOTRUB() 'CALCULO LOS TOTALES DE LOS RUBROS
  '
  NUPRECA$ = TRIM$(Text1)
  If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
  '
  Call CATOCAM 'CALCULO DE TOTAL DE CAMARAS
  Call CATOPUE 'CALCULO DE TOTAL DE PUERTAS
  Call CATOACC 'CALCULO DE TOTAL DE ACCESORIOS
  Call CATOEQU 'CALCULO DE TOTAL DE EQUIPOS
  Call CATOVAR 'CALCULO DE TOTAL DE VARIOS
  Call CATOMOB 'CALCULO DE TOTAL DE MOB
  Call CATOTRAN 'CALCULO DE TOTAL DE TRANSPORTE
  '
End Sub
Sub CATOPUE() 'CALCULO DE TOTAL DE PUERTAS
   '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 2 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label42 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub
Sub CATOACC()
  '
    NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 3 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label43 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub
Sub CATOEQU()
  '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 4 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label44 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub
Sub CATOVAR()
  '
    NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 5 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label47 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub
Sub CATOCAM() 'CALCULO DE TOTAL DE CAMARAS
  '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 1 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label40 = FORMATNUM(TOT1#, "F11.2")
End Sub
Sub CATOTRAN() 'CALCULO DE TOTAL DE TRANSPORTE
 '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 12 "
  SQLX$ = SQLX$ + "ORDER BY DESCRI_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label63 = FORMATNUM(TOT1#, "F11.2")
End Sub
Sub CATOMOB() 'CALCULO DEL TOTAL DE MANO DE OBRA
  '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 11 "
  SQLX$ = SQLX$ + "ORDER BY DESCRI_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label62 = FORMATNUM(TOT1#, "F11.2")
End Sub
Sub TOTMONT() 'CALCULO DEL TOTAL DE MONTAJE E INSTALACION
  '
  TOTMOB1# = Val(Label62)
  TOTRAN1# = Val(Label63)
  TOTMONT1# = TOTMOB1# + TOTRAN1#
  Label65 = FORMATNUM(TOTMONT1#, "F11.2")
  '
End Sub
Sub TOTMATCOM() 'CALCULO DEL TOTAL DE MATERIALES Y COMPONENTES
  '
  TOTCAM1# = Val(Label40)
  TOTPUE1# = Val(Label42)
  TOTACC1# = Val(Label43)
  TOTEQU1# = Val(Label44)
  TOTVAR1# = Val(Label47)
  TOMACO1# = TOTCAM1# + TOTPUE1# + TOTACC1# + TOTEQU1# + TOTVAR1#
  Label52 = FORMATNUM(TOMACO1#, "F11.2")
  '
End Sub
Sub CATOGG() 'CALCULO DEL TOTAL DE GASTOS GENERALES
  '
  TOTMAT# = Val(Label68)
  TOTOB# = Val(Label69)
  TOTGG# = TOTMAT# + TOTOB#
  Label70 = FORMATNUM(TOTGG#, "F11.2")
  '
End Sub
Sub CATONET() 'CALCULO DEL TOTAL NETO
  '
  TOTMAT# = Val(Label52)
  TOTMON# = Val(Label65)
  TOTGAGE# = Val(Label70)
  TOTDEBO# = Val(Label83)
  TOTNET# = TOTMAT# + TOTMON# + TOTGAGE# - TOTDEBO#
  Label49 = FORMATNUM(TOTNET#, "F11.2")
  '
End Sub
Sub CALCIVA()
  '
10 IMPTAS1# = Val(Text23)
   PORTAS1# = Val(PORIVA(0))
   If IMPTAS1# > Val(Label49) Then Text23 = Val(Label49): GoTo 10
   IMPIVA1# = IMPTAS1# * PORTAS1# / 100
   Label54 = FORMATNUM(IMPIVA1#, "F11.2")
  '
End Sub
Sub CALCIVA2()
  '
  IMPTAS2# = Val(Text27)
  PORTAS2# = Val(PORIVA(1))
  IMPIVA2# = IMPTAS2# * PORTAS2# / 100
  Label71 = FORMATNUM(IMPIVA2#, "F11.2")
  '
End Sub
Sub TOTGEN()
  '
  IMPNET# = Val(Label49)
  IMPIVA1# = Val(Label54)
  IMPIVA2# = Val(Label71)
  IMPPER# = Val(Label76)
  TOTGEN1# = IMPNET# + IMPIVA1# + IMPIVA2# + IMPPER#
  Label77 = FORMATNUM(TOTGEN1#, "F11.2")
  '
End Sub
Sub CARGOPLACO() 'CARGO PLAN DE COBRA. DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM PLACOB "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With OTTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do Until .EOF = True
         NOITOT1$ = CStr(!NuOrItem)
         DESCRIOT1$ = CStr(!Refe_Pcob)
         IMPOT1$ = FORMATNUM(!Impo_Pcob, "F10.2")
         FEOT1$ = FECHATEX(CVINT(!FeVen_Pcob))
         Y$ = Space$(64)
         Call REPLA(Y$, AJUSTD$(NOITOT1$, 3), 1, 3)
         Call REPLA(Y$, AJUSTI$(DESCRIOT1$, 15), 6, 15)
         Call REPLA(Y$, IMPOT1$, 21, 10)
         Call REPLA(Y$, FEOT1$, 33, 8)
         List1.AddItem Y$
        .MoveNext
      Loop
      
      '
   End With
End Sub

Sub CARGOLISGRU() 'CARGO LISTA DE GRUPOS DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM GRUPOTRA "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   REBLA$ = REGBLAN$("GRUORTRA")
   JJ& = 0
   List3.Clear
   '
   With OTTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do Until .EOF = True
         NROGRU$ = !Nro_Grup
         If IsNull(!PROG_GRUP) Then
            PROGRU$ = ""
         Else
            PROGRU$ = !PROG_GRUP
         End If
         If IsNull(!COMP_GRUP) Then
            COMGRU$ = ""
         Else
            COMGRU$ = !COMP_GRUP
         End If
         If IsNull(!DENO_GRUP) Then
            DENGRU$ = ""
         Else
            DENGRU$ = !DENO_GRUP
         End If
         If IsNull(!OBSERVA_GRUP) Then
            OBSGRU$ = ""
         Else
            OBSGRU$ = !OBSERVA_GRUP
         End If
         '
         Z$ = REBLA$
         Call REPLA(Z$, NROGRU$, 1, 8)
         Call REPLA(Z$, PROGRU$, 9, 12)
         Call REPLA(Z$, COMGRU$, 21, 12)
         Call REPLA(Z$, DENGRU$, 33, 64)
         Call REPLA(Z$, OBSGRU$, 97, 64)
         JJ& = JJ& + 1
         Call GRAREG("!GRUORTRA", Z$, JJ&)
         '
         Y$ = Space$(128)
         Call REPLA(Y$, NROGRU$, 1, 8)
         Call REPLA(Y$, PROGRU$, 10, 6)
         Call REPLA(Y$, COMGRU$, 17, 6)
         Call REPLA(Y$, DENGRU$, 24, 48)
         Call REPLA(Y$, OBSGRU$, 73, 24)
         List3.AddItem Y$
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Call SETULTREG("!GRUORTRA", JJ&)
   '
End Sub
Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   'PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Private Sub TICAMBIS_GOTFOCUS()
  '
  Command15.SetFocus
  '
End Sub
'\\\OT-05-0394-WAR-10/06/05///
Sub PRIDETPRE() ' IMPRESION DE DETALLE PRESUPUETOS X RUBRO
    '
    FORIPRE$ = CONTROL$("", "DETAPRE")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Label31) + " / " + TRIM$(Label2)
      TIPODOC$ = "Lista de Materiales "
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "NUME-COM" 'NUMERO DE PRESU
      CODPARAM$(2) = "FECH-EMI" 'FECHA DE EMISION
      CODPARAM$(3) = "FECH-VEN" 'FECHA DE VALIDEZ
      CODPARAM$(4) = "FECH-LAN" 'FECHA DE PED. COTIZ.
      CODPARAM$(5) = "NRO-AENT" 'PEDIDO COTIZACION
      CODPARAM$(6) = "FECH-REC" 'FECHA DE O/C
      CODPARAM$(7) = "NRO-OCOM" 'O/C
      
      CAPARDOC% = 7
      
      DOCPARAM(1) = TRIM(Text1)
      DOCPARAM(2) = Text9
      DOCPARAM(3) = Text10
      DOCPARAM(4) = Text24
      DOCPARAM(5) = Text25
      DOCPARAM(6) = Text12
      DOCPARAM(7) = Text13
      '
      CAITAB1% = 0
      CAITCRU% = 0
      '
 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
      CODTABU1$(1) = "REF-MAT1" 'RUBRO
      CODTABU1$(2) = "REF-MAT2" 'SUB-RUBRO
      CODTABU1$(3) = "COD-MAT" 'CODIGO
      CODTABU1$(4) = "DES-MAT" 'DESCRIPCION
      CODTABU1$(5) = "REF-MAT3" 'METROS
      CODTABU1$(6) = "CANTIDAD" 'CANTIDAD
      CACOLTAB1% = 6
      '
      'ACA PONGO EL DETALLE
      NROPRE$ = TRIM(Text1)
      Call ABREPRECAM
      SQLX$ = "SELECT * FROM DETAPRESU"
      SQLX$ = SQLX$ + " WHERE Nro_Presu = '" & NROPRE$ & "'"
      SQLX$ = SQLX$ + " AND  Rubro_Det >=  1 "
      SQLX$ = SQLX$ + " AND  Rubro_Det <=  5 "
      SQLX$ = SQLX$ + " ORDER BY Rubro_Det, SubRub_Det"
      Set DETMP = PresuCam.OpenRecordset(SQLX$, 4)
      With DETMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          GoTo 36
        End If
        On Error GoTo 0
        J& = 0
        Do While .EOF = False
        '
        Select Case !rubro_det
          Case 1
            RUB$ = "CAMARA"
          Case 2
            RUB$ = "PUERTAS"
          Case 3
            RUB$ = "ACCESORIOS"
          Case 4
            RUB$ = "EQUIPOS"
          Case 5
            RUB$ = "VARIOS"
        End Select
        Select Case !SUBRUB_DET
          Case 1
            SUBRUB$ = "LATERALES"
          Case 2
            SUBRUB$ = "PISO"
          Case 3
            SUBRUB$ = "TECHO"
          Case 0
            SUBRUB$ = ""
        End Select
        '
        If J& = 0 Then
          RUB1$ = RUB$
          SUBRUB1$ = SUBRUB$
        End If
        If J& > 0 Then
         If RUB1$ = RUB$ Then
           RUB$ = ""
         Else
           RUB1$ = RUB$
         End If
        End If
        If J& > 0 Then
         If SUBRUB1$ = SUBRUB$ Then
           SUBRUB$ = ""
         Else
           SUBRUB1$ = SUBRUB$
         End If
        End If
        '
        CI% = !CODINTE_DET
        CE1$ = CODEXT$(CI%)
        DA$ = DESCRIT0$(CI%)
        '
        J& = J& + 1
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = RUB$
        TETABU1$(2, CAITAB1%) = SUBRUB$
        TETABU1$(3, CAITAB1%) = CE1$
        TETABU1$(4, CAITAB1%) = DA$
        TETABU1$(5, CAITAB1%) = !MED1_DET
        TETABU1$(6, CAITAB1%) = !Cantid_Det
        '
        .MoveNext
      Loop
      '
      End With
      'ACA IMPRIMO OBSERVACION
35   Call FRATEXTO(Text29, 70)
      For KKU& = 1 To CARETEX%
       Call REGAPP("!OBSERVOF", RETEX$(KKU&))
      Next KKU&
      Call CIERRARCH("!OBSERVOF")
      '
36    NOCONFIRM% = 1
      '
      Call PRINTDOC("DETAPRE") ' DETALLE DE PRESUPUESTO
      '
      'LIMPIA LA LISTA
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
End Sub
'\\\OT-05-0394-WAR-FIN///
