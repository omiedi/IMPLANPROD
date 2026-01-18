VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form AMASTO00 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos - A-B-M-L Maestro de Items de Stock"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "AMASTO00.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   10740
      TabIndex        =   114
      Top             =   4880
      Width           =   1040
      Begin VB.CommandButton Command13 
         Caption         =   "Class"
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
         Height          =   690
         Left            =   125
         Picture         =   "AMASTO00.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Clasificación y Búsqueda por Atributos Técnicos"
         Top             =   2385
         Width           =   765
      End
      Begin VB.CommandButton Command39 
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
         Left            =   125
         Picture         =   "AMASTO00.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   1695
         Width           =   765
      End
      Begin VB.CommandButton Command38 
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
         Left            =   125
         Picture         =   "AMASTO00.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Emisión de Listados"
         Top             =   1005
         Width           =   765
      End
      Begin VB.CommandButton Command37 
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
         Left            =   125
         Picture         =   "AMASTO00.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1110
      Left            =   7140
      ScaleHeight     =   1050
      ScaleWidth      =   60
      TabIndex        =   112
      Top             =   315
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   113
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   10740
      TabIndex        =   97
      Top             =   1575
      Width           =   1040
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
         Left            =   125
         Picture         =   "AMASTO00.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   1005
         Width           =   765
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
         Left            =   125
         Picture         =   "AMASTO00.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   315
         Width           =   765
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Subst"
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
         Left            =   125
         Picture         =   "AMASTO00.frx":165E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   1695
         Width           =   765
      End
      Begin VB.CommandButton Command24 
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
         Left            =   125
         Picture         =   "AMASTO00.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   2385
         Width           =   765
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      MaxLength       =   15
      TabIndex        =   2
      Top             =   315
      Width           =   2040
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
      Height          =   330
      Left            =   3600
      Picture         =   "AMASTO00.frx":1AB2
      TabIndex        =   3
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7455
      TabIndex        =   1
      Top             =   210
      Width           =   4320
      Begin VB.CommandButton Command25 
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
         Height          =   750
         Left            =   105
         Picture         =   "AMASTO00.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   315
         Width           =   750
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
         Height          =   750
         Index           =   4
         Left            =   3400
         Picture         =   "AMASTO00.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command4 
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
         Height          =   750
         Left            =   2650
         Picture         =   "AMASTO00.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restore"
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
         Left            =   1760
         Picture         =   "AMASTO00.frx":251A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command1 
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
         Height          =   750
         Left            =   1010
         Picture         =   "AMASTO00.frx":2824
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6420
      Left            =   150
      TabIndex        =   0
      Top             =   1680
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   11324
      _Version        =   327681
      Tab             =   2
      TabHeight       =   520
      BackColor       =   12648447
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "AMASTO00.frx":2B2E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MARCOFICHA(0)"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Label17"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Fórmula / Estructura"
      TabPicture(1)   =   "AMASTO00.frx":2B4A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Secuencia / Hoja de Ruta"
      TabPicture(2)   =   "AMASTO00.frx":2B66
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "3"
         Height          =   5895
         Index           =   3
         Left            =   105
         TabIndex        =   137
         Top             =   420
         Width           =   10095
         Begin VB.Frame Frame4 
            Caption         =   "Detalle de la Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5580
            Left            =   4830
            TabIndex        =   140
            Top             =   210
            Width           =   5160
            Begin VB.CommandButton Command28 
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
               Height          =   330
               Index           =   0
               Left            =   2100
               TabIndex        =   65
               Top             =   1785
               Width           =   175
            End
            Begin VB.TextBox Text5 
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
               Left            =   1260
               TabIndex        =   63
               Top             =   780
               Width           =   3615
            End
            Begin VB.TextBox Text7 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3885
               TabIndex        =   62
               Text            =   " "
               Top             =   420
               Width           =   990
            End
            Begin VB.TextBox Text9 
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
               Index           =   0
               Left            =   1260
               MaxLength       =   6
               TabIndex        =   64
               Text            =   " "
               Top             =   1785
               Width           =   885
            End
            Begin VB.TextBox Text11 
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
               Left            =   3045
               TabIndex        =   66
               Top             =   1785
               Width           =   1830
            End
            Begin VB.TextBox Text15 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1260
               TabIndex        =   67
               Top             =   2835
               Width           =   465
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1785
               TabIndex        =   68
               Top             =   2835
               Width           =   360
            End
            Begin VB.TextBox Text26 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   2205
               TabIndex        =   69
               Top             =   2835
               Width           =   360
            End
            Begin VB.TextBox Text15 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1260
               TabIndex        =   70
               Top             =   3150
               Width           =   465
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1785
               TabIndex        =   71
               Top             =   3150
               Width           =   360
            End
            Begin VB.TextBox Text26 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   2205
               TabIndex        =   72
               Top             =   3150
               Width           =   360
            End
            Begin VB.TextBox Text15 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1260
               TabIndex        =   73
               Top             =   3465
               Width           =   465
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1785
               TabIndex        =   74
               Top             =   3465
               Width           =   360
            End
            Begin VB.TextBox Text26 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   2205
               TabIndex        =   75
               Top             =   3465
               Width           =   360
            End
            Begin VB.TextBox Text15 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1260
               TabIndex        =   76
               Top             =   3780
               Width           =   465
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1785
               TabIndex        =   77
               Top             =   3780
               Width           =   360
            End
            Begin VB.TextBox Text26 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   2205
               TabIndex        =   78
               Top             =   3780
               Width           =   360
            End
            Begin VB.CommandButton Command28 
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
               Index           =   1
               Left            =   4800
               TabIndex        =   80
               Top             =   2625
               Width           =   175
            End
            Begin VB.TextBox Text9 
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
               Index           =   1
               Left            =   4030
               MaxLength       =   6
               TabIndex        =   79
               Top             =   2625
               Width           =   780
            End
            Begin VB.CommandButton Command28 
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
               Index           =   2
               Left            =   4800
               TabIndex        =   82
               Top             =   2940
               Width           =   175
            End
            Begin VB.TextBox Text9 
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
               Index           =   2
               Left            =   4035
               MaxLength       =   6
               TabIndex        =   81
               Top             =   2940
               Width           =   780
            End
            Begin VB.CommandButton Command28 
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
               Index           =   3
               Left            =   4800
               TabIndex        =   84
               Top             =   3255
               Width           =   175
            End
            Begin VB.TextBox Text9 
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
               Index           =   3
               Left            =   4035
               MaxLength       =   6
               TabIndex        =   83
               Top             =   3255
               Width           =   780
            End
            Begin VB.TextBox Text29 
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
               Left            =   4305
               TabIndex        =   85
               Text            =   " "
               Top             =   3780
               Width           =   570
            End
            Begin VB.TextBox Text30 
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
               Left            =   1260
               TabIndex        =   86
               Text            =   " "
               Top             =   4410
               Width           =   570
            End
            Begin VB.TextBox Text33 
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
               Left            =   2415
               TabIndex        =   87
               Text            =   " "
               Top             =   4410
               Width           =   360
            End
            Begin VB.CommandButton Command29 
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
               Left            =   2775
               TabIndex        =   88
               Top             =   4410
               Width           =   175
            End
            Begin VB.TextBox Text15 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1260
               TabIndex        =   89
               Top             =   5040
               Width           =   465
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1785
               TabIndex        =   90
               Top             =   5040
               Width           =   360
            End
            Begin VB.TextBox Text26 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   2205
               TabIndex        =   91
               Top             =   5040
               Width           =   360
            End
            Begin VB.TextBox Text34 
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
               Left            =   4200
               TabIndex        =   92
               Text            =   " "
               Top             =   5040
               Width           =   675
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Operación:"
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
               Left            =   -735
               TabIndex        =   174
               Top             =   525
               Width           =   1905
            End
            Begin VB.Label Label43 
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
               Height          =   285
               Left            =   1260
               TabIndex        =   173
               Top             =   2160
               Width           =   3615
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Standard:"
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
               Left            =   1890
               TabIndex        =   172
               Top             =   525
               Width           =   1905
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo Principal:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   171
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "HPF:"
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
               Left            =   2205
               TabIndex        =   170
               Top             =   1890
               Width           =   750
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Preparación:"
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
               Left            =   -735
               TabIndex        =   169
               Top             =   2940
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   1680
               TabIndex        =   168
               Top             =   2835
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   2100
               TabIndex        =   167
               Top             =   2835
               Width           =   120
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "T.Std.Fijo:"
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
               Index           =   1
               Left            =   -735
               TabIndex        =   166
               Top             =   3255
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   1680
               TabIndex        =   165
               Top             =   3150
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   2100
               TabIndex        =   164
               Top             =   3150
               Width           =   120
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "T.Var.Lote:"
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
               Index           =   2
               Left            =   -735
               TabIndex        =   163
               Top             =   3570
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   1680
               TabIndex        =   162
               Top             =   3465
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   2100
               TabIndex        =   161
               Top             =   3465
               Width           =   120
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "T.Limpieza:"
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
               Index           =   3
               Left            =   -735
               TabIndex        =   160
               Top             =   3885
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   1680
               TabIndex        =   159
               Top             =   3780
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   2100
               TabIndex        =   158
               Top             =   3780
               Width           =   120
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Eq.Aux (1):"
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
               Index           =   1
               Left            =   3045
               TabIndex        =   157
               Top             =   2730
               Width           =   960
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Eq.Aux (2):"
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
               Index           =   2
               Left            =   3045
               TabIndex        =   156
               Top             =   3045
               Width           =   960
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Eq.Aux (3):"
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
               Index           =   3
               Left            =   3045
               TabIndex        =   155
               Top             =   3360
               Width           =   960
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% Demora Distribuida:"
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
               Left            =   3150
               TabIndex        =   154
               Top             =   3675
               Width           =   1065
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Operarios:"
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
               Left            =   -735
               TabIndex        =   153
               Top             =   4515
               Width           =   1905
            End
            Begin VB.Label Label51 
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
               Height          =   285
               Left            =   3045
               TabIndex        =   152
               Top             =   4410
               Width           =   1830
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Cat:"
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
               Left            =   1785
               TabIndex        =   151
               Top             =   4515
               Width           =   540
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mant.Herr:"
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
               Index           =   4
               Left            =   -735
               TabIndex        =   150
               Top             =   5145
               Width           =   1905
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   1680
               TabIndex        =   149
               Top             =   5040
               Width           =   120
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   2100
               TabIndex        =   148
               Top             =   5040
               Width           =   120
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ciclo M-H:"
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
               Left            =   2205
               TabIndex        =   147
               Top             =   5145
               Width           =   1905
            End
            Begin VB.Label Label54 
               BackStyle       =   0  'Transparent
               Caption         =   "    H.    M.    S."
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
               Left            =   1260
               TabIndex        =   146
               Top             =   2610
               Width           =   1380
            End
            Begin VB.Label Label55 
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
               Height          =   285
               Left            =   1260
               TabIndex        =   145
               Top             =   420
               Width           =   570
            End
            Begin VB.Label Label60 
               Alignment       =   2  'Center
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
               Height          =   285
               Left            =   1260
               TabIndex        =   144
               Top             =   1155
               Width           =   885
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:"
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
               Left            =   -735
               TabIndex        =   143
               Top             =   1240
               Width           =   1905
            End
            Begin VB.Label Label62 
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
               Height          =   285
               Left            =   3045
               TabIndex        =   142
               Top             =   1155
               Width           =   1830
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Resp:"
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
               Left            =   1050
               TabIndex        =   141
               Top             =   1245
               Width           =   1905
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Operaciones Standard"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Left            =   105
            TabIndex        =   139
            Top             =   210
            Width           =   4530
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   105
               TabIndex        =   55
               Top             =   420
               Width           =   4320
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00E0E0E0&
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
               Height          =   370
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   59
               ToolTipText     =   "Copia la Secuencia de Operaciones de un Código a Otro"
               Top             =   2520
               Width           =   945
            End
            Begin VB.CommandButton Command32 
               BackColor       =   &H00E0E0E0&
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
               Height          =   370
               Left            =   1390
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Elimina la Operación Activa"
               Top             =   2520
               Width           =   630
            End
            Begin VB.CommandButton Command33 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   370
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   2520
               Width           =   630
            End
            Begin VB.CommandButton Command42 
               BackColor       =   &H00E0E0E0&
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
               Height          =   370
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   60
               ToolTipText     =   "Emite listado de Secuencia de Operaciones por Impresora"
               Top             =   2520
               Width           =   945
            End
            Begin VB.CommandButton Command43 
               BackColor       =   &H00E0E0E0&
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
               Height          =   370
               Left            =   745
               Style           =   1  'Graphical
               TabIndex        =   57
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
               Top             =   2520
               Width           =   630
            End
         End
         Begin VB.Frame Frame10 
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
            Height          =   2430
            Left            =   105
            TabIndex        =   138
            Top             =   3360
            Width           =   4530
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
               Height          =   1905
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   315
               Width           =   4320
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   5895
         Index           =   2
         Left            =   -74895
         TabIndex        =   120
         Top             =   420
         Width           =   10200
         Begin VB.Frame Frame2 
            Caption         =   "Actualizar Fórmula / Estructura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            Left            =   210
            TabIndex        =   129
            Top             =   3990
            Width           =   8100
            Begin VB.CommandButton Command27 
               Caption         =   "Append"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "AMASTO00.frx":2B82
               TabIndex        =   43
               ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
               Top             =   410
               Width           =   855
            End
            Begin VB.TextBox Text32 
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
               Left            =   4935
               TabIndex        =   41
               Text            =   " "
               Top             =   420
               Width           =   1170
            End
            Begin VB.TextBox Text31 
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
               Left            =   1260
               TabIndex        =   39
               Top             =   420
               Width           =   1800
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Insert"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "AMASTO00.frx":2E8C
               TabIndex        =   44
               ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
               Top             =   695
               Width           =   855
            End
            Begin VB.TextBox Text27 
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
               Left            =   6510
               TabIndex        =   42
               Text            =   " "
               Top             =   420
               Width           =   330
            End
            Begin VB.CommandButton Command19 
               Caption         =   "Delete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "AMASTO00.frx":3196
               TabIndex        =   45
               ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
               Top             =   970
               Width           =   855
            End
            Begin VB.CommandButton Command31 
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
               Height          =   360
               Left            =   3045
               TabIndex        =   40
               Top             =   420
               Width           =   175
            End
            Begin VB.CommandButton Command18 
               Caption         =   "Edit"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   275
               Left            =   7035
               Picture         =   "AMASTO00.frx":34A0
               TabIndex        =   46
               ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Uso:"
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
               TabIndex        =   136
               Top             =   525
               Width           =   855
            End
            Begin VB.Label Label25 
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
               Height          =   225
               Left            =   210
               TabIndex        =   135
               Top             =   525
               Width           =   960
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "U-M:"
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
               Left            =   2940
               TabIndex        =   134
               Top             =   525
               Width           =   855
            End
            Begin VB.Label Label23 
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
               Height          =   225
               Left            =   105
               TabIndex        =   133
               Top             =   945
               Width           =   1065
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "D:"
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
               Left            =   5880
               TabIndex        =   132
               Top             =   525
               Width           =   540
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   1260
               TabIndex        =   131
               Top             =   945
               Width           =   5580
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Left            =   3885
               TabIndex        =   130
               Top             =   420
               Width           =   435
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   210
            ScaleHeight     =   375
            ScaleWidth      =   8025
            TabIndex        =   127
            Top             =   390
            Width           =   8085
            Begin VB.Label Label27 
               BackColor       =   &H00E0E0E0&
               Caption         =   " Codigo                   Descripcion                                       U-M              Uso     D"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   128
               Top             =   105
               Width           =   7995
            End
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   210
            TabIndex        =   126
            Top             =   840
            Width           =   8100
         End
         Begin VB.CommandButton Command41 
            Height          =   435
            Left            =   8715
            Picture         =   "AMASTO00.frx":37AA
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   380
            Width           =   1170
         End
         Begin VB.Frame Frame12 
            Caption         =   "Consultas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1210
            Left            =   8505
            TabIndex        =   125
            Top             =   4515
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Implosión"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   6
               Left            =   210
               TabIndex        =   54
               ToolTipText     =   "Consulta de Conjuntos a los que Pertenece el Elemento Activo"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   4
               Left            =   210
               TabIndex        =   53
               ToolTipText     =   "Muestra Estructura Multinivel"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame Frame11 
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
            Height          =   1530
            Left            =   8505
            TabIndex        =   124
            Top             =   2940
            Width           =   1590
            Begin VB.CommandButton Command40 
               Caption         =   "Enlaces"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   0
               Left            =   210
               TabIndex        =   50
               ToolTipText     =   "Listado de Estucturas y Enlaces"
               Top             =   380
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Multinivel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   1
               Left            =   210
               TabIndex        =   51
               ToolTipText     =   "Lista de Materiales Multinivel"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "L.Compras"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   2
               Left            =   210
               TabIndex        =   52
               ToolTipText     =   "Lista de Compras"
               Top             =   1100
               Width           =   1170
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Utilidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   8505
            TabIndex        =   123
            Top             =   1785
            Width           =   1590
            Begin VB.CommandButton Command40 
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
               Height          =   280
               Index           =   3
               Left            =   210
               TabIndex        =   49
               ToolTipText     =   "Control de Integridad de la Base de Datos"
               Top             =   735
               Width           =   1170
            End
            Begin VB.CommandButton Command40 
               Caption         =   "Copiar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Index           =   5
               Left            =   210
               TabIndex        =   48
               ToolTipText     =   "Nueva Fórmula por Copia de una Pre-Existente"
               Top             =   360
               Width           =   1170
            End
         End
         Begin VB.Frame FRAME15 
            Caption         =   "Revisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   8505
            TabIndex        =   121
            Top             =   945
            Width           =   1590
            Begin VB.Label Label59 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   210
               TabIndex        =   122
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Line Line1 
            X1              =   8295
            X2              =   8295
            Y1              =   380
            Y2              =   840
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   5685
         Index           =   0
         Left            =   -74895
         TabIndex        =   101
         Top             =   525
         Width           =   10215
         Begin VB.PictureBox Picture1 
            Height          =   380
            Left            =   1050
            ScaleHeight     =   315
            ScaleWidth      =   420
            TabIndex        =   188
            Top             =   4000
            Visible         =   0   'False
            Width           =   480
            Begin VB.Image Image1 
               Height          =   330
               Left            =   0
               Picture         =   "AMASTO00.frx":3AB4
               Stretch         =   -1  'True
               ToolTipText     =   "Doble-Click para Acceso a Planera Electrónica"
               Top             =   0
               Width           =   435
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Expedicion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   210
            TabIndex        =   185
            Top             =   4725
            Visible         =   0   'False
            Width           =   1905
            Begin VB.OptionButton Option2 
               Caption         =   "Componentes"
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
               TabIndex        =   187
               Top             =   550
               Width           =   1485
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Conjunto"
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
               TabIndex        =   186
               Top             =   330
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin VB.TextBox Text28 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   2310
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   3255
            Width           =   7785
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Cortar"
            Height          =   260
            Left            =   5145
            TabIndex        =   22
            Top             =   5385
            Width           =   1170
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Copiar"
            Height          =   260
            Left            =   6330
            TabIndex        =   23
            Top             =   5385
            Width           =   1170
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Pegar"
            Height          =   260
            Left            =   7500
            TabIndex        =   24
            Top             =   5385
            Width           =   1275
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Selecc Todo"
            Height          =   260
            Left            =   8790
            TabIndex        =   25
            Top             =   5385
            Width           =   1320
         End
         Begin VB.CommandButton Command7 
            Height          =   260
            Left            =   2310
            TabIndex        =   118
            Top             =   5385
            Width           =   2825
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   2310
            ScaleHeight     =   375
            ScaleWidth      =   7725
            TabIndex        =   116
            Top             =   2835
            Width           =   7780
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Descripción General - Especificaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   117
               Top             =   105
               Width           =   7680
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Reposición"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   5145
            TabIndex        =   103
            Top             =   105
            Width           =   4950
            Begin VB.Frame Frame1 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   780
               Left            =   105
               TabIndex        =   179
               Top             =   1680
               Width           =   4700
               Begin VB.TextBox Text24 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1365
                  TabIndex        =   18
                  Top             =   60
                  Width           =   1200
               End
               Begin VB.TextBox Text10 
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   3675
                  MaxLength       =   5
                  TabIndex        =   19
                  Top             =   60
                  Width           =   780
               End
               Begin VB.CommandButton Command14 
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
                  Left            =   4455
                  TabIndex        =   20
                  Top             =   60
                  Width           =   175
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kilos por Unidad:"
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
                  Left            =   0
                  TabIndex        =   182
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Proveedor Habitual:"
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
                  Left            =   2100
                  TabIndex        =   181
                  ToolTipText     =   "Doble-Click para Control Cambios Proveedor Habitual"
                  Top             =   0
                  Width           =   1500
               End
               Begin VB.Label Label33 
                  BackColor       =   &H00E0E0E0&
                  BorderStyle     =   1  'Fixed Single
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   210
                  TabIndex        =   180
                  Top             =   420
                  Width           =   4440
               End
            End
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               Caption         =   "Frame9"
               Height          =   645
               Left            =   210
               TabIndex        =   175
               Top             =   945
               Width           =   4650
               Begin VB.TextBox Text19 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   16
                  Top             =   270
                  Width           =   1305
               End
               Begin VB.TextBox Text18 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   105
                  TabIndex        =   15
                  Top             =   270
                  Width           =   1305
               End
               Begin VB.TextBox Text20 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   3255
                  TabIndex        =   17
                  Top             =   270
                  Width           =   1305
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Económico:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1365
                  TabIndex        =   178
                  Top             =   0
                  Width           =   1245
               End
               Begin VB.Label Label7 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Stock Minimo:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   0
                  TabIndex        =   177
                  Top             =   0
                  Width           =   1395
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lote Máximo:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   3045
                  TabIndex        =   176
                  Top             =   0
                  Width           =   1245
               End
            End
            Begin VB.TextBox Text8 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1260
               MaxLength       =   2
               TabIndex        =   13
               Text            =   " "
               Top             =   420
               Width           =   360
            End
            Begin VB.CommandButton Command12 
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
               Left            =   1575
               TabIndex        =   14
               Top             =   420
               Width           =   175
            End
            Begin VB.Label Label32 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1890
               TabIndex        =   105
               Top             =   420
               Width           =   2895
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Criterio de Reposicion:"
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
               Left            =   -210
               TabIndex        =   104
               Top             =   360
               Width           =   1380
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Características"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   210
            TabIndex        =   102
            Top             =   105
            Width           =   4740
            Begin VB.TextBox Text36 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3780
               TabIndex        =   12
               Top             =   2100
               Width           =   780
            End
            Begin VB.TextBox Text12 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   420
               TabIndex        =   11
               Top             =   2100
               Width           =   2670
            End
            Begin VB.TextBox Text25 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3420
               TabIndex        =   9
               Text            =   " "
               Top             =   1680
               Width           =   945
            End
            Begin VB.CommandButton Command20 
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
               Left            =   4380
               TabIndex        =   10
               Top             =   1680
               Width           =   175
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
               Left            =   1700
               TabIndex        =   8
               Top             =   975
               Width           =   175
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   7
               Text            =   " "
               Top             =   975
               Width           =   570
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
               Left            =   1700
               TabIndex        =   6
               Top             =   525
               Width           =   175
            End
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   5
               Text            =   " "
               Top             =   525
               Width           =   570
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rev:"
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
               Left            =   1050
               TabIndex        =   184
               Top             =   2160
               Width           =   2655
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nro. de Plano / Molde:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   -170
               TabIndex        =   183
               Top             =   1845
               Width           =   2325
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Revisión:"
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
               Left            =   2205
               TabIndex        =   111
               Top             =   1550
               Width           =   1170
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Material:"
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
               TabIndex        =   110
               Top             =   420
               Width           =   1380
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unidad de Medida:"
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
               Left            =   -105
               TabIndex        =   109
               Top             =   840
               Width           =   1185
            End
            Begin VB.Label Label30 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1995
               TabIndex        =   94
               Top             =   975
               Width           =   2535
            End
            Begin VB.Label Label29 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1995
               TabIndex        =   93
               Top             =   525
               Width           =   2535
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción y Especificaciones:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   315
            TabIndex        =   119
            Top             =   3255
            Width           =   1905
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   99
         Top             =   4305
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   98
         Top             =   3360
         Width           =   15
      End
   End
   Begin VB.Label Label58 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   210
      TabIndex        =   115
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3465
      TabIndex        =   108
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   107
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   5040
      TabIndex        =   106
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6195
      TabIndex        =   100
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   105
      TabIndex        =   96
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   375
      TabIndex        =   95
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "AMASTO00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%(15), MODIFIC%(15)
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
'
Dim RESECO$(128), DESOPE$(128)
Dim MOREPO%, AREPO$
'
Private Sub Command1_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
End Sub

Private Sub Command10_Click()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Command11_Click()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Command12_Click()
   Call SELECHO(AREPO$)
   VALONUE$ = VALACT1$(AREPO$)
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   TTXXA$ = Clipboard.GetText
   ACONEC$ = "BATRIB03": If TRIM$(Text1) <> "" Then ACONEC$ = ACONEC$ + "/" + TRIM$(Text1)
   Call CONECRUN(ACONEC$, "Clasificación Técnica")
   TTXXD$ = Clipboard.GetText
   If TTXXD$ <> TTXXA$ Then
     Text1 = TTXXD$
     Clipboard.SetText (TTXXA$)
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
   End If
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command18_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Edición no Válida")
       Exit Sub
    End If
    '
    Text31.TEXT = Left$(List1.TEXT, 16)
    CDYY% = CODINT%(Text31.TEXT)
    Label40.Caption = DESCRIT0$(CDYY%)
    Label41.Caption = UNIMED$(CDYY%)
    Text32.TEXT = TRIM$(Mid$(List1.TEXT, 49, 12))
    Text27.TEXT = TRIM$(Mid$(List1.TEXT, 61, 4))
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command18.Enabled = False
    Command19.Enabled = False
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    Text32.SetFocus
    MODOEDIT% = 1
    MODIFIC%(1) = 1
    '
End Sub

Private Sub Command19_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Borrado no Válida")
       Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    MODIFIC%(1) = 1
    '
End Sub

Private Sub Command2_Click()
   '
   If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
      Exit Sub
   End If
   '
   For Index = 0 To 15
     YACAR%(Index) = 0
     MODIFIC%(Index) = 0
   Next Index
   '
   Command18.Enabled = True
   Command19.Enabled = True
   MODOEDIT% = 0
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Command20_Click()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub COMMAND21_CLICK()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command21.Enabled = False
   Call GRADATITEM
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
      CODINUE$ = Mid$(VDEF$, 77, 16)
      Text1.TEXT = CODINUE$
      SSTab1.Tab = 0
   End If
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   CII% = CODINT%(Text1.TEXT)
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   If ULTREG&("TABLATA") > 0 Then
     TTXX$ = "Elija Modalidad de Alta por Copia:\\Copia Simple\Generación por Tabla de Talles\Generación por Tabla de Colores"
     OPXX% = COMALTER%(TTXX$)
     If OPXX% < 1 Or OPXX% > 3 Then GoTo 99
     If OPXX% = 2 Then
       Call GENERTALLES
       GoTo 99
     End If
     If OPXX% = 3 Then
       FORSELCOL.Show 1
       GoTo 99
     End If
   End If
   '
   REMAS& = 1 + CII%
   If REMAS& < 2 Or REMAS& > ULTREG&("MASTER") Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   VDEF$ = REGLEIDO$("MASTER", REMAS&)
   CINUE% = 1 + NUMAS%
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, Space$(16), 77, 16)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
      Call COPYFORMU(CII%, CINUE%)
      Call COPYRUTA(CII%, CINUE%)
      Call COPYSPEC(CII%, CINUE%)
      CODINUE$ = Mid$(VDEF$, 77, 16)
      Text1.TEXT = CODINUE$
      SSTab1.Tab = 0
   End If
   '
99 Command22.Enabled = True
   '
End Sub

Sub GENERTALLES()
   Dim COTALLE$(), DETALLE$(), FRECUTA()
   '
   CATAX% = ULTREG&("TABLATA")
   ReDim COTALLE$(CATAX%), DETALLE$(CATAX%), FRECUTA(CATAX%)
   '
   FORSELTA.Caption = TRIM$(AMASTO00.Text1) + " - " + TRIM$(AMASTO00.Text3)
   FORSELTA.LICOLTA.Clear
   '
   For U& = 1 To ULTREG&("TABLATA")
     X$ = REGLEIDO$("TABLATA", U&)
     CODTALLE$ = AJUSTI$(Mid$(X$, 13, 2), 2)
     FORSELTA.LICOLTA.AddItem CODTALLE$ + " - " + TRIM$(Mid$(X$, 15, 16))
   Next U&
   '
   FORSELTA.Show 1
   '
   U1& = 0
   For U& = 1 To CATAX%
     If FORSELTA.LICOLTA.Selected(U& - 1) = True Then
       U1& = U1& + 1
       XX$ = REGLEIDO$("TABLATA", U&)
       COTALLE$(U1&) = TRIM$(Mid$(XX$, 13, 2))
       DETALLE$(U1&) = TRIM$(Mid$(XX$, 15, 16))
       FRECUTA(U1&) = CVS(Mid$(XX$, 31, 4))
     End If
   Next U&
   If U1& < 1 Then Exit Sub
   '
   CATAX% = U1&
   CODECO$ = TRIM$(Text1)
   DEDECO$ = TRIM$(Text3)
   '
   If COMALTER%("Está Seguro de Generar Códigos por Talle\Para '" + CODECO$ + " - " + DEDECO$ + "' ?\\No, Cancelar\Si, Generar") <> 2 Then
     Exit Sub
   End If
   '
   CIORI% = CODINT%(Text1)
   JJ1& = 0: REBLX$ = REGBLAN$("DISTALLE")
   For U& = 1 To CATAX%
     COTA$ = COTALLE$(U&)
     DETA$ = DETALLE$(U&)
     If COTA$ <> "" Then
       If DETA$ <> "" Then
         CODINUE$ = CODECO$ + "/" + COTA$
         DESCRINUE$ = DEDECO$ + " - " + DETA$
         CINUE% = CODINT%(CODINUE$)
         If CINUE% > 0 Then
             REMANU& = CINUE% + 1
             XMX$ = REGLEIDO$("MASTER", REMANU&)
             Call REPLA(XMX$, DESCRINUE$, 5, 64)
             Call REPLA(XMX$, CODINUE$, 77, 16)
             Call GRAREG("MASTER", XMX$, REMANU&)
           Else
             NMS% = NUMAS% + 1
             REMANU& = NMS% + 1
             REMANAN& = CIORI% + 1
             CINUE% = NMS%
             XMX$ = REGLEIDO$("MASTER", REMANAN&)
             Call REPLA(XMX$, MKI$(CINUE%), 1, 2)
             Call REPLA(XMX$, DESCRINUE$, 5, 64)
             Call REPLA(XMX$, CODINUE$, 77, 16)
             Call GRAREG("MASTER", XMX$, REMANU&)
             '
             XMX$ = REGLEIDO$("MASTER", 1&)
             Call REPLA(XMX$, MKI$(NMS%), 1, 2)
             Call GRAREG("MASTER", XMX$, 1&)
         End If
         XYX$ = REBLX$
         Call REPLA(XYX$, MKI$(CIORI%), 1, 2)
         Call REPLA(XYX$, MKI$(CINUE%), 3, 2)
         Call REPLA(XYX$, MKS$(FRECUTA(U&)), 5, 4)
         Call REPLA(XYX$, COTA$, 33, 2)
         Call REPLA(XYX$, DETA$, 35, 16)
         JJ1& = JJ1& + 1
         Call GRAREG("!DISTALLE", XYX$, JJ1&)
       End If
     End If
   Next U&
   Call SETULTREG("!DISTALLE", JJ1&)
   Call FILTERUPDATE("DISTALLE", "", 1, MKI$(CIORI%))
   '
   For U& = 1 To CATAX%
     COTA$ = COTALLE$(U&)
     DETA$ = DETALLE$(U&)
     If COTA$ <> "" Then
       If DETA$ <> "" Then
         Call AREXPLO(CIORI%, 1, 1)
         For KKI% = 1 To CAIT%
           CODECO$ = TRIM$(Left$(EXPLOLIN$(KKI%), 16))
           CODINUE$ = CODECO$ + "/" + COTA$
           DEDECO$ = TRIM$(DESCRIT0$(CODINT%(CODECO$)))
           DESCRINUE$ = DEDECO$ + " - " + DETA$
           '
           CINUE% = CODINT%(CODINUE$)
           If CINUE% > 0 Then
               REMANU& = CINUE% + 1
               XMX$ = REGLEIDO$("MASTER", REMANU&)
               Call REPLA(XMX$, DESCRINUE$, 5, 64)
               Call REPLA(XMX$, CODINUE$, 77, 16)
               Call GRAREG("MASTER", XMX$, REMANU&)
             Else
               NMS% = NUMAS% + 1
               REMANU& = NMS% + 1
               REMANAN& = CODINT%(CODECO$) + 1
               XMX$ = REGLEIDO$("MASTER", REMANAN&)
               Call REPLA(XMX$, MKI$(NMS%), 1, 2)
               Call REPLA(XMX$, DESCRINUE$, 5, 64)
               Call REPLA(XMX$, CODINUE$, 77, 16)
               Call GRAREG("MASTER", XMX$, REMANU&)
               '
               XMX$ = REGLEIDO$("MASTER", 1&)
               Call REPLA(XMX$, MKI$(NMS%), 1, 2)
               Call GRAREG("MASTER", XMX$, 1&)
           End If
         Next KKI%
       End If
     End If
   Next U&
   '
   Call CIERRARCH("MASTER")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   ATENA% = ATENTI%
   ATENTI% = (-123)
   Call Command25_Click
   ATENTI% = ATENA%
   '
   '
   CIORI% = CODINT%(Text1)
   For U& = 1 To CATAX%
     COTA$ = COTALLE$(U&)
     DETA$ = DETALLE$(U&)
     If COTA$ <> "" Then
       If DETA$ <> "" Then
         CICOP% = CODINT%(Text1 + "/" + COTA$)
         If CICOP% > 0 Then
           Call COPYRUTA(CIORI%, CICOP%)
           Call COPYSPEC(CIORI%, CICOP%)
           Call COPYPLANOS(CIORI%, CICOP%)
         End If
         '
         Call AREXPLO(CIORI%, 1, 1)
         For KKI% = 1 To CAIT%
           CODECO$ = TRIM$(Left$(EXPLOLIN$(KKI%), 16))
           CODINU$ = CODECO$ + "/" + COTA$
           CODORI% = CODINT%(CODECO$)
           CODEST% = CODINT%(CODINU$)
           If CODORI% > 0 Then
             If CODEST% > 0 Then
               Call COPYFORMU(CODORI%, CODEST%)
               Call COPYRUTA(CODORI%, CODEST%)
               Call COPYSPEC(CODORI%, CODEST%)
               Call COPYPLANOS(CIORI%, CICOP%)
             End If
           End If
           '
           TTXX$ = EXPLOLIN$(KKI%)
           Call REPLA(TTXX$, CODINU$, 1, 16)
           EXPLOLIN$(KKI%) = TTXX$
         Next KKI%
         Call AREXPLO(CICOP%, 1, 2)
         '
       End If
     End If
   Next U&
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call COMUNI("Generación Completada")
   '
End Sub

Private Sub Command23_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Command23.Enabled = False
   If VALCODART% < 1 Then GoTo 99
   '
   VDEF$ = ""
10 OBX$ = OBJPLA$("ENTRACODNUE", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   CODINUE$ = TRIM$(OBX$)
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 16) = CODIVIE$ Then
         Mid$(PUNCODIG$, U&, 16) = AJUSTI$(CODINUE$, 16)
         GoTo 20
      End If
   Next U&
   Call COMUNI("Operación Cancelada (1) - Reintente")
   GoTo 99
   '
20 REMAS& = 1 + (U& + 15) / 16
   NMSI% = REMAS& - 1
   XX$ = REGLEIDO$("MASTER", REMAS&)
   If Mid$(XX$, 77, 15) <> Left$(CODIVIE$, 15) Then
      Call COMUNI("Operación Cancelada (2) - Reintente")
      GoTo 99
   End If
   Call REPLA(XX$, CODINUE$, 77, 16)
   Call GRAREG("MASTER", XX$, REMAS&)
   Call CIERRARCH("MASTER")
   '
   PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   '
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Text3.TEXT, 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) = CODIVIE Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 64)
         GoTo 30
      End If
   Next U&
   '
30 For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 40
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
40 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
50 NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) = CODIVIE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 18)
         GoTo 60
      End If
   Next U&
   '
60 For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   Screen.MousePointer = 1
   '
   Text1.TEXT = CODINUE$
   '
99 Screen.MousePointer = 1
   Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Dim EL As String * 2048
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), 15)
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call BLANFORMU
   Screen.MousePointer = 11
   Text1.TEXT = ""
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   REGMA& = CII% + 1
   X$ = REGLEIDO$("MASTER", REGMA&)
   If Mid$(X$, 77, 15) <> AJUSTI$(COD$, 15) Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Call REPLA(X$, Chr$(96), 92, 1)
   Call GRAREG("MASTER", X$, REGMA&)
   Call CIERRARCH("MASTER")
   '
10 NVII& = 1 + NUINV%
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 15) = COD$ Then
         Mid$(TTXX$, U&, 16) = COD$ + Chr$(96)
         GoTo 20
      End If
   Next U&
   Close #FIECO%
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
20 JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 15) = COD$ Then
         Mid$(PUNCODIG$, U&, 16) = COD$ + Chr$(96)
         GoTo 25
      End If
   Next U&
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
25 PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   LOTOTE& = 64 * ULTREG&("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 15) = COD$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 64)
         GoTo 30
      End If
   Next U&
   '
30 While (Len(TTXX$) Mod 2048) <> 0: TTXX$ = TTXX$ + Chr$(0): Wend
   JJ& = 0
   
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   Call FRESHECHO("MASTER")
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    If NUMAS% < 1 Then GoTo 99
    '
    If ATENTI% = (-123) Then GoTo 20
    If COMALTER%("Realmente Desea Controlar el Maestro\\No, Cancelar\Si, Controlar") = 2 Then
20    Call BLOQARCH("MASTER")
      Call REVIMASTER
      Call REPUNCODIG
    End If
99  Command25.Enabled = True
End Sub

Private Sub Command26_Click()
   '
   If VALCODART% < 1 Then Exit Sub
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(Val(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
   CAIT% = CAIT% + 1
   ILISTA& = List1.ListIndex
   If ILISTA& < 0 Then ILISTA& = 0
   '
   List1.AddItem XXX$, ILISTA&
   List1.ListIndex = ILISTA&
   List1.Refresh
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text27.TEXT = ""
   Text31.SetFocus
   MODIFIC%(1) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
   If VALCODART% < 1 Then Exit Sub
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32.TEXT)
   If Abs(USO) < 0.0001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32.SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   USOX$ = CSTRING$(MKS$(USO), 4, 12, 4, 2)
   PD = Int(Val(Text27.TEXT))
   If PD < 1 Then PD = 0
   If PD > 9 Then PD = 9
   PDX$ = CSTRING$(Chr$(PD), 1, 4, 0, 2)
   '
   XXX$ = Space$(64)
   Call REPLA(XXX$, CODXX$, 1, 16)
   Call REPLA(XXX$, DEXX$, 17, 28)
   Call REPLA(XXX$, UNXX$, 47, 2)
   Call REPLA(XXX$, USOX$, 49, 12)
   Call REPLA(XXX$, PDX$, 61, 4)
   CAIT% = CAIT% + 1
   '
   List1.AddItem XXX$
   List1.ListIndex = List1.ListCount - 1
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   Text32.TEXT = ""
   Text27.TEXT = ""
   Text31.SetFocus
   MODIFIC%(1) = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command28_Click(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command29_Click()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Command3_Click(Index As Integer)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command30_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command30.Enabled = False
   Call GRADATITEM
   For Index = 0 To 15: YACAR%(Index) = 0: Next Index
   Call COPYRUTAS
   Command30.Enabled = True
End Sub

Private Sub COMMAND31_CLICK()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
    '
End Sub

Private Sub Command32_Click()
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   ILISTA& = List3.ListIndex
   If ILISTA& >= 0 Then
     List3.RemoveItem (ILISTA&)
     For UI& = ILISTA& To List3.ListCount - 1
       RESECO$(UI&) = RESECO$(UI& + 1)
       DESOPE$(UI&) = DESOPE$(UI& + 1)
     Next UI&
   End If
   '
   Label55.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Label51.Caption = ""
   MODIFIC(2) = 1
   '
   If List3.ListCount > 0 Then
     List3.ListIndex = 0
   End If
   '
End Sub

Private Sub Command33_Click()
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   NOPY% = 10
   If List3.ListCount > 0 Then
     ULI& = List3.ListCount
     NOPY% = 10 * (1 + Int(Val(Left$(List3.List(ULI& - 1), 9) / 10)))
   End If
   RESECO$(List3.ListCount) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(List3.ListCount) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Agregada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(2) = 1
   List3.AddItem SECO$
   List3.ListIndex = List3.ListCount - 1
   Text5.SetFocus
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   'OPMAST00.Show 1
   Call CONECRUN("FLSETUP/OPMAST00", "Configuración de Funcionamiento")
   Command37.Enabled = True
End Sub

Private Sub Command38_Click()
   Command38.Enabled = False
   OPLMAS00.Show 1
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   FORIPRE$ = TRIM$(CONTROL$("", "FICHATEC"))
   If FORIPRE$ = "" Or EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       PRINTFORM
     Else
       Call PRIFICHTEC
   End If
   Command39.Enabled = True
End Sub

Sub PRIFICHTEC()
   '
   CIXI% = CODINT%(Text1)
   If CIXI% < 1 Then Exit Sub
   '
   FORIPRE$ = TRIM$(CONTROL$("", "FICHATEC"))
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
   '
   FECHDOC$ = FECHATEX$(HOY(HOS$))
   NUMEDOC$ = ""
   TIPODOC$ = "Ficha Tecnica " + CODEXT$(CIXI%)
   '
   CODPARAM$(1) = "COD-MAT"
   CODPARAM$(2) = "DES-MAT"
   CAPARDOC% = 2
   '
   DOCPARAM$(1) = CODEXT$(CIXI%)
   DOCPARAM$(2) = DESCRIT0$(CIXI%)
   '
   CODTABU1$(1) = "COD-MPRI"
   CODTABU1$(2) = "DES-MPRI"
   CODTABU1$(3) = "UNI-MPRI"
   CODTABU1$(4) = "USO-MPRI"
   CACOLTAB1% = 4
   '
   CODTABU2$(1) = "ESP-MAT"
   CACOLTAB2% = 1
   '
   DESTIDOC% = 0
   CAITAB1% = 0
   CAITAB2% = 0
   CAITCRU% = 0
   '
   Call EXPLOCOM(CIXI%, 1)
   For KKK% = 1 To CONTAMUL%
     CIII% = CIK%(KKK%)
     CAITAB1% = CAITAB1% + 1
     '
     TETABU1$(1, CAITAB1%) = CODEXT$(CIII%)
     TETABU1$(2, CAITAB1%) = DESCRIT0$(CIII%)
     TETABU1$(3, CAITAB1%) = UNIMED$(CIII%)
     TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANT(KKK%)))
     '
   Next KKK%
   '
   TTXXX$ = ESPECIFI$(CIXI%)
   Call FRATEXTO(TTXXX$, 128)
   CAITAB2% = CARETEX%
   If CAITAB2% > 32 Then CAITAB2% = 32
   '
   For KKK% = 1 To CAITAB2%
     TETABU2$(1, KKK%) = RETEX$(KKK%)
   Next KKK%
   '
   CAIMAGEN% = 0
   REBLA$ = REGBLAN$("INDIPLAN")
   RCIP$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CIXI%))
   If RCIP$ <> REBLA$ Then
     For KII% = 0 To 3
       COPLANO$ = TRIM$(Mid$(RCIP$, 65 + 48 * KII%, 48))
       If COPLANO$ <> "" Then
         COPLANO$ = RUTARGRA$ + "\" + COPLANO$
         If EXISTE%(COPLANO$) > 0 Then
           CAIMAGEN% = CAIMAGEN% + 1
           CODIMAGEN$(CAIMAGEN%) = "DOCUGRA" + TRIM$(Str$(KII%))
           ARCHIMAGEN$(CAIMAGEN%) = COPLANO$
         End If
       End If
     Next KII%
   End If
   '
   DESTIDOC% = 0
   TIPODOC$ = "Ficha Tecnica " + TRIM$(CODEXT$(CIXI%))
   Call PRINTDOC("FICHATEC")
   '
End Sub
'
Private Sub Command40_Click(Index As Integer)
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   '
   Command40(Index).Enabled = False
   '
   Call GRADATITEM
   Select Case Index
      Case 0
        Call LIESTRU
      Case 1
        Call LIMULTI
      Case 2
        Call LIMULIN
      Case 3
        Call REDESPIE
      Case 4
        Call COMULTI
      Case 5
        Call COPYSTRUC
      Case 6
        Call LIMPLOSI
   End Select
   '
   Command40(Index).Enabled = True
   '
End Sub

Private Sub Command41_Click()
  '
  If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
  Call GRADATITEM
  '
  If IUCOD% > 0 Then
    IUCOD% = IUCOD% - 1
    Text1.TEXT = UCOD$(IUCOD% + 1)
  End If
End Sub

Private Sub Command42_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command42.Enabled = False
   Call GRADATITEM
   Call LISTRUTAS
   Command42.Enabled = True
End Sub

Private Sub Command43_Click()
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then
      Exit Sub
   End If
   '
   ILISTAX& = List3.ListIndex
   If ILISTAX& < 0 Or ILISTAX& >= List3.ListCount Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   NOPEP% = Val(Left$(List3.List(ILISTAX&), 4))
   NOPEA% = 0
   If ILISTAX& > 0 Then
     NOPEA% = Val(Left$(List3.List(ILISTAX& - 1), 4))
   End If
   NOPY% = (NOPEA% + NOPEP%) / 2
   If NOPY% <= NOPEA% Or NOPY% >= NOPEP% Then
      Call MENSERR(24, "Imposible Insertar")
      Exit Sub
   End If
   '
   For UI& = List3.ListCount To ILISTAX& Step -1
     If UI& > 0 Then
       RESECO$(UI&) = RESECO$(UI& - 1)
       DESOPE$(UI&) = DESOPE$(UI& - 1)
     End If
   Next UI&
   '
   RESECO$(ILISTAX&) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(ILISTAX&) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Insertada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(2) = 1
   List3.AddItem SECO$, ILISTAX&

   List3.ListIndex = ILISTAX&
   Text5.SetFocus
   '
End Sub


Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    SSTab1.Tab = 0
    'Call TRANSLABELS(Name)
    'Show
    'Refresh
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    'Call GRATITFOR(Caption)
    '
    If InStr(EPAC$, "FAEM") > 0 Then
      Frame16.Visible = True
    End If
    '
    If EXISTE%("BATRIB03.EXE") > 0 Then
      Command13.Enabled = True
    End If
    '
    If EXISTE%("PLANEL02.EXE") > 0 Then
      Picture1.Visible = True
    End If
    '
    Refresh
    '
    Call CARTAVALISTO
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
    MOREPO% = 0: AREPO$ = "TACRIRE"
    'lo siguiente se deshabilita porque siempre aparece proveedor habitual
    'If CONTROL$("", "MOREPOSI") = "PROHABIT" Then
    '  MOREPO% = 1
    '  AREPO$ = "PROVED1"
    '  Label11.Caption = "Proveedor Habitual"
    '  Text8.MaxLength = 5
    '  Text8.Width = Text8.Width + 500
    '  Command12.Left = Command12.Left + 500
    '  Label32.Width = Label32.Width - 500
    '  Label32.Left = Label32.Left + 500
    'End If
    '
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   IMPLOSI.Hide
   On Error GoTo 0
   DoEvents
   If WindowState = 1 Then
       Call CIERRARCH("*.*")
       Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MODOEDIT% = 1 Then
     Cancel = 1
     Call MENSAEDIT
     Exit Sub
   End If
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
   Hide
End Sub

Private Sub Image1_DblClick()
  If EXISTE%(App.Path + "\PLANEL02.EXE") > 0 Then
    ACONEC$ = "PLANEL02"
    If CODINT%(Text1) > 0 Then
      ACONEC$ = ACONEC$ + "/" + Text1
    End If
    Call CONECRUN(ACONEC$, "Planera Electrónica")
  End If
End Sub

Private Sub List1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   If IUCOD% < 128 Then
     If TRIM$(Text1) <> "" Then
       If Text1 <> UCOD$(IUCOD%) Then
         IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = Text1.TEXT
                              ICOMPO%(IUCOD%) = List1.ListIndex
       End If
     End If
   End If
   '
   Call GRADATITEM
   COD$ = TRIM$(Left$(List1.TEXT, 16))
   If COD$ <> "" Then
     Text1.TEXT = COD$
     Call CARDATITEM
   End If
   '
End Sub

Private Sub List3_Click()
   Call MUEDATOPER
End Sub


Function VALCODART%()
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
End Function
'
Sub BLANFORMU()
   '
   Text3.TEXT = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
   Text4.TEXT = ""
   Label30.Caption = ""
   Text6.TEXT = ""
   Label29.Caption = ""
   Text8.TEXT = ""
   Label32.Caption = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   Text20.TEXT = ""
   Text24.TEXT = ""
   Text25.TEXT = ""
   '
   Text28.TEXT = ""
   '
   List1.Clear
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   List3.Clear
   Label55.Caption = ""
   Text5.TEXT = ""
   Text7.TEXT = ""
   Label43 = ""
   Text11 = ""
   On Error Resume Next
     For KKI% = 0 To 4
       Text9(KKI%).TEXT = ""
       Text15(KKI%).TEXT = ""
       Text17(KKI%).TEXT = ""
       Text26(KKI%).TEXT = ""
     Next KKI%
   On Error GoTo 0
   '
   Text2.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Label51.Caption = ""
   '
End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
          Call MENSERR(24, "Código Inexistente o no Válido")
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    MODIFIC(0) = 0
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENITEM
        Call CARESPECI
      Case 1
        Call CARFORMULA
      Case 2
        Call CARHORUTA
    End Select
    Screen.MousePointer = 1
    '
End Sub

Sub CARGENITEM()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      REG& = CI% + 1
      REMAS$ = REGLEIDO$("MASTER", REG&)
      '
        UNID$ = UNIMED$(CI%)
      Text4.TEXT = UNID$
      Label30.Caption = ECOARCH$("UMEDIDA", UNID$)
         TIMA% = Asc(Mid$(REMAS$, 93, 1))
      Text6.TEXT = TRIM$(Str$(TIMA%))
      Label29.Caption = ECOARCH$("TATIMAT", Chr$(TIMA%))
         CRIT% = Asc(Mid$(REMAS$, 94, 1))
      Text8.TEXT = TRIM$(Str$(CRIT%))
      NPRO% = CVI(Mid$(REMAS$, 123, 2))
      Text10 = TRIM$(Str$(NPRO%))
      ' Label32.Caption = ECOARCH$("TACRIRE", Chr$(CRIT%))
      Text18.TEXT = CSTRING$(MKS$(STOMIN(CI%)), 3, 10, 1, 2)
      Text19.TEXT = CSTRING$(MKS$(LOTESTAN(CI%)), 3, 10, 1, 2)
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(CI%)), 3, 10, 1, 2)
         CP = CVS(Mid$(REMAS$, 99, 4))
      Text24.TEXT = LTITPE$(CP)
         FEX = FEREVI%(CI%)
      Text25.TEXT = FECHATEX$(FEX)
         'CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
         'Text25.TEXT = TRIM$(Str$(CCOMPO%))
         'Label39.Caption = TRIM$(ECOARCH$("TACOMPO", Chr$(CCOMPO%)))
      Text36 = RVPLANO$(CI%)
      Text12 = NUPLANO$(CI%)
      '
      Option1.Value = True
      If DESCOMPO%(CI%) = 1 Then Option2.Value = True
      '
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

Sub CARESPECI()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Text28.TEXT = "": TTYZ$ = ""
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            Screen.MousePointer = 11
            RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
            URN& = ULTREG&("ESPECIFI")
            '
            For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                  If RENOTA& <= URN& Then
                     TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                     If TTYZ$ <> "" Then
                        TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
                     End If
                     TTYZ$ = TTYZ$ + RTrim$(TTYY$)
                  End If
               End If
            Next U&
            Screen.MousePointer = 1
         End If
      End If
      LOCADE& = Len(TTYZ$)
      For U& = LOCADE& To 1 Step -1
        If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
      Next U&
      U& = 0
30    Text28.TEXT = Left$(TTYZ$, U&)
      Text28.Refresh
      '
      YACAR%(0) = 1
      MODIFIC%(0) = 0
      '
   End If
   '
End Sub
'
Sub CARFORMULA()
   '
   If YACAR%(1) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      Call LEEEXPLO(CODD$, 1)
      Label59.Caption = FECHATEX$(REVIFORMU)
      List1.Clear
      For I% = 1 To CAIT%
         List1.AddItem EXPLOLIN$(I%)
      Next I%
      INDEXI% = 0
      If IUCOD% >= 0 Then
        If ICOMPO%(IUCOD% + 1) >= 0 Then
          If ICOMPO%(IUCOD% + 1) < List1.ListCount Then
            INDEXI% = ICOMPO%(IUCOD% + 1)
          End If
        End If
      End If
      If List1.ListCount > 0 Then
         List1.ListIndex = INDEXI%
         If List1.TopIndex < List1.ListIndex - 8 Then
           List1.TopIndex = List1.ListIndex - 8
         End If
      End If
      '
      YACAR%(1) = 1
      MODIFIC%(1) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub CARHORUTA()
   '
   If YACAR%(2) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      List3.Clear
      Label55 = ""
      Text5 = ""
      Text7 = ""
      Text11 = ""
      Label43 = ""
      Text29 = ""
      Text30 = ""
      Text33 = ""
      Text34 = ""
      Label51 = ""
      For III% = 0 To 4
        Text15(III%) = ""
        Text17(III%) = ""
        Text26(III%) = ""
      Next III%
      For III% = 0 To 3
        Text9(III%) = ""
      Next III%
      '
      ULTOP% = 0
10    RFFX$ = RECFILT$("SECOPER", 1, MKI$(CI%))
      For I1& = 1 To Len(RFFX$) Step 4
        I& = CVS(Mid$(RFFX$, I1&, 4))
        X$ = REGLEIDO$("SECOPER", I&)
        If CVI(Left$(X$ + String$(2, 0), 2)) <> CI% Then
           X$ = REGLEIDO$("SECOPER", 1)
           Call GRAREG("SECOPER", X$, 1)
           Call CIERRARCH("SECOPER")
           GoTo 10
        End If
        '
        NOPY% = CVI(Mid$(X$, 3, 2))
        If NOPY% > ULTOP% Then ULTOP% = NOPY%
        NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
        DOPE$ = Mid$(X$, 5, 48)
        REGO& = I&
        REGOX$ = AJUSTD$(Str$(REGO&), 6)
        SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
        List3.AddItem SECO$
        RESECO$(List3.ListCount - 1) = Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16)
        DESOPE$(List3.ListCount - 1) = ""
      Next I1&
      '
      RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
      URN& = ULTREG&("OBSEOPER")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
            For KKI% = 1 To List3.ListCount
              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
                DESOPE$(KKI% - 1) = DESOPE$(KKI% - 1) + Mid$(TTZZ$, 5)
              End If
            Next KKI%
          End If
        End If
      Next U&
      '
      If List3.ListCount > 0 Then
        List3.ListIndex = 0
      End If
      '
      YACAR%(2) = 1
      MODIFIC%(2) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub MUEDATOPER()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   '
   RETEMODI3% = MODIFIC%(2)
   '
   CABAS = LOTESTAN(CI%)
   If CABAS < 1 Then CABAS = 1
   If CABAS > 3000000 Then CABAS = 1
   '
   REGI& = Val(Mid$(List3.TEXT, 59, 6))
   NOPY% = Val(Left$(List3.TEXT, 4))
   DOPE$ = Mid$(List3.TEXT, 10, 48)
   '
   VDEF$ = RESECO$(List3.ListIndex)
   '
   Label55.Caption = TRIM$(Str$(NOPY%))
   Text5.TEXT = DOPE$
   Text7.TEXT = CSTRING(MKS$(CABAS), 3, 7, 0, 2)
   '
   Text9(0).TEXT = AJUSTI$(Mid$(VDEF$, 53, 6), 6)
   Label43.Caption = ECOARCH$("PADMAQ", Text9(0).TEXT)
   Text9(1).TEXT = Mid$(VDEF$, 103, 6)
   Text9(2).TEXT = Mid$(VDEF$, 109, 6)
   Text9(3).TEXT = Mid$(VDEF$, 115, 6)
   '
   Text11.TEXT = Mid$(VDEF$, 129, 16)
   '
   TPREP = CVS(Mid$(VDEF$, 59, 4))
   TSTT$ = HORMINSE$(60 * 60 * TPREP)
   Text15(0).TEXT = Left$(TSTT$, 3)
   Text17(0).TEXT = Mid$(TSTT$, 5, 2)
   Text26(0).TEXT = Mid$(TSTT$, 8, 2)
   '
   TFIJO = CVS(Mid$(VDEF$, 63, 4))
   TSTT$ = HORMINSE$(60 * 60 * TFIJO)
   Text15(1).TEXT = Left$(TSTT$, 3)
   Text17(1).TEXT = Mid$(TSTT$, 5, 2)
   Text26(1).TEXT = Mid$(TSTT$, 8, 2)
   '
   TISTA = CABAS * CVS(Mid$(VDEF$, 67, 4))
   TSTT$ = HORMINSE$(60 * TISTA)
   Text15(2).TEXT = Left$(TSTT$, 3)
   Text17(2).TEXT = Mid$(TSTT$, 5, 2)
   Text26(2).TEXT = Mid$(TSTT$, 8, 2)
   '
   TLIMP = CVS(Mid$(VDEF$, 79, 4))
   TSTT$ = HORMINSE$(60 * 60 * TLIMP)
   Text15(3).TEXT = Left$(TSTT$, 3)
   Text17(3).TEXT = Mid$(TSTT$, 5, 2)
   Text26(3).TEXT = Mid$(TSTT$, 8, 2)
   '
   TMHER = CVS(Mid$(VDEF$, 95, 4))
   TSTT$ = HORMINSE$(60 * 60 * TMHER)
   Text15(4).TEXT = Left$(TSTT$, 3)
   Text17(4).TEXT = Mid$(TSTT$, 5, 2)
   Text26(4).TEXT = Mid$(TSTT$, 8, 2)
   '
   Text29.TEXT = CSTRING$(Mid$(VDEF$, 75, 4), 3, 4, 1, 2)
   Text30.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 83, 4), 3, 5, 2, 2))
   Text33.TEXT = TRIM$(Str$(Asc(Mid$(VDEF$, 87, 1))))
   Text34.TEXT = TRIM$(CSTRING$(Mid$(VDEF$, 99, 4), 3, 5, 0, 2))
   Label51.Caption = ECOARCH$("CATOPER", Chr$(Text33.TEXT))
   '
   Text2.TEXT = RTrim$(DESOPE$(List3.ListIndex))
   '
   Label60.Caption = ""
   Label62.Caption = ""
   RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
   URN& = ULTREG&("DATADOPS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         If CVI(Left$(TTXX$, 2)) = CI% Then
           FEX = CVI(Mid$(TTXX$, 25, 2))
           Label60.Caption = FECHATEX$(FEX)
           Label62.Caption = TRIM$(Mid$(TTXX$, 37, 16))
         End If
       End If
     End If
   End If
   '
   MODIFIC%(2) = RETEMODI3%
   '
End Sub
'
Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If MODIFIC%(0) > 0 Then
       On Error Resume Next
       Command9.SetFocus
       On Error GoTo 0
       Call GRAGENITEM
       ' Call GRAESPECI
     ElseIf MODIFIC%(1) > 0 Or (SSTab1.Tab = 1 And GRABATODO% = 1) Then
       Command9.SetFocus
       Call GRAFORMULA
     ElseIf MODIFIC%(2) > 0 Then
       Command9.SetFocus
       Call GRAHORUTA
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'
Sub GRAGENITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") < 2 Then
      Exit Sub
   End If
   '
   If YACAR%(0) < 1 Then    ' SOLO CAMBIO LA DESCRIPCION
      Screen.MousePointer = 11
      REG& = CI% + 1
      REMAS$ = REGLEIDO$("MASTER", REG&)
      '
      MODESCRI% = 0
      DE$ = TRIM$(Text3.TEXT)
      If DE$ = "" Then DE$ = TRIM$(Mid$(REMAS$, 5, 64))
      MODESCRI% = 1
      Call REPLA(REMAS$, DE$, 5, 64)
      GoTo 30
   End If
   '
   If YACAR%(0) < 1 Then Exit Sub
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(0) = 0
       MODIFIC(0) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   TTITPE$ = TRIM$(Text24.TEXT)
   PESU = Val(TTITPE$)
   If InStr(TTITPE$, "/") > 0 Then
      TTQP$ = TTITPE$
      TTQP$ = REPLACAR$(TTQP$, "/", ".")
      PESU = (-1) * Abs(Val(TTQP$))
   End If
   '
   REG& = CI% + 1
   REMAS$ = REGLEIDO$("MASTER", REG&)
   '
   MODESCRI% = 0
   DE$ = TRIM$(Text3.TEXT)
   If DE$ = "" Then DE$ = TRIM$(Mid$(REMAS$, 5, 64))
   If AJUSTI$(DE$, 64) <> Mid$(REMAS$, 5, 64) Then
      MODESCRI% = 1
   End If
   Call REPLA(REMAS$, DE$, 5, 64)
   '
   TIMA% = Val(Text6.TEXT)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   '
   CRIT% = Val(Text8.TEXT)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   '
   NPRO% = Val(Text10)
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   '
   CCOMPO% = 0: If Option2.Value = True Then CCOMPO% = 1
   Call REPLA(REMAS$, Chr$(CCOMPO%), 98, 1)
   '
   Call REPLA(REMAS$, MKS$(PESU), 99, 4)
   '
   STMII% = (-1) * Val(Text18) / 100: If Val(Text18) < 32000 Then STMII% = Val(Text18)
   LOREI% = (-1) * Val(Text19) / 100: If Val(Text19) < 32000 Then LOREI% = Val(Text19)
   LOTEMAI% = (-1) * Val(Text20) / 100: If Val(Text20) < 32000 Then LOTEMAI% = Val(Text20)
   If Abs(LOTEMAI%) < Abs(LOREI%) Then LOTEMAI% = LOREI%
   '
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   '
   UNID$ = TRIM$(Text4.TEXT)
   Call REPLA(REMAS$, UNID$, 109, 2)
   '
   FERE% = FECHANUM(Text25.TEXT)
   Call REPLA(REMAS$, MKI$(FERE%), 121, 2)
   '
30 Call GRAREG("MASTER", REMAS$, REG&)
   Call CIERRARCH("MASTER")
   '
   If MODESCRI% > 0 Then
      CODARTIC$ = AJUSTI$(Text1.TEXT, 16)
      LOTOTE& = 64 * ULTREG&("ECOMAST")
      FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
      Dim EL As String * 2048
      TTXX$ = String$(LOTOTE&, 0)
      '
      For U& = 1 To 1 + LOF(FIECO%) / 2048
         Get #FIECO%, U&, EL$
         On Error Resume Next
         Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
         On Error GoTo 0
      Next U&
      TTXX$ = Left$(TTXX$, LOTOTE&)
      '
      For U& = 1 To Len(TTXX$) Step 64
        If Mid$(TTXX$, U&, 16) = CODARTIC$ Then
          Mid$(TTXX$, U& + 16, 46) = AJUSTI$(DE$, 46)
          GoTo 40
        End If
      Next U&
      '
40    JJ& = 0
      For U& = 1 To Len(TTXX$) Step 2048
        EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
        If Len(EL$) < 2048 Then
          EL$ = Left$(EL$ + String$(2048, 0), 2048)
        End If
        JJ& = JJ& + 1
        Put #FIECO%, JJ&, EL$
      Next U&
      Close #FIECO%
      Call GENECOMASNUE
      'Call FRESHECHO("ECOMAST")
      '
   End If
   '
   NUPLA$ = TRIM$(Text12)
   REVIS$ = TRIM$(Text36)
   If NUPLA$ <> "" Or REVIS$ <> "" Then
     RFG$ = RECFILT$("INDIPLAN", 1, MKI$(CI%))
     REIN& = 1 + ULTREG&("INDIPLAN"): REGIX$ = REGBLAN$("INDIPLAN")
     If Len(RFG$) >= 4 Then
       REIN& = CVS(Mid$(RFG$, 1, 4))
       REGIX$ = REGLEIDO$("INDIPLAN", REIN&)
     End If
     Call REPLA(REGIX$, MKI$(CI%), 1, 2)
     Call REPLA(REGIX$, NUPLA$, 3, 16)
     Call REPLA(REGIX$, REVIS$, 19, 4)
     Call GRAREG("INDIPLAN", REGIX$, REIN&)
     Call CIERRARCH("INDIPLAN")
   End If
   '
   If CI% > 1 Then     ' fuerza actualizacion de descripcion
      DDD = DESCRIT$(CI% - 1)
   End If
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Call GRAESPECI
   '
   Screen.MousePointer = 1
   MODIFIC%(0) = 0
   MODESCRI% = 0
   '
End Sub
'
Sub GRAESPECI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(0) < 1 Then Exit Sub
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   If DERACCE%("MODISPEC", "Modificación de Especificaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       MODIFIC(0) = 0
       Call CARESPECI
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("ESPECIFI")
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 62)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CI%) + RETEX$(LL%) + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ESPECIFI", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("ESPECIFI", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("ESPECIFI")
   MODIFIC(0) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAFORMULA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     RESPU% = COMALTER%("Se ha Modificado la Fórmula del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar")
     If RESPU% <> 1 Then
       YACAR%(1) = 0
       MODIFIC(1) = 0
       Call CARFORMULA
       GoTo 99
     End If
   End If
   '
   RFF1$ = RECFILT$("CARFORLI", 1, MKI$(CI%))
   RFF2$ = RECFILT$("CARFORLI", 4, MKI$(CI%))
   If RFF1$ <> "" Or RFF2$ <> "" Then
     Call MENSERR(24, "Artículo Componente de Fórmula por Lista.\  \Esta Fórmula solo Puede Modificarse\a Partir de la Lista de Componentes, Materiales y Medidas.")
     YACAR%(1) = 0
     MODIFIC(1) = 0
     Call CARFORMULA
     GoTo 99
   End If
   '
   Screen.MousePointer = 11
   CAIT% = List1.ListCount
   '
   Dim EXPLOLI1$(256)
   CAIT1% = CAIT%
   For JJI% = 1 To 256
      EXPLOLI1$(JJI%) = ""
      If JJI% <= List1.ListCount Then
         EXPLOLI1$(JJI%) = List1.List(JJI% - 1)
      End If
   Next JJI%
   '
   For I& = 1 To CAIT%
     EXPLOLIN(I&) = List1.List(I& - 1)
     CICOMPO% = CODINT%(Left$(EXPLOLIN(I&), 16))
     Call EXPLOMUL(CICOMPO%)
     For JJI% = 1 To CONTAMUL%
       If CIK%(JJI%) = CI% Then
         Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".")
         YACAR%(2) = 0
         MODIFIC(1) = 0
         SSTab1.Tab = 2
         Call CARFORMULA
         GoTo 99
       End If
     Next JJI%
   Next I&
   '
   CAIT% = CAIT1%
   For JJI% = 0 To 256
      EXPLOLIN$(JJI%) = EXPLOLI1$(JJI%)
   Next JJI%
   '
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   CABAS = 1
   CAIT% = List1.ListCount
   '
   If CABAS > 0.01 Then
      Call EXPLOGR(CODD$, CABAS)
      MODIFIC%(1) = 0
   End If
   '
99 Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   Command18.Enabled = True
   Command19.Enabled = True
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAHORUTA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(2) < 1 Then MODIFIC%(2) = 0: Exit Sub
   If MODIFIC%(2) < 1 Then Exit Sub
   '
   If DERACCE%("MODIRUTA", "Modificación Hojas de Ruta") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(2) = 0
       MODIFIC(2) = 0
       Call CARHORUTA
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("SECOPER")
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CI%))
   KK& = (-3)
   '
   For LL% = 1 To List3.ListCount
      TTYY$ = MKI$(CI%) + Mid$(RESECO$(LL% - 1), 3, 126)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
           XY$ = REGLEIDO$("SECOPER", RENOTA&)
           If CVI(Left$(XY$, 2)) <> CI% Then
              Call MENSERR(24, "Error de Grabación de Secuencia de Operaciones.")
              URN& = URN& + 1
              RENOTA& = URN&
           End If
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("SECOPER", TTYY$, RENOTA&)
      '
      NOPEX% = CVI(Mid$(TTYY$, 3, 2))
      NHPF$ = Mid$(RESECO$(LL% - 1), 129, 16)
      Call GRADATOPER(CI%, NOPEX%, NHPF$)  ' DATOS ADICIONALES OPERACION
      '
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("SECOPER", String$(128, 0), RENOTA&)
   Wend
   '
   URN& = ULTREG&("OBSEOPER")
   RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CI%))
   KK& = (-3)
   '
   For LL% = 1 To List3.ListCount
     '
     If TRIM$(DESOPE$(LL% - 1)) <> "" Then
       TEXTO$ = RTrim$(DESOPE$(LL% - 1))
       While Len(TEXTO$) > 0
         TTYY$ = MKI$(CI%) + Mid$(RESECO$(LL% - 1), 3, 2) + Left$(TEXTO$, 252)
         KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
             RENOTA& = CVS(Mid$(RNC$, KK&, 4))
           Else
             URN& = URN& + 1
             RENOTA& = URN&
         End If
         Call GRAREG("OBSEOPER", TTYY$, RENOTA&)
         TEXTO$ = Mid$(TEXTO$, 253)
       Wend
     End If
   Next LL%
   '
   While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("OBSEOPER", String$(256, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH("SECOPER")
   Call CIERRARCH("OBSEOPER")
   '
   MODIFIC(2) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRADATOPER(CI%, NOPE%, NHPF$)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         URN& = ULTREG&("DATADOPS")
         RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
         KK& = (-3)
         '
         TTYY$ = REGBLAN$("DATADOPS")
         Call REPLA(TTYY$, MKI$(CI%), 1, 2)
         Call REPLA(TTYY$, MKI$(NOPE%), 3, 2)
         Call REPLA(TTYY$, NHPF$, 5, 16)
         'Call REPLA(TTYY$, Text2.TEXT, 21, 4)
         FEXI% = HOY(HOS$)
         Call REPLA(TTYY$, MKI$(FEXI%), 25, 2)
         Call REPLA(TTYY$, MKI$(FEXI%), 27, 2)
         'Call REPLA(TTYY$, Text5.TEXT, 29, 8)
         USUA$ = USERNAME$ 'TRIM$(Text6.TEXT): If USUA$ = "" Then USUA$ = USERNAME$
         Call REPLA(TTYY$, USUA$, 37, 16)
         'Call REPLA(TTYY$, Text4(0).TEXT, 53, 48)
         '
15       KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
              RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Else
              URN& = URN& + 1
              RENOTA& = URN&
              If TRIM$(NHPF$) = "" Then
                If KK& > 1 Then
                  GoTo 20
                End If
              End If
         End If
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         NOPI% = CVI(Mid$(TTXX$, 3, 2))
         If NOPI% > 0 Then
            If NOPI% <> NOPE% Then
               GoTo 15
            End If
         End If
         Call GRAREG("DATADOPS", TTYY$, RENOTA&)
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
            If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
               Call GRAREG("DATADOPS", MKI$(CI%) + MKI$(0) + Space$(124), RENOTA&)
            End If
         Wend
         '
         Call CIERRARCH("DATADOPS")
      End If
   End If
   '
End Sub

Private Sub Option1_Click()
   MODIFIC%(0) = 1
End Sub

Private Sub Option2_Click()
   MODIFIC(0) = 1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      '
      If MODOEDIT% = 1 Then
        If PreviousTab = 1 Then
          SSTab1.Tab = PreviousTab
          Call MENSAEDIT
        End If
        Exit Sub
      End If
      '
      Call GRADATITEM
      Call FRESHALL
      Call CARDATITEM
      '
   End If
End Sub

Private Sub Text1_Change()
   '
   For Index = 0 To 15: YACAR%(Index) = 0: Next Index
   '
   'If IUCOD% < 128 Then
   '  If TRIM$(Text1) <> "" Then
   '    If Text1 <> UCOD$(IUCOD%) Then
   '      IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = COD$
   '                           ICOMPO%(IUCOD%) = List1.ListIndex
   '    End If
   '  End If
   'End If
   '
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
End Sub

Private Sub Text10_Change()
   Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text10.TEXT))))
   MODIFIC%(0) = 1
End Sub

Private Sub Text10_DblClick()
   Call Command14_Click
End Sub

Private Sub Text11_Change()
   If List3.ListIndex >= 0 Then
     Mid$(RESECO$(List3.ListIndex), 129, 16) = AJUSTI$(Text11.TEXT, 16)
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11.TEXT)
End Sub

Private Sub Text12_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text15_Change(Index As Integer)
   If Val(Text15(Index).TEXT) > 999 Then Text15(Index).TEXT = "999"
   If Val(Text15(Index).TEXT) < 0 Then Text15(Index).TEXT = "  0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index).TEXT)
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text15(Index).TEXT)), 3, 3, 0, 2))
   If TTTT$ = "" Then TTTT$ = "0"
   While Len(TTTT$) < 3: TTTT$ = " " + TTTT$: Wend
   Text15(Index).TEXT = TTTT$
End Sub

Private Sub Text17_Change(Index As Integer)
   If Val(Text17(Index).TEXT) > 59 Then Text17(Index).TEXT = "59"
   If Val(Text17(Index).TEXT) < 0 Then Text17(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text17_GotFocus(Index As Integer)
   Text17(Index).SelStart = 0
   Text17(Index).SelLength = Len(Text17(Index).TEXT)
End Sub

Private Sub Text17_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text17(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text17(Index).TEXT = TTTT$
End Sub

Private Sub Text18_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text18_LostFocus()
   Text18.TEXT = CSTRING$(MKS$(Val(Text18.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
   If Val(Text20.TEXT) < Val(Text19.TEXT) Then
     Text20.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text2_Change()
   On Error Resume Next
   DESOPE$(List3.ListIndex) = Text2
   On Error GoTo 0
   MODIFIC%(2) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_LostFocus()
   Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_LostFocus()
   Text24.TEXT = CSTRING$(MKS$(Val(Text24.TEXT)), 3, 9, 5, 2)
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_DblClick()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Text26_Change(Index As Integer)
   If Val(Text26(Index).TEXT) > 59 Then Text26(Index).TEXT = "59"
   If Val(Text26(Index).TEXT) < 0 Then Text26(Index).TEXT = " 0"
   If List3.ListIndex >= 0 Then
     RESECOA$ = RESECO$(List3.ListIndex)
     TPAR = Val(Text15(Index).TEXT) + Val(Text17(Index).TEXT) / 60 + Val(Text26(Index).TEXT) / 3600
     If Index = 2 Then TPAR = 60 * TPAR / Val(Text7.TEXT)
     If Index <= 2 Then
         Mid$(RESECO$(List3.ListIndex), 59 + 4 * Index, 4) = MKS$(TPAR)
       ElseIf Index = 3 Then
         Mid$(RESECO$(List3.ListIndex), 79, 4) = MKS$(TPAR)
       ElseIf Index = 4 Then
         Mid$(RESECO$(List3.ListIndex), 95, 4) = MKS$(TPAR)
     End If
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(2) = 1
   End If
End Sub

Private Sub Text26_GotFocus(Index As Integer)
   Text26(Index).SelStart = 0
   Text26(Index).SelLength = Len(Text26(Index).TEXT)
End Sub

Private Sub Text26_LostFocus(Index As Integer)
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text26(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text26(Index).TEXT = TTTT$
End Sub

Private Sub Text28_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text29_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 75, 4) = MKS(Val(Text29.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text29_GotFocus()
   Text29.SelStart = 0
   Text29.SelLength = Len(Text29.TEXT)
End Sub

Private Sub Text29_LostFocus()
   Text29.TEXT = CSTRING$(MKS$(Val(Text29.TEXT)), 3, 4, 1, 2)
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text30_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 83, 4) = MKS(Val(Text30.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text30_GotFocus()
   Text30.SelStart = 0
   Text30.SelLength = Len(Text30.TEXT)
End Sub

Private Sub Text30_LostFocus()
   Text30.TEXT = CSTRING$(MKS$(Val(Text30.TEXT)), 3, 5, 2, 2)
End Sub

Private Sub Text33_GotFocus()
   Text33.SelStart = 0
   Text33.SelLength = Len(Text33.TEXT)
End Sub

'
Private Sub Text34_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 99, 4) = MKS(Val(Text34.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text34_GotFocus()
   Text34.SelStart = 0
   Text34.SelLength = Len(Text34.TEXT)
End Sub

Private Sub Text34_LostFocus()
   Text34.TEXT = TRIM$(CSTRING$(MKS$(Val(Text34.TEXT)), 3, 5, 0, 2))
End Sub

Private Sub Text31_DblClick()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32.TEXT) = "" Then Text32.TEXT = Str$(CABAS)
     Text32.SetFocus
   End If
   '
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      CODY$ = TRIM$(Text31.TEXT)
      CIY% = CODINT%(CODY$)
      If CIY% > 0 Then
         CODE$ = CODY$: CIE% = CIY%
         Text31.TEXT = CODY$
         CODE$ = CODY$: CIE% = CIY%
         Label40.Caption = DESCRIT0$(CIE%)
         Label41.Caption = UNIMED$(CIE%)
         Text32.SetFocus
        Else
         Call MENSERR(24, "Código no Válido")
         Text31.TEXT = CODE$
      End If
   End If
   '
End Sub

Private Sub Text32_GotFocus()
   Text32.SelStart = 0
   Text32.SelLength = Len(Text32.TEXT)
End Sub

Private Sub Text33_Change()
   VALONU = Val(Text33.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text33.TEXT = "": Exit Sub
   End If
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 87, 1) = Chr$(VALONU)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
   Label51.Caption = ECOARCH$("CATOPER", Chr$(VALONU))
End Sub

Private Sub Text33_DblClick()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Text36_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text4_Change()
   Label37.Caption = Text4.TEXT
   Label30.Caption = ECOARCH$("UMEDIDA", Text4.TEXT)
   MODIFIC%(0) = 1
End Sub

Private Sub Text4_DblClick()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Text5_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 5, 48) = AJUSTI$(Text5.TEXT, 48)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(2) = 1
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5.TEXT)
End Sub

Private Sub Text5_LostFocus()
   ILISTA& = List3.ListIndex
   For UI& = 1 To List3.ListCount
     If Val(Left$(List3.List(UI& - 1), 4)) = Val(Label55) Then
       TTXT$ = List3.List(UI& - 1)
       List3.RemoveItem (UI& - 1)
       Call REPLA(TTXT$, Text5.TEXT, 10, 48)
       List3.AddItem TTXT$, UI& - 1
     End If
   Next UI&
   On Error Resume Next
   List3.ListIndex = ILISTA&
   On Error GoTo 0
End Sub

Private Sub Text6_Change()
   VALONU = Val(Text6.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text6.TEXT = "": Exit Sub
   End If
   Label29.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_DblClick()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Text7_Change()
   '
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     CABAS = Val(Text7.TEXT)
     If CABAS < 1 Then
        CABAS = 1
        Text7.TEXT = "      1"
     End If
     '
     TISTA = CABAS * CVS(Mid$(VDEF$, 67, 4))
     TSTT$ = HORMINSE$(60 * TISTA)
     Text15(2).TEXT = Left$(TSTT$, 3)
     Text17(2).TEXT = Mid$(TSTT$, 5, 2)
     Text26(2).TEXT = Mid$(TSTT$, 8, 2)
   End If
   '
   MODIFIC%(2) = 1
   '
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub Text8_Change()
   VALONU = Val(Text8.TEXT)
   If VALONU < 0 Or VALONU > 32567 Then
      Text8.TEXT = "": Exit Sub
   End If
   If AREPO$ <> "PROVED1" Then
       Label32.Caption = ECOARCH$(AREPO$, Chr$(VALONU))
     Else
       Label32.Caption = ECOARCH$(AREPO$, MKI$(Int(VALONU)))
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text8_DblClick()
   Call SELECHO(AREPO$)
   VALONUE$ = VALACT1$(AREPO$)
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Text9_Change(Index As Integer)
   If List3.ListIndex >= 0 Then
     If TRIM$(Text9(Index).TEXT) = "" Or ECOARCH$("PADMAQ", Text9(Index).TEXT) <> "" Then
        If Index = 0 Then
            Label43.Caption = ECOARCH$("PADMAQ", Text9(Index).TEXT)
            Mid$(RESECO$(List3.ListIndex), 53, 6) = AJUSTI$(Text9(Index).TEXT, 6)
          Else
            Mid$(RESECO$(List3.ListIndex), 97 + 6 * Index, 6) = AJUSTI$(Text9(Index).TEXT, 6)
        End If
     End If
   End If
   MODIFIC(2) = 1
End Sub

Private Sub Text9_DblClick(Index As Integer)
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text9(Index).TEXT = VDEVU$
   End If
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   NMSI% = NUMAS% + 1
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'
Sub COPYFORMU(CODORI%, CODEST%)
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   '
   J& = 0
   PUNTCI$ = ""
   ESP1$ = "": ESP2$ = "": ESP3$ = ""
   For I& = 1 To ULTREG&("ESTRUNUE")
      EST1$ = REGLEIDO$("ESTRUNUE", I&)
      CIV% = CVI(Left$(EST1$, 2))
      EST2$ = REGLEIDO$("ESTRUDOS", I&)
      If CIV% > 0 Then
        If CIV% <= NUMAS% Then
          If CIV% <> CODEST% Then
            J& = J& + 1
            Call GRAREG("ESTRUNUE", EST1$, J&)
            Call GRAREG("ESTRUDOS", EST2$, J&)
            PUNTCI$ = PUNTCI$ + MKI$(CIV%)
          End If
        End If
      End If
      '
      If CIV% = CODORI% Then
        If ESP1$ = "" Then
          ESP1$ = EST1$
          ESP2$ = EST2$
        End If
      End If
   Next I&
   '
   Call REPLA(ESP1$, MKI$(CODEST%), 1, 2)
   EST1$ = ESP1$
   EST2$ = ESP2$
   J& = J& + 1
   Call GRAREG("ESTRUNUE", EST1$, J&)
   Call GRAREG("ESTRUDOS", EST2$, J&)
   PUNTCI$ = PUNTCI$ + MKI$(CODEST%)
   '
   For I& = J& + 1 To ULTREG&("ESTRUNUE")
     Call GRAREG("ESTRUNUE", String$(128, 0), I&)
     Call GRAREG("ESTRUDOS", String$(128, 0), I&)
   Next I&
   '
   Call CIERRARCH("*.*")
   '
   J& = 0
   For I& = 1 To ULTREG&("ESTRULAR")
      EST3$ = REGLEIDO$("ESTRULAR", I&)
      CIV% = CVI(Left$(EST3$, 2))
      If CIV% > 0 Then
        If CIV% <= NUMAS% Then
          If CIV% <> CODEST% Then
            J& = J& + 1
            Call GRAREG("ESTRULAR", EST3$, J&)
          End If
        End If
      End If
      '
      If CIV% = CODORI% Then
        If ESP3$ = "" Then
          ESP3$ = EST3$
        End If
      End If
   Next I&
   '
   If ESP3$ <> "" Then
     Call REPLA(ESP3$, MKI$(CODEST%), 1, 2)
     EST3$ = ESP3$
     J& = J& + 1
     Call GRAREG("ESTRULAR", EST3$, J&)
   End If
   Call CIERRARCH("ESTRULAR")
   '
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
End Sub

Sub COPYRUTA(CODORI%, CODEST%)
   '
   Dim REGTX$(256)
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   RFFX$ = RECFILT$("SECOPER", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("SECOPER", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("SECOPER")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("SECOPER", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("SECOPER")
     End If
     VALANT$ = REGLEIDO$("SECOPER", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("SECOPER", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("SECOPER")
     Call GRAREG("SECOPER", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("SECOPER")
   '
   RFFX$ = RECFILT$("OBSEOPER", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("OBSEOPER", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("OBSEOPER")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("OBSEOPER", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("OBSEOPER")
     End If
     VALANT$ = REGLEIDO$("OBSEOPER", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("OBSEOPER", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("OBSEOPER")
     Call GRAREG("OBSEOPER", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("OBSEOPER")
   '
   RFFX$ = RECFILT$("PLACONTR", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PLACONTR", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PLACONTR")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PLACONTR", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PLACONTR")
     End If
     VALANT$ = REGLEIDO$("PLACONTR", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PLACONTR", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PLACONTR")
     Call GRAREG("PLACONTR", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PLACONTR")
   '
   RFFX$ = RECFILT$("PUEPUNTO", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PUEPUNTO", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PUEPUNTO")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PUEPUNTO", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PUEPUNTO")
     End If
     VALANT$ = REGLEIDO$("PUEPUNTO", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PUEPUNTO", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PUEPUNTO")
     Call GRAREG("PUEPUNTO", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PUEPUNTO")
   '
   RFFX$ = RECFILT$("DATADOPS", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("DATADOPS", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("DATADOPS")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("DATADOPS", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("DATADOPS")
     End If
     VALANT$ = REGLEIDO$("DATADOPS", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("DATADOPS", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("DATADOPS")
     Call GRAREG("DATADOPS", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("DATADOPS")
   '
End Sub
'
Sub COPYSPEC(CODORI%, CODEST%)
   '
   Dim REGTX$(1024)
   Screen.MousePointer = 11
   RFFX$ = RECFILT$("ESPECIFI", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("ESPECIFI", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("ESPECIFI")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("ESPECIFI", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("ESPECIFI")
     End If
     VALANT$ = REGLEIDO$("ESPECIFI", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("ESPECIFI", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("ESPECIFI")
     Call GRAREG("ESPECIFI", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("ESPECIFI")
   Screen.MousePointer = 1
   '
End Sub
'
Sub COPYPLANOS(CODORI%, CODEST%)
   REBLA$ = REGBLAN$("INDIPLAN")
   If CODORI% > 0 Then
     If CODEST% > 0 Then
       If CODORI% <> CODEST% Then
         XYZ$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CODORI%))
         If XYZ$ <> REBLA$ Then
           Call REPLA(XYZ$, MKI$(CODEST%), 1, 2)
           Call FILTERSETRECORD("INDIPLAN", 1, MKI$(CODEST%), XYZ$)
           Call CIERRARCH("INDIPLAN")
         End If
       End If
     End If
   End If
End Sub

Sub REVIMASTER()                         ' revision del maestro
        '
        Screen.MousePointer = 11
        HAYINCO% = 0
        '
        Call BLOQARCH("MASTER")
        Call BLOQARCH("INVERT")
        Call BLOQARCH("INVENUE")
        '
        Dim CODISEL$(32767), CISEL%(32767)
        '
        If Left$(UCase$(EMPREAC$), 5) = "BOREN" Then
           ' ejecuta conversion de codigos
           FIN& = NUMAS%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("MASTER", I& + 1)
              CODES$ = Mid$(X$, 77, 16)
              If Mid$(CODES$, 16, 1) <> " " Then
                If Mid$(CODES$, 16, 1) <> Chr$(96) Then
                  Call COMUNI("Codigo '" + CODES$ + "' No Válido\(excede longitud = 15 car.)\  \Pulse 'Aceptar' para Marcarlo de Baja.")
                  Call REPLA(CODES$, Chr$(96), 16, 1)
                  Call REPLA(X$, Chr$(96), 92, 1)
                  Call GRAREG("MASTER", X$, I& + 1)
                End If
              End If
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 77, 16)
                 Call GRAREG("MASTER", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("INVERT", I& + 1)
              CODES$ = Mid$(X$, 1, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 1, 16)
                 Call GRAREG("INVERT", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("INVERT")
           '
        End If
        '
        NM% = ULTREG&("MASTER") - 1
        If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
              X$ = REGLEIDO$("MASTER", 1&)
              Call REPLA(X$, MKI$(NM%), 1, 2)
              Call GRAREG("MASTER", X$, 1&)
              If ATENTI% <> (-123) Then Call COMUNI("Número de Registros\de Archivo Maestro:\  \Indice Restaurado.")
              HAYINCO% = 1
          Else
              If ATENTI% <> (-123) Then Call COMUNI("Número de Registros\de Archivo Maestro:\  \Control O.K.")
        End If
        '
        Screen.MousePointer = 11
        NVL& = ULTREG&("INVERT") - 1
        FIN& = NVL&
        For II& = 1 To NVL&
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("INVERT", II& + 1)
            COD$ = Left$(X$, 16): CII% = CVI(Mid$(X$, 17, 2))
            If Mid$(COD$, 16, 1) = Chr$(96) Then
               If CII% > 0 Then
                  If CII% <= NUMAS% Then
                     REGMA& = CII% + 1
                     Y$ = REGLEIDO$("MASTER", REGMA&)
                     If Mid$(Y$, 77, 15) = Left$(COD$, 15) Then
                        Call REPLA(Y$, Chr$(96), 92, 1)
                        Call GRAREG("MASTER", Y$, REGMA&)
                     End If
                  End If
               End If
            End If
        Next II&
        '
        Screen.MousePointer = 11
        FORMALTER.LINOSORT.Clear
        FORMALTER.LISORT.Clear
        FIN& = NM%
        RECUMAX& = 0
        For II& = 1 To NM%
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("MASTER", II& + 1)
            If CVI(Left$(X$, 2)) <> II& Then
                If ATENTI% <> (-123) Then Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                Screen.MousePointer = 11
                cix% = II&
                Call REPLA(X$, MKI$(cix%), 1, 2)
                Call GRAREG("MASTER", X$, II& + 1)
            End If
            If Mid$(X$, 92, 1) <> " " Then
              If Mid$(X$, 92, 1) <> Chr$(96) Then
                CODES$ = TRIM$(Mid$(X$, 77, 16))
                Call COMUNI("Codigo '" + CODES$ + "' No Válido\(excede longitud = 15 car.)\  \Pulse 'Aceptar' para Marcarlo de Baja.")
                Call REPLA(X$, Chr$(96), 92, 1)
                Call GRAREG("MASTER", X$, I& + 1)
                HAYINCO% = 1
              End If
            End If
            '
            COD$ = AJUSTI$(TRIM$(Mid$(X$, 77, 15)), 15) + Mid$(X$, 92, 1)
            If Left$(COD$, 4) = "zzz-" Then
              If Val(Mid$(COD$, 5, 6)) > RECUMAX& Then
                RECUMAX& = Val(Mid$(COD$, 5, 6))
              End If
            End If
            '
            ZZ$ = "A" + COD$ + CSTRING$(MKS$(II&), 3, 7, 0, 2)
            FORMALTER.LISORT.AddItem ZZ$
        Next II&
        '
        FIN& = FORMALTER.LISORT.ListCount
        For II& = 1 To FORMALTER.LISORT.ListCount
           Call TRACE(20, II&, FIN&)
           ZZ$ = FORMALTER.LISORT.List(II& - 1)
           Call AGRELISCOD(ZZ$)
        Next II&
        '
        NVL& = ULTREG&("INVERT") - 1
        If NVL& < 1 Then
             Call REGAPP("INVERT", String$(16, 0) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = FORMALTER.LINOSORT.ListCount
        X$ = REGBLAN$("INVERT")
        Call REPLA(X$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", X$, 1)
        '
        J& = 1
        FIN& = NV% - 1
        For II& = 0 To FORMALTER.LINOSORT.ListCount - 1
            Call TRACE(20, II&, FIN&)
            FORMALTER.LINOSORT.ListIndex = II&
            CODISEL$(II&) = Mid$(FORMALTER.LINOSORT.TEXT, 2, 16)
            CISEL%(II&) = Val(Mid$(FORMALTER.LINOSORT.TEXT, 18))
            X$ = REGBLAN$("INVERT")
            Call REPLA(X$, CODISEL$(II&), 1, 16)
            Call REPLA(X$, MKI$(CISEL%(II&)), 17, 2)
            Call GRAREG("INVERT", X$, II& + 2)
        Next II&
        Call SETULTREG("INVERT", 1& + NV%)
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        CODA$ = ""
        FIN& = NV%
        For II& = 1 To NV%
           Call TRACE(20, II&, FIN&)
           IK$ = REGLEIDO$("INVERT", II& + 1)
           CODL$ = Left$(IK$, 15)
           If TRIM$(CODL$) = "" Or CODL$ = CODA$ Then
                'If TRIM$(CODL$) = "" Then GoTo 19
                REPLAI% = REPLAI% + 1
                'CODNUE$ = TRIM$(CODL$) + "-" + Chr$(64 + REPLAI%)
                CODNUE$ = "zzz-" + TRIM$(Str$(RECUMAX& + 1))
                RECUMAX& = RECUMAX& + 1
                If Left$(CODNUE$, 1) = "-" Then CODNUE$ = "0" + CODNUE$
                RGM& = 1 + CVI(Mid$(IK$, 17, 2))
                M$ = REGLEIDO$("MASTER", RGM&)
                Call REPLA(M$, CODNUE$, 77, 16)
                Call GRAREG("MASTER", M$, RGM&)
                Call REPLA(IK$, CODNUE$, 1, 15)
                Call REPLA(IK$, Chr$(96), 16, 1)
                Call GRAREG("INVERT", IK$, II& + 1)
                If ATENTI% <> (-123) Then Call COMUNI("Codigo '" + TRIM$(CODL$) + "' Duplicado o no Válido.\  \Reemplazado por '" + TRIM$(CODNUE$) + "'.")
                Screen.MousePointer = 11
                HAYINCO% = 1
              Else
                REPLAI% = 0
           End If
           '
           If REPLAI% = 0 Then CODA$ = CODL$
19      Next II&
        '
        Call CIERRARCH("*.*")
        Call BLANARCH("INVENUE")
        '
17000   'Call BLIMESS("Actualizando Archivos de Validación ...")
        '
        'Call BLOQARCH("ECOMAST")
        FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 64)
        Dim ECM As String * 64
        '
        'Call BLOQARCH("CIACTI")
        FICIA% = FILEOPEN%(LUDAT$ + "CIACTI.DAT", 0, 2)
        Dim CCT As String * 2
        '
        NV% = ULTREG&("INVERT") - 1
        CCT$ = MKI$(NV%)
        Put #FICIA%, 1, CCT$
        '
        K% = 0
        FIN& = NV%
        For XL& = 1 To NV%
             Call TRACE(20, XL&, FIN&)
             IK$ = REGLEIDO$("INVERT", XL& + 1)
             If Mid$(IK$, 16, 1) <> Chr$(96) Then
               If TRIM$(Left$(IK$, 16)) <> "" Then
                 '
                 CI% = CVI(Mid$(IK$, 17, 2))
                 DE$ = DESCRIT$(CI%)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, Left$(IK$, 16), 1, 16)
                 Call REPLA(EC$, DE$, 17, 46)
                 '
                 K% = K% + 1: ECM$ = EC$: Put #FIECO%, K%, ECM$
                              CCT$ = MKI$(CI%): Put #FICIA%, K% + 1, CCT$
                 '
               End If
             End If
        Next XL&
        '
        CCT$ = MKI$(K%): Put #FICIA%, 1, CCT$
        '
        For K1% = K% + 1 To LOF(FIECO%) / 64
           ECM$ = String$(64, 0): Put #FIECO%, K1%, ECM$
           CCT$ = String$(2, 0): Put #FICIA%, K1% + 1, CCT$
        Next K1%
        '
        Close #FIECO%: FIECO% = 0 'Call CIERRARCH("ECOMAST")
        Close #FICIA%: FICIA% = 0 'Call CIERRARCH("CIACTI")
        FIECO% = 0: FICIA% = 0
        '
        Call GRACONTROL("", "AMASTO", "NORMAL")
        '
        Call CIERRARCH("*.*")
        Call LIBARCH("MASTER")
        '
        Call FRESHECHO("MASTER")
        Call FRESHECHO("ECOMAST")
        '
        'Call BLIMESS("-")
        '
        Screen.MousePointer = 1
        If HAYINCO% = 0 Then
            If ATENTI% <> (-123) Then Call COMUNI("No se han Detectado Inconsistencias\en la Revisión del Maestro de Items.")
        End If
        Exit Sub
        '
        '
12500   If ULTREG&("INVERT") < 1 Then
             Call REGAPP("INVERT", String$(15, 0) + Chr$(1) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = ULTREG&("INVERT") - 1
        If CVI(Mid$(REGLEIDO$("INVERT", 1&), 17, 2)) <> NV% Then
            Call COMUNI("Posible Inconsistencia (1) !!!\Ejecute Opción de Control.")
            Call FINAL("")    ' GoTo 2000
        End If
        '
        ABRE$ = AJUSTI$(COD$, 15)
        LI& = 1: LS& = NV% + 2
        ENCU% = 0
        '
12510   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 12580
        L& = LI& + E&: If L& < 2 Or L& > NV% + 1 Then GoTo 12580
        IK$ = REGLEIDO$("INVERT", L&): I1$ = Left$(IK$, 15)
                                  If I1$ < ABRE$ Then LI& = L&: GoTo 12510
                                  If I1$ > ABRE$ Then LS& = L&: GoTo 12510
        ENCU% = 1: CI% = CVI(Mid$(IK$, 17, 2)): RGI& = L&
                If Mid$(IK$, 16, 1) = Chr$(96) Then
                    ENCU% = (-1)
                End If
        '
12580   Return
        '
        '
End Sub

Sub AGRELISCOD(RAGRE$)
   '
   ULIS& = FORMALTER.LINOSORT.ListCount
   If ULIS& < 1 Or Left$(RAGRE$, 17) > FORMALTER.LINOSORT.List(ULIS& - 1) Then
       FORMALTER.LINOSORT.AddItem RAGRE$
       GoTo 10
     Else
       For J& = ULIS& - 1 To 0 Step -1
         If Left$(FORMALTER.LINOSORT.List(J&), 17) <= Left$(RAGRE$, 17) Then
           FORMALTER.LINOSORT.AddItem RAGRE$, J& + 1
           GoTo 10
         End If
       Next J&
       FORMALTER.LINOSORT.AddItem RAGRE$, 0
10 End If
   '
End Sub

Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta Definir Depósitos")
        'OPMAST00.Show 1
        Call CONECRUN("FLSETUP/OPMAST00", "Configuración de Funcionamiento")
        GoTo 10
    End If
    '
    If ULTREG&("UMEDIDA") < 2 Then
        Call BLANARCH("UMEDIDA")
        Call REGAPP("UMEDIDA", "U.Unidades")
        Call REGAPP("UMEDIDA", "KgKilogramos")
        Call REGAPP("UMEDIDA", "grGramos")
        Call REGAPP("UMEDIDA", "gKGramos - Kilos")
        Call REGAPP("UMEDIDA", "M.Metros")
        Call REGAPP("UMEDIDA", "M2M.Cuadrados")
        Call REGAPP("UMEDIDA", "M3M.Cúbicos")
        Call REGAPP("UMEDIDA", "mmMilímetros")
        Call REGAPP("UMEDIDA", "cmCentímetros")
        Call REGAPP("UMEDIDA", "LtLitros")
        Call CIERRARCH("UMEDIDA")
    End If
    XX$ = REGLEIDO$("UMEDIDA", 1)
    '
    If ULTREG&("UNIMED") < 2 Then
        Call BLANARCH("UNIMED")
        Call REGAPP("UNIMED", "U.  Unidades" + String$(256, 0))
        Call REGAPP("UNIMED", "Kg  Kilogramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gr  Gramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gK  Gramos - Kilos" + String$(256, 0))
        Call REGAPP("UNIMED", "M.  Metros" + String$(256, 0))
        Call REGAPP("UNIMED", "M2  M.Cuadrados" + String$(256, 0))
        Call REGAPP("UNIMED", "M3  M.Cúbicos" + String$(256, 0))
        Call REGAPP("UNIMED", "mm  Milímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "cm  Centímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "Lt  Litros" + String$(256, 0))
        Call CIERRARCH("UNIMED")
    End If
    XX$ = REGLEIDO$("UNIMED", 1)
    '
    If ULTREG&("TATIMAT") < 1 Then
        Call BLANARCH("TATIMAT")
        Call REGAPP("TATIMAT", Chr$(1) + "Materia Prima")
        Call REGAPP("TATIMAT", Chr$(2) + "Pieza Comercial")
        Call REGAPP("TATIMAT", Chr$(3) + "Pieza Elaborada")
        Call REGAPP("TATIMAT", Chr$(4) + "Semielaborado")
        Call REGAPP("TATIMAT", Chr$(5) + "Prod.Terminado")
        Call REGAPP("TATIMAT", Chr$(8) + "Servicios")
        Call REGAPP("TATIMAT", Chr$(9) + "Improductivo")
        Call CIERRARCH("TATIMAT")
    End If
    XX$ = REGLEIDO$("TATIMAT", 1)
    '
    If ULTREG&("TABTIMA") < 1 Then
       DPPRE% = Asc(Left$(REGLEIDO$("ECODEP", 1), 1))
       For U& = 1 To ULTREG&("TATIMAT")
          X$ = REGLEIDO$("TATIMAT", U&)
          Y$ = REGBLAN$("TABTIMA")
          Call REPLA(Y$, X$, 1, 16)
          Call REPLA(Y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", Y$)
       Next U&
       Call CIERRARCH("TABTIMA")
    End If
    XX$ = REGLEIDO$("TABTIMA", 1)
    '
    If ULTREG&("TACRIRE") < 4 Then
        Call BLANARCH("TACRIRE")
        Call REGAPP("TACRIRE", Chr$(1) + "Por Pedido")
        Call REGAPP("TACRIRE", Chr$(2) + "Por Mínimos")
        Call REGAPP("TACRIRE", Chr$(3) + "Cons.Estadístico")
        Call REGAPP("TACRIRE", Chr$(4) + "Programación")
        Call CIERRARCH("TACRIRE")
    End If
    XX$ = REGLEIDO$("TACRIRE", 1)
    '
    If ULTREG&("TAMONED") < 2 Then
        Call BLANARCH("TAMONED")
        Call REGAPP("TAMONED", Chr$(1) + AJUSTI$("Pesos", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(2) + AJUSTI$("Dolares U$S", 27) + MKS$(1))
        Call CIERRARCH("TAMONED")
    End If
    XX$ = REGLEIDO$("TAMONED", 1)
    '
    If ULTREG&("TACOMPO") < 2 Then
        Call BLANARCH("TACOMPO")
        Call REGAPP("TACOMPO", Chr$(0) + "Manual")
        Call REGAPP("TACOMPO", Chr$(1) + "Automática")
        Call CIERRARCH("TACOMPO")
    End If
    XX$ = REGLEIDO$("TACOMPO", 1)
    '
    If ULTREG&("GRUPIVA") < 1 Then
        Call SETULTREG("GRUPIVA", 0)
        Call REGAPP("GRUPIVA", Chr$(0) + "IVA 21 %   " + MKS$(21))
        Call REGAPP("GRUPIVA", Chr$(1) + "IVA 10,5 % " + MKS$(10.5))
        Call REGAPP("GRUPIVA", Chr$(2) + "IVA 27 %   " + MKS$(27))
        Call CIERRARCH("GRUPIVA")
    End If
    XX$ = REGLEIDO$("GRUPIVA", 1)
    '
    Screen.MousePointer = 1
    '
End Sub
'


Sub COPYSTRUC()
   '
10 CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     GoTo 99
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   Call LEEEXPLO(CODOR$, 1)
   If CAIT% < 1 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("Copia Imposible.\Código Origen " + TRIM$(CODOR$) + " sin Estructura.\\Repetir\Abandonar")
     If OPX% = 1 Then GoTo 10
     GoTo 99
   End If
   '
   Call LEEEXPLO(CODES$, 1)
   If CAIT% > 0 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("ATENCION: Código Destino " + TRIM$(CODES$) + " con Estructura.\\Cancelar\Sobre-escribir")
     If OPX% <> 2 Then GoTo 10
   End If
   '
   Call COPYFORMU(CODORI%, CODEST%)
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   '
99 End Sub

Sub COPYRUTAS()
   '
Dim REG&(256), REGTX$(256)
   '
10 Screen.MousePointer = 1
   CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     Call CIERRARCH("*.*")
     Exit Sub
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Call BLOQARCH("SECOPER")
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   RFFX$ = RECFILT$("SECOPER", 1, MKI$(CODORI%))
   If Len(RFFX$) < 4 Then
     Screen.MousePointer = 1
     Call COMUNI("Copia Imposible.\Código Origen sin Operaciones.")
     GoTo 10
   End If
   '
   RFFY$ = RECFILT$("SECOPER", 1, MKI$(CODEST%))
   If Len(RFFY$) >= 4 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("ATENCION: Código Destino con Operaciones !!!\\Cancelar\Sobre-escribir")
     If OPX% <> 2 Then GoTo 10
   End If
   '
   Call COPYRUTA(CODORI%, CODEST%)
   '
   Call CIERRARCH("SECOPER")
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'
Sub LISTRUTAS()
   '
   XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Dim PUNCOD$(32767)
   '
   CI1% = CVI(Left$(X$, 2))
   cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2))
   cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   V& = 0
   FIN& = ULTREG&("SECOPER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("SECOPER", U&)
     CI0% = CVI(Left$(X$, 2))
     If CI0% > 0 Then
       If CI0% <= NUMAS% Then
         PUNCOD$(CI0%) = PUNCOD$(CI0%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   FIN& = NUINV%
   For IU& = 1 To NUINV%
     Call TRACE(20, IU&, FIN&)
     Z$ = REGLEIDO$("INVERT", IU& + 1)
     CODEX$ = Left$(Z$, 16)
     If Mid$(Z$, 16, 1) <> "`" Then
       CODEX$ = TRIM$(CODEX$)
       If CODEX$ >= cod1$ Then
         If CODEX$ <= cod2$ Then
           CI0% = CVI(Mid$(Z$, 17, 2))
           If Len(PUNCOD$(CI0%)) >= 4 Then
              '
              LOSTAND = LOTESTAN(CI0%)
              DESCRI$ = TRIM$("Lt.Std." + TRIM$(Str$(LOSTAND)) + " " + TRIM$(UNIMED$(CI0%)))
              DESCRI$ = " (" + DESCRI$ + ")"
                LOMADE% = 60 - Len(DESCRI$)
                If LOMADE% > 0 Then
                  DESCRI$ = Left$(TRIM$(DESCRIT0$(CI0%)), LOMADE%) + DESCRI$
                End If
                '
              Y$ = REGBLAN$("SECOPSEL")
              Call REPLA(Y$, MKI$(CI0%), 1, 2)
              Call REPLA(Y$, DESCRI$, 3, 60)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              Call REPLA(Y$, Space$(60), 3, 60)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              For IV% = 1 To Len(PUNCOD$(CI0%)) Step 4
                U& = CVS(Mid$(PUNCOD$(CI0%), IV%, 4))
                If U& > 0 Then
                  If U& <= ULTREG&("SECOPER") Then
                    Y1$ = REGLEIDO$("SECOPER", U&)
                    NOPES$ = TRIM$(CSTRING$(Mid$(Y1$, 3, 2), 1, 6, 0, 2))
                    While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
                    DOPES$ = Mid$(Y1$, 5, 62)
                    TVAR = CVS(Mid$(Y1$, 67, 4)) * LOSTAND
                    Y$ = Y1$
                    Call REPLA(Y$, NOPES$, 3, 6)
                    Call REPLA(Y$, DOPES$, 9, 62)
                    Call REPLA(Y$, MKS$(TVAR), 71, 4)
                    V& = V& + 1
                    Call GRAREG("SECOPSEL", Y$, V&)
                  End If
                End If
              Next IV%
              '
              Call REPLA(Y$, String$(126, 0), 3, 126)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              Y$ = REGBLAN$("SECOPSEL")
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
           End If
         End If
       End If
     End If
   Next IU&
   '
   Call SETULTREG("SECOPSEL", V&)
   Call CIERRARCH("SECOPER")
   Call CIERRARCH("SECOPSEL")
   Screen.MousePointer = 1
   Call FINAL("*LISECOP")
   '
End Sub

Sub REDESPIE()
    '
    CAERR% = 0
    FIN& = ULTREG&("ESTRUNUE")
    FINDOS& = ULTREG&("ESTRUDOS")
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      R0$ = REGLEIDO$("ESTRUNUE", I&)
      R1$ = Left$(R0$, 2)
      DPS$ = Mid$(R0$, 3, 120)
      If I& <= FINDOS& Then
        DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", I&)
        CI1% = CVI(R1$)
        If CI1% > 0 Then
          If CI1% <= NUMAS% Then
            A$ = CODEXT$(CI1%) + " " + DESCRIT$(CI1%)
            '
            J% = 0
            For K% = 1 To Len(DPS$) Step 8
              DPSX$ = Mid$(DPS$, K%, 8)
              CI2% = CVI(Left$(DPSX$, 2))
              If CI2% > 0 Then
                If CI2% <= NUMAS% Then
                  J% = J% + 1
                  CII1%(J%) = CI2%
                End If
              End If
            Next K%

            For K% = 1 To J%
              If CII1%(K%) = CI1% Then
                MENSA$ = "Recursividad en Código " + TRIM$(CODEXT$(CI1%)) + "\" + DESCRIT0$(CI1%)
                If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
              End If
              '
              For K1% = K% + 1 To J%
                If CII1%(K1%) = CII1%(K%) Then
                  CAERR% = CAERR% + 1
                  cod1$ = TRIM$(CODEXT$(CI1%))
                  cod2$ = TRIM$(CODEXT$(CII1%(K%)))
                  If CODINT%(cod1$) = CI1% Then ' SUPRIME ITEMS DE BAJA
                    MENSA$ = "Repetición del Código " + cod2$ + "\en Fórmula de Codigo " + cod1$
                    If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
                  End If
                  GoTo 89
                End If
              Next K1%
            Next K%
          End If
        End If
      End If
89  Next I&
    '
    If CAERR% < 1 Then
        Call COMUNI("Control O.K. - No se Detectaron Repeticiones")
      Else
        Call MENSERR(20, "Se Encontraron" + Str$(CAERR%) + " Repeticiones en Despieces")
    End If
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub


Sub LIESTRU()      ' listado estructuras y enlaces por codigo
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    FIECO% = FILENBR%("LIESTRU")
    Dim ECM As String * 16
    '
    Call HEADERS("LIESTRU", "")
    Call HEADERS("LIESTRU", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LIESTRU", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
    '
    Screen.MousePointer = 11
    NESTRU& = ULTREG&("ESTRUNUE")
    J& = 0
    '
    For I& = 1 To NESTRU&
      Call TRACE(20, I&, NESTRU&)
      CI% = CVI(Left$(REGLEIDO$("ESTRUNUE", I&), 2))
      ACTIX% = 0
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          CODYY$ = CODEXT$(CI%)
          If Mid$(CODYY$, 16, 1) <> Chr$(96) Then ' SUPRIME ITEMS DE BAJA
            If CODYY$ >= cod1$ Then
              If CODYY$ <= cod2$ Then
                Call LEEEXPLO(CODYY$, 1)
                If CAIT% > 0 Then
                  For K% = 1 To CAIT%
                    CI1% = CIJ%(K%)
                    CAN = USOI(K%)
                    PD% = PDI%(K%)
                    If CI1% > 0 Then
                      If CI1% <= NUMAS% Then
                        If Abs(CAN) > 0.0005 Then
                          UM$ = UNIMED$(CI1%)
                          Z$ = REGBLAN$("LIESTRU")
                          Call REPLA(Z$, MKI$(CI1%), 3, 2)
                          Call REPLA(Z$, UM$, 5, 2)
                          Call REPLA(Z$, MKI$(CI%), 1, 2)
                          Call REPLA(Z$, MKS$(CAN), 7, 4)
                          Call REPLA(Z$, Chr$(PD%), 11, 1)
                          '
                          J& = J& + 1
                          Call GRAREG("LIESTRU", Z$, J&)
                          ACTIX% = 1
                          '
                        End If
                      End If
                    End If
                  Next K%
                End If
              End If
            End If
          End If
        End If
      End If
      '
      If ACTIX% > 0 Then
        Z$ = REGBLAN$("LIESTRU")
        Call REPLA(Z$, MKI$(CI%), 1, 2)
        J& = J& + 1
        Call GRAREG("LIESTRU", Z$, J&)
      End If
      '
    Next I&
    '
    Call SETULTREG("LIESTRU", J&)
    Call CIERRARCH("LIESTRU")
    '
    Screen.MousePointer = 1
    Call FINAL("*SOLIEST")
    '
End Sub
'

Sub LIMULTI()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        Call TRACE(20, REG&, FIN&)
        X$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(X$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(X$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
'    FORMULA.List2.Clear
    '
    FINV& = NUINV%
    For REG& = 1 To NUINV%
        '
        Call TRACE(20, REG&, FINV&)
        X$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(X$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(X$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOMUL(CI%)
                If CONTAMUL% > 1 Then
                  '
                  If REGQ& > 0 Then
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIMULTI", String$(128, "="), REGQ&)
                  End If
                  '
                  For I% = 1 To CONTAMUL%
                    '
                    If I% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If CANT(I%) > 0.00005 Then  ' SUPRIME DESARMADO
                       XX$ = REGBLAN$("LIMULTI")
                       Call REPLA(XX$, MKI$(I% - 1), 1, 2)
                       Call REPLA(XX$, CODEXT$(CIK%(I%)), 3, 16)
                       Call REPLA(XX$, DESCRIT$(CIK%(I%)), 17 + 2 * NIVE%(I%), 48)
                       Call REPLA(XX$, CSTRING$(MKS$(CANT(I%)), 4, 10, 5, 2), 75, 10)
                       Call REPLA(XX$, CSTRING$(MKI$(PAD%(I%)), 1, 6, 0, 2), 85, 6)
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If I% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                 Next I%
                  '
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    Call SETULTREG("LIMULTI", REGQ&)
    Call CIERRARCH("LIMULTI")
    '
    Call HEADERS("LIMULTI", "")
    Screen.MousePointer = 1
    Call FINAL("*LIMULTI")
    '
End Sub
'

Sub LIMULIN()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        Call TRACE(20, REG&, FIN&)
        X$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(X$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(X$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
    '
    J& = 0
    For REG& = 1 To NUINV%
        '
        X$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(X$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(X$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOCOM(CI%, 1)
                If CONTAMUL% > 1 Then
                  '
                  For K% = 1 To CONTAMUL%
                    CI1% = CIK%(K%)
                    CAN = CANT(K%)
                    PD% = 0
                    If CI1% > 0 Then
                      If CI1% <= NUMAS% Then
                        If Abs(CAN) > 0.0005 Then
                          UM$ = UNIMED$(CI1%)
                          Z$ = REGBLAN$("LIESTRU")
                          Call REPLA(Z$, MKI$(CI1%), 3, 2)
                          Call REPLA(Z$, UM$, 5, 2)
                          Call REPLA(Z$, MKI$(CI%), 1, 2)
                          Call REPLA(Z$, MKS$(CAN), 7, 4)
                          Call REPLA(Z$, Chr$(PD%), 11, 1)
                          '
                          J& = J& + 1
                          Call GRAREG("LIESTRU", Z$, J&)
                          ACTIX% = 1
                          '
                        End If
                      End If
                    End If
                  Next K%
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    Call SETULTREG("LIESTRU", J&)
    Call CIERRARCH("LIESTRU")
    '
    Call HEADERS("LIESTRU", "")
    Screen.MousePointer = 1
    Call FINAL("*LIESTRU")
    '
End Sub
'
Private Sub Text9_GotFocus(Index As Integer)
   Text9(Index).SelStart = 0
   Text9(Index).SelLength = Len(Text9(Index).TEXT)
End Sub

Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   Command26.SetFocus
End Sub

Sub LIMPLOSI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
20 Screen.MousePointer = 11
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Implosión " + DESCRIT$(CI%)
   '
   UREIMP& = ULTREG&("IMPLOSI")
   RFF$ = RECFILT$("IMPLOSI", 2, MKI$(CI%))
   For IK& = 1 To Len(RFF$) Step 4
     RIMPLO& = CVS(Mid$(RFF$, IK&, 4))
     If RIMPLO& > 0 Then
       If RIMPLO& <= UREIMP& Then
         IMPLOX$ = REGLEIDO$("IMPLOSI", RIMPLO&)
         CICONJU% = CVI(Left$(IMPLOX$, 2))
         If CICONJU% > 0 Then
           If CICONJU% <= NUMAS% Then
             Z$ = Space$(64)
             Call REPLA(Z$, CODEXT$(CICONJU%), 1, 16)
             Call REPLA(Z$, DESCRIT0$(CICONJU%), 17, 32)
             Call REPLA(Z$, UNIMED$(CICONJU%), 51, 2)
             Call REPLA(Z$, CSTRING$(Mid$(IMPLOX$, 5, 4), 3, 10, 4, 2), 55, 10)
             IMPLOSI.List1.AddItem Z$
             On Error Resume Next
             IMPLOSI.List1.TopIndex = IMPLOSI.List1.ListCount - 10
             On Error GoTo 0
             DoEvents
           End If
         End If
       End If
     End If
   Next IK&
   '
99 Screen.MousePointer = 1
   '
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI.Show 1
   '
End Sub
'
Sub COMULTI()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   CICON% = CI%
   '
   Screen.MousePointer = 11
   Call EXPLOMUL(CI%)
   If CONTAMUL% < 2 Then Exit Sub
   '
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   '
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
     Call REPLA(Z$, SP$ + DESCRIT0$(CIK%(KKI%)), 17, 32)
     Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 51, 2)
     Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 4, 2), 55, 10)
     IMPLOSI.List1.AddItem Z$
   Next KKI%
   Screen.MousePointer = 1
   '
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI.Show 1
   '
End Sub

Sub GENECOMASNUE()
   '
   'Static YAGENEC%
   'If YAGENEC% > 0 Then Exit Sub
   'If HAYACMAS% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   '
   FIN& = ULTREG&("MASTER")
   If FIN& < 2 Then
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   JI& = 1  ' PARA INVERT
   REBLI$ = Space$(16) + MKI$(0)
   JE& = 0  ' PARA ECOMAST
   REBLE$ = Space$(62) + MKI$(0)
   '
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, String$(16, 0), 1, 16)
   Call GRAREG("INVERT", XI$, 1)
   '
   For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     XM$ = REGLEIDO$("MASTER", U&)
     CII% = U& - 1
     If CVI(Left$(XM$, 2)) <> CII% Then
       Call REPLA(XM$, MKI$(CII%), 1, 2)
       Call GRAREG("MASTER", XM$, U&)
     End If
     '
     CODE$ = Mid$(XM$, 77, 16)
     DEDE$ = Mid$(XM$, 5, 64)
     '
     XI$ = REBLI$
     Call REPLA(XI$, CODE$, 1, 16)
     Call REPLA(XI$, MKI$(CII%), 17, 2)
     JI& = JI& + 1
     Call GRAREG("INVERT", XI$, JI&)
     '
     If Mid$(CODE$, 16, 1) <> Chr$(96) Then
       XE$ = REBLE$
       Call REPLA(XE$, CODE$, 1, 16)
       '
       RFF$ = RECFILT$("INDIPLAN", 1, MKI$(CII%))
       NUMREG& = CVS(RFF$)
       XX$ = REGLEIDO$("INDIPLAN", NUMREG&)
       PLANO$ = TRIM$(Mid$(XX$, 3, 16))
       If PLANO$ <> "" Then
         If PLANO$ <> TRIM$(CODE$) Then
           While Len(DEDE$ + PLANO$) > 43
             DEDE$ = Left$(DEDE$, Len(DEDE$) - 1)
           Wend
           DEDE$ = DEDE$ + " (" + PLANO$ + ")"
         End If
       End If
       Call REPLA(XE$, DEDE$, 17, 46)
       Call REPLA(XE$, MKI$(CII%), 63, 2)
       JE& = JE& + 1
       Call GRAREG("ECOMAST", XE$, JE&)
       '
     End If
   Next U&
   '
   NMS% = FIN& - 1
   XM$ = REGLEIDO$("MASTER", 1)
   Call REPLA(XM$, MKI$(NMS%), 1, 2)
   Call GRAREG("MASTER", XM$, 1)
   Call CIERRARCH("MASTER")
   '
   NVI% = JI& - 1
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, MKI$(NVI%), 17, 2)
   Call GRAREG("INVERT", XI$, 1)
   Call SETULTREG("INVERT", JI&)
   Call CIERRARCH("INVERT")
   Call FILESORT("INVERT", "", 1, 1, "ASC")
   '
   Call SETULTREG("ECOMAST", JE&)
   Call CIERRARCH("ECOMAST")
   Call FILESORT("ECOMAST", "", 1, 1, "ASC")
   '
   Call REPUNCODIG
   'YAGENEC% = 1
   Screen.MousePointer = 1
   '
End Sub
