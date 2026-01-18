VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ANACOSTO 
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos Standard - Análisis y Presupuestacion de Productos"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   225
      Left            =   4095
      TabIndex        =   284
      Top             =   105
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Text45 
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
      Left            =   5145
      MaxLength       =   25
      TabIndex        =   282
      Top             =   315
      Width           =   2250
   End
   Begin VB.Frame Frame27 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Import"
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
      Left            =   7800
      TabIndex        =   280
      Top             =   270
      Width           =   885
      Begin VB.CommandButton Command6 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         Picture         =   "ANACOSTO.frx":0000
         TabIndex        =   281
         ToolTipText     =   "Importación de Archivos Maestros desde Excel"
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0FFC0&
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
      Height          =   2535
      Left            =   10740
      TabIndex        =   54
      Top             =   5565
      Width           =   1040
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
         Picture         =   "ANACOSTO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Imprime Ficha del Articulo"
         Top             =   1680
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
         Picture         =   "ANACOSTO.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Complementos, Consultas, Listados"
         Top             =   990
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
         Picture         =   "ANACOSTO.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   300
         Width           =   765
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1100
      Left            =   7560
      ScaleHeight     =   1035
      ScaleWidth      =   60
      TabIndex        =   52
      Top             =   340
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0FFC0&
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
      Height          =   3890
      Left            =   10740
      TabIndex        =   10
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
         Picture         =   "ANACOSTO.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   990
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
         Picture         =   "ANACOSTO.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   300
         Width           =   750
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
         Picture         =   "ANACOSTO.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   1680
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
         Picture         =   "ANACOSTO.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   2370
         Width           =   765
      End
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
         Height          =   690
         Left            =   125
         Picture         =   "ANACOSTO.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   3060
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
      TabIndex        =   7
      Top             =   880
      Width           =   4920
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
      TabIndex        =   6
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
      Picture         =   "ANACOSTO.frx":1AB2
      TabIndex        =   5
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFC0&
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
      Left            =   8790
      TabIndex        =   1
      Top             =   240
      Width           =   2970
      Begin VB.CommandButton COMMAND3 
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
         Left            =   2160
         Picture         =   "ANACOSTO.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   315
         Width           =   645
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
         Left            =   1530
         Picture         =   "ANACOSTO.frx":1F06
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Rest"
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
         Left            =   770
         Picture         =   "ANACOSTO.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   630
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
         Left            =   140
         Picture         =   "ANACOSTO.frx":251A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   630
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
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12648384
      TabCaption(0)   =   "Insumos"
      TabPicture(0)   =   "ANACOSTO.frx":2824
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MARCOFICHA(0)"
      Tab(0).Control(1)=   "Frame22"
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(3)=   "Label17"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Componentes"
      TabPicture(1)   =   "ANACOSTO.frx":2840
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MARCOFICHA(2)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Elaboración"
      TabPicture(2)   =   "ANACOSTO.frx":285C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MARCOFICHA(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Costos"
      TabPicture(3)   =   "ANACOSTO.frx":2878
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame16"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame16 
         BorderStyle     =   0  'None
         Caption         =   "Frame16"
         Height          =   5895
         Left            =   105
         TabIndex        =   135
         Top             =   420
         Width           =   10200
         Begin VB.Frame Frame26 
            Caption         =   "Cobertura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2745
            Left            =   0
            TabIndex        =   206
            Top             =   3150
            Width           =   2325
            Begin VB.TextBox Text44 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   315
               TabIndex        =   261
               Top             =   640
               Width           =   540
            End
            Begin VB.Label Label118 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   315
               TabIndex        =   275
               Top             =   1980
               Width           =   795
            End
            Begin VB.Label Label117 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   1420
               TabIndex        =   213
               Top             =   1980
               Width           =   795
            End
            Begin VB.Label Label113 
               Caption         =   "Mg.Bto(%):    Mg.Neto(%):   "
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   105
               TabIndex        =   274
               Top             =   1780
               Width           =   2100
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   267
               Top             =   965
               Width           =   1275
            End
            Begin VB.Label Label64 
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
               Left            =   105
               TabIndex        =   266
               Top             =   730
               Width           =   2010
            End
            Begin VB.Line Line6 
               BorderColor     =   &H80000005&
               X1              =   0
               X2              =   2310
               Y1              =   1695
               Y2              =   1695
            End
            Begin VB.Line Line5 
               BorderColor     =   &H80000003&
               X1              =   0
               X2              =   2310
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "I.C.Pr:"
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
               Index           =   4
               Left            =   -420
               TabIndex        =   215
               Top             =   2400
               Width           =   1170
            End
            Begin VB.Label Label119 
               Alignment       =   1  'Right Justify
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
               Height          =   285
               Left            =   840
               TabIndex        =   214
               Top             =   2325
               Width           =   795
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Margen:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   -1260
               TabIndex        =   212
               Top             =   1365
               Width           =   2115
            End
            Begin VB.Label Label116 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   211
               Top             =   1290
               Width           =   1275
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pr.Neto:"
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
               Index           =   1
               Left            =   -1260
               TabIndex        =   210
               Top             =   1070
               Width           =   2115
            End
            Begin VB.Label Label115 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   209
               Top             =   640
               Width           =   1275
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   945
               TabIndex        =   208
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pr.Lista:"
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
               Index           =   0
               Left            =   -1260
               TabIndex        =   207
               Top             =   420
               Width           =   2115
            End
         End
         Begin VB.Frame Frame21 
            Caption         =   "Ultimo Cálculo Integral"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   6720
            TabIndex        =   175
            Top             =   2100
            Width           =   3480
            Begin VB.CommandButton Command55 
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
               Height          =   630
               Left            =   2835
               Picture         =   "ANACOSTO.frx":2894
               Style           =   1  'Graphical
               TabIndex        =   176
               ToolTipText     =   "Re-Cálculo General de Costos"
               Top             =   210
               Width           =   540
            End
            Begin VB.Label Label123 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "El"
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
               TabIndex        =   220
               Top             =   600
               Width           =   330
            End
            Begin VB.Label Label122 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1540
               TabIndex        =   219
               Top             =   480
               Width           =   120
            End
            Begin VB.Label Label121 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "hs"
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
               Left            =   2415
               TabIndex        =   218
               Top             =   600
               Width           =   330
            End
            Begin VB.Label Label95 
               Alignment       =   2  'Center
               BackColor       =   &H80000016&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   1680
               TabIndex        =   217
               Top             =   525
               Width           =   750
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               BackColor       =   &H80000016&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   420
               TabIndex        =   216
               Top             =   525
               Width           =   1065
            End
         End
         Begin VB.Frame Frame20 
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
            Height          =   2010
            Left            =   6720
            TabIndex        =   168
            Top             =   0
            Width           =   2325
            Begin VB.PictureBox Picture7 
               BackColor       =   &H80000016&
               Height          =   645
               Left            =   210
               ScaleHeight     =   585
               ScaleWidth      =   1845
               TabIndex        =   172
               Top             =   500
               Width           =   1905
               Begin VB.CommandButton Command54 
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
                  Height          =   645
                  Left            =   0
                  Picture         =   "ANACOSTO.frx":2B9E
                  Style           =   1  'Graphical
                  TabIndex        =   173
                  Top             =   0
                  Width           =   645
               End
               Begin VB.Label Label94 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Análisis de Costos"
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
                  TabIndex        =   174
                  Top             =   105
                  Width           =   855
               End
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H80000016&
               Height          =   645
               Left            =   210
               ScaleHeight     =   585
               ScaleWidth      =   1845
               TabIndex        =   169
               Top             =   1200
               Width           =   1905
               Begin VB.CommandButton Command50 
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
                  Height          =   645
                  Left            =   0
                  Picture         =   "ANACOSTO.frx":2EA8
                  Style           =   1  'Graphical
                  TabIndex        =   170
                  Top             =   0
                  Width           =   645
               End
               Begin VB.Label Label91 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Listado Resumen"
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
                  TabIndex        =   171
                  Top             =   105
                  Width           =   1170
               End
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Costo de Suministros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1850
            Left            =   0
            TabIndex        =   154
            Top             =   0
            Width           =   6630
            Begin VB.Line Line7 
               BorderColor     =   &H80000005&
               X1              =   0
               X2              =   6600
               Y1              =   435
               Y2              =   435
            End
            Begin VB.Label Label90 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2415
               TabIndex        =   167
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label89 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Materias Primas:"
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
               TabIndex        =   166
               Top             =   735
               Width           =   2115
            End
            Begin VB.Label Label88 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2415
               TabIndex        =   165
               Top             =   1005
               Width           =   1275
            End
            Begin VB.Label Label87 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Piezas Compradas:"
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
               TabIndex        =   164
               Top             =   1110
               Width           =   2115
            End
            Begin VB.Label Label86 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2415
               TabIndex        =   163
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label85 
               Alignment       =   1  'Right Justify
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
               Height          =   330
               Left            =   5145
               TabIndex        =   162
               ToolTipText     =   "Doble-click para Ver Analisis de Costos de Materiales y Servicios"
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   3780
               TabIndex        =   161
               Top             =   1365
               Width           =   1275
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   3780
               TabIndex        =   160
               Top             =   1005
               Width           =   1275
            End
            Begin VB.Label Label82 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   3780
               TabIndex        =   159
               Top             =   630
               Width           =   1275
            End
            Begin VB.Label Label81 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Trabajos de Terceros:"
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
               TabIndex        =   158
               Top             =   1470
               Width           =   2115
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000003&
               X1              =   0
               X2              =   6600
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "C.Propios"
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
               Index           =   3
               Left            =   2415
               TabIndex        =   157
               Top             =   210
               Width           =   1170
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Recibidos"
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
               Index           =   1
               Left            =   3780
               TabIndex        =   156
               Top             =   210
               Width           =   1170
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Totales"
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
               Index           =   2
               Left            =   5145
               TabIndex        =   155
               Top             =   210
               Width           =   1170
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Costo de Elaboración"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Index           =   4
            Left            =   0
            TabIndex        =   146
            Top             =   1890
            Width           =   6630
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   5145
               TabIndex        =   153
               ToolTipText     =   "Doble-click Para Ver Análisis de Costos de Mano de Obra "
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2415
               TabIndex        =   152
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipos de Producción:"
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
               TabIndex        =   151
               Top             =   795
               Width           =   2325
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mano de Obra Directa:"
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
               TabIndex        =   150
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2415
               TabIndex        =   149
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   3780
               TabIndex        =   148
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   3780
               TabIndex        =   147
               Top             =   315
               Width           =   1275
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Opción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   9135
            TabIndex        =   143
            Top             =   0
            Width           =   1065
            Begin VB.CommandButton Command45 
               BackColor       =   &H00E0E0E0&
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
               Height          =   645
               Left            =   230
               Picture         =   "ANACOSTO.frx":31B2
               Style           =   1  'Graphical
               TabIndex        =   144
               ToolTipText     =   "Padron Maestro de Proveedores "
               Top             =   500
               Width           =   645
            End
            Begin VB.CommandButton Command48 
               BackColor       =   &H00E0E0E0&
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
               Height          =   645
               Left            =   230
               Picture         =   "ANACOSTO.frx":34BC
               Style           =   1  'Graphical
               TabIndex        =   145
               ToolTipText     =   "Listas de Precios de Venta"
               Top             =   1200
               Width           =   645
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Costo Total Calculado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2745
            Left            =   2415
            TabIndex        =   136
            Top             =   3150
            Width           =   7785
            Begin VB.TextBox Text41 
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
               Left            =   1680
               TabIndex        =   255
               Top             =   690
               Width           =   750
            End
            Begin VB.TextBox Text43 
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
               Left            =   5460
               TabIndex        =   259
               Top             =   1065
               Width           =   645
            End
            Begin VB.TextBox Text42 
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
               Left            =   5460
               TabIndex        =   258
               Top             =   690
               Width           =   645
            End
            Begin VB.TextBox Text39 
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
               Left            =   5460
               TabIndex        =   260
               Top             =   1785
               Width           =   645
            End
            Begin VB.TextBox Text38 
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
               Left            =   1650
               TabIndex        =   257
               Top             =   1815
               Width           =   750
            End
            Begin VB.TextBox Text36 
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
               Left            =   1680
               TabIndex        =   256
               Top             =   1065
               Width           =   750
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6405
               TabIndex        =   265
               Top             =   1440
               Width           =   1275
            End
            Begin VB.Label Label100 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Total Bruto:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3255
               TabIndex        =   264
               Top             =   1515
               Width           =   3060
            End
            Begin VB.Label Label124 
               BackStyle       =   0  'Transparent
               Caption         =   "GG.Elaboración:               %:"
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
               Left            =   180
               TabIndex        =   263
               Top             =   1170
               Width           =   2535
            End
            Begin VB.Label Label93 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   262
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6405
               TabIndex        =   205
               Top             =   690
               Width           =   1275
            End
            Begin VB.Label Label109 
               BackStyle       =   0  'Transparent
               Caption         =   " Costo s/Venta:             %:"
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
               Left            =   4050
               TabIndex        =   204
               Top             =   810
               Width           =   2325
            End
            Begin VB.Label Label108 
               BackStyle       =   0  'Transparent
               Caption         =   " Utilidad:             %:"
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
               Left            =   4620
               TabIndex        =   203
               Top             =   1920
               Width           =   1905
            End
            Begin VB.Label Label107 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   202
               Top             =   1815
               Width           =   1275
            End
            Begin VB.Label Label106 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Standard Total:"
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
               Left            =   4200
               TabIndex        =   201
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label105 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6405
               TabIndex        =   200
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label104 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6405
               TabIndex        =   199
               Top             =   1065
               Width           =   1275
            End
            Begin VB.Label Label103 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   6405
               TabIndex        =   198
               Top             =   2250
               Width           =   1275
            End
            Begin VB.Label Label102 
               BackStyle       =   0  'Transparent
               Caption         =   " Impuestos:             %:"
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
               Left            =   4410
               TabIndex        =   197
               Top             =   1170
               Width           =   2115
            End
            Begin VB.Label Label101 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Precio Básico de Venta:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3360
               TabIndex        =   196
               Top             =   2370
               Width           =   2955
            End
            Begin VB.Label Label99 
               Alignment       =   1  'Right Justify
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
               Height          =   225
               Left            =   2350
               TabIndex        =   195
               Top             =   1935
               Width           =   330
            End
            Begin VB.Label Label98 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Standard Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -420
               TabIndex        =   194
               Top             =   2370
               Width           =   3060
            End
            Begin VB.Label Label97 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gastos Grales Admin. y Ventas:"
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
               Left            =   105
               TabIndex        =   193
               Top             =   1785
               Width           =   1485
            End
            Begin VB.Label Label96 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2730
               TabIndex        =   192
               Top             =   2250
               Width           =   1275
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   191
               Top             =   1815
               Width           =   1275
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2730
               TabIndex        =   142
               Top             =   315
               Width           =   1275
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Costos Directos:"
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
               Left            =   525
               TabIndex        =   141
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   2730
               TabIndex        =   140
               Top             =   1065
               Width           =   1275
            End
            Begin VB.Label Label67 
               BackStyle       =   0  'Transparent
               Caption         =   "GG.Materiales:               %:"
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
               TabIndex        =   139
               Top             =   810
               Width           =   2430
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Std de Manufactura:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -420
               TabIndex        =   138
               Top             =   1515
               Width           =   3060
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
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
               Left            =   2730
               TabIndex        =   137
               Top             =   1440
               Width           =   1275
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "3"
         Height          =   5895
         Index           =   3
         Left            =   -74790
         TabIndex        =   84
         Top             =   420
         Width           =   10095
         Begin VB.Frame Frame4 
            Caption         =   "Análisis de Costo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5800
            Left            =   4830
            TabIndex        =   88
            Top             =   50
            Width           =   5160
            Begin VB.Frame Frame25 
               Caption         =   "Costo Tercerizacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1665
               Index           =   1
               Left            =   210
               TabIndex        =   186
               Top             =   4020
               Width           =   4770
               Begin VB.TextBox Text37 
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
                  Left            =   735
                  TabIndex        =   111
                  Top             =   1260
                  Width           =   1410
               End
               Begin VB.ListBox List2 
                  BeginProperty Font 
                     Name            =   "Fixedsys"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   510
                  Left            =   105
                  TabIndex        =   187
                  Top             =   315
                  Width           =   4530
               End
               Begin VB.Label Label63 
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
                  Left            =   3730
                  TabIndex        =   113
                  ToolTipText     =   "Doble-Click para Seleccionar Fecha"
                  Top             =   915
                  Width           =   885
               End
               Begin VB.Label Label48 
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
                  Left            =   735
                  TabIndex        =   110
                  ToolTipText     =   "Doble-click para Seleccionar Proveedor"
                  Top             =   915
                  Width           =   2295
               End
               Begin VB.Label Label49 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Prov:                                        Fecha:"
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
                  Left            =   255
                  TabIndex        =   268
                  Top             =   990
                  Width           =   3480
               End
               Begin VB.Label Label46 
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
                  Left            =   2205
                  TabIndex        =   112
                  ToolTipText     =   "Doble-Click para Seleccionar Moneda"
                  Top             =   1260
                  Width           =   2415
               End
               Begin VB.Label Label22 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Precio:"
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
                  Left            =   105
                  TabIndex        =   188
                  Top             =   1335
                  Width           =   1905
               End
            End
            Begin VB.Frame Frame25 
               Caption         =   "Costo Elaboración"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1065
               Index           =   0
               Left            =   2730
               TabIndex        =   185
               Top             =   2310
               Width           =   2250
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "$/U 3's:"
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
                  Left            =   -1155
                  TabIndex        =   278
                  Top             =   705
                  Width           =   1905
               End
               Begin VB.Label Label52 
                  Alignment       =   1  'Right Justify
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
                  Left            =   840
                  TabIndex        =   277
                  Top             =   630
                  Width           =   1200
               End
               Begin VB.Label Label73 
                  Alignment       =   1  'Right Justify
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
                  Left            =   840
                  TabIndex        =   190
                  Top             =   270
                  Width           =   1200
               End
               Begin VB.Label Label71 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "$/Unid:"
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
                  Left            =   -1470
                  TabIndex        =   189
                  Top             =   345
                  Width           =   2220
               End
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
               Height          =   330
               Index           =   0
               Left            =   2100
               TabIndex        =   92
               Top             =   1575
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
               TabIndex        =   90
               Top             =   735
               Width           =   3720
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
               TabIndex        =   89
               Text            =   " "
               Top             =   420
               Width           =   1095
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
               TabIndex        =   91
               Text            =   " "
               Top             =   1575
               Width           =   885
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
               TabIndex        =   96
               Top             =   2625
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
               TabIndex        =   97
               Top             =   2625
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
               TabIndex        =   98
               Top             =   2625
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
               TabIndex        =   99
               Top             =   2940
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
               TabIndex        =   100
               Top             =   2940
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
               TabIndex        =   101
               Top             =   2940
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
               TabIndex        =   102
               Top             =   3255
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
               TabIndex        =   103
               Top             =   3255
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
               TabIndex        =   104
               Top             =   3255
               Width           =   360
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
               Left            =   4515
               TabIndex        =   109
               Text            =   " "
               Top             =   3570
               Width           =   465
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
               TabIndex        =   93
               Text            =   " "
               Top             =   1890
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
               TabIndex        =   94
               Text            =   " "
               Top             =   1890
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
               TabIndex        =   95
               Top             =   1890
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
               TabIndex        =   105
               Top             =   3570
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
               TabIndex        =   106
               Top             =   3570
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
               TabIndex        =   107
               Top             =   3570
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
               Left            =   3200
               TabIndex        =   108
               Text            =   " "
               Top             =   3570
               Width           =   600
            End
            Begin VB.Label Label42 
               BackStyle       =   0  'Transparent
               Caption         =   "Operación:                       Lote Standard:"
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
               Left            =   255
               TabIndex        =   134
               Top             =   525
               Width           =   3585
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
               Left            =   2415
               TabIndex        =   133
               Top             =   1575
               Width           =   2565
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Equipo:"
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
               Index           =   0
               Left            =   210
               TabIndex        =   132
               Top             =   1680
               Width           =   960
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "Preparación:         :      :"
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
               Left            =   105
               TabIndex        =   131
               Top             =   2700
               Width           =   2220
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   " T.Std.Fijo:         :      :"
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
               Left            =   240
               TabIndex        =   130
               Top             =   3015
               Width           =   2220
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   "T.Var.Lote:         :      :"
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
               Left            =   220
               TabIndex        =   129
               Top             =   3330
               Width           =   2430
            End
            Begin VB.Label Label50 
               BackStyle       =   0  'Transparent
               Caption         =   "Operarios:              Cat:"
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
               Left            =   280
               TabIndex        =   128
               Top             =   1995
               Width           =   2220
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
               TabIndex        =   127
               Top             =   1890
               Width           =   1935
            End
            Begin VB.Label Label47 
               BackStyle       =   0  'Transparent
               Caption         =   " Mant.Herr:         :      :"
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
               Left            =   225
               TabIndex        =   126
               Top             =   3645
               Width           =   2430
            End
            Begin VB.Label Label53 
               BackStyle       =   0  'Transparent
               Caption         =   "Ciclo:             % DD:"
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
               Left            =   2680
               TabIndex        =   125
               Top             =   3650
               Width           =   1800
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
               TabIndex        =   124
               Top             =   2400
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
               TabIndex        =   123
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
               TabIndex        =   122
               Top             =   1050
               Width           =   885
            End
            Begin VB.Label Label61 
               BackStyle       =   0  'Transparent
               Caption         =   "Revisión:                      Resp:"
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
               Left            =   380
               TabIndex        =   121
               Top             =   1120
               Width           =   2955
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
               TabIndex        =   120
               Top             =   1050
               Width           =   1935
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
            TabIndex        =   86
            Top             =   50
            Width           =   4530
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
               Left            =   1395
               Style           =   1  'Graphical
               TabIndex        =   116
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
               TabIndex        =   114
               ToolTipText     =   "Agrega una Operación al Final de la Lista"
               Top             =   2520
               Width           =   630
            End
            Begin VB.CommandButton Command5 
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
               Left            =   750
               Style           =   1  'Graphical
               TabIndex        =   115
               ToolTipText     =   "Intercala una Operación Antes de la Activa"
               Top             =   2520
               Width           =   630
            End
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
               TabIndex        =   87
               Top             =   420
               Width           =   4320
            End
            Begin VB.CommandButton Command30 
               BackColor       =   &H00E0E0E0&
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
               Height          =   370
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   117
               Top             =   2520
               Width           =   945
            End
            Begin VB.CommandButton Command42 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Listar"
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
               TabIndex        =   118
               Top             =   2520
               Width           =   945
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
            Height          =   2650
            Left            =   105
            TabIndex        =   85
            Top             =   3200
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
               Height          =   2220
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   119
               Top             =   315
               Width           =   4320
            End
         End
         Begin VB.Frame Frame24 
            Caption         =   "No Usados"
            Height          =   3900
            Left            =   4830
            TabIndex        =   177
            Top             =   105
            Visible         =   0   'False
            Width           =   3270
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
               Left            =   2625
               TabIndex        =   184
               Top             =   2310
               Width           =   360
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
               Left            =   2205
               TabIndex        =   183
               Top             =   2310
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
               Left            =   1680
               TabIndex        =   182
               Top             =   2310
               Width           =   465
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
               Left            =   1050
               TabIndex        =   181
               Top             =   840
               Width           =   1830
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
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   180
               Top             =   1890
               Width           =   780
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
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   179
               Top             =   1575
               Width           =   780
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
               Left            =   1305
               MaxLength       =   6
               TabIndex        =   178
               Top             =   1260
               Width           =   780
            End
         End
      End
      Begin VB.Frame MARCOFICHA 
         BorderStyle     =   0  'None
         Caption         =   "2"
         Height          =   5895
         Index           =   2
         Left            =   -74895
         TabIndex        =   59
         Top             =   420
         Width           =   10200
         Begin VB.CommandButton Command101 
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
            Left            =   9380
            Picture         =   "ANACOSTO.frx":3606
            Style           =   1  'Graphical
            TabIndex        =   286
            ToolTipText     =   "Búsqueda de Componente por Código o Descripcion"
            Top             =   380
            Width           =   690
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
            TabIndex        =   68
            Top             =   810
            Width           =   8100
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
            TabIndex        =   271
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
               TabIndex        =   273
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
               TabIndex        =   272
               ToolTipText     =   "Muestra Estructura Multinivel"
               Top             =   360
               Width           =   1170
            End
         End
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
            TabIndex        =   71
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
               Picture         =   "ANACOSTO.frx":496C
               TabIndex        =   79
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
               TabIndex        =   78
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
               TabIndex        =   77
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
               Picture         =   "ANACOSTO.frx":4C76
               TabIndex        =   76
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
               TabIndex        =   75
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
               Picture         =   "ANACOSTO.frx":4F80
               TabIndex        =   74
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
               TabIndex        =   73
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
               Picture         =   "ANACOSTO.frx":528A
               TabIndex        =   72
               ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
               Top             =   1260
               Width           =   855
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
               TabIndex        =   80
               Top             =   420
               Width           =   435
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
               TabIndex        =   83
               Top             =   945
               Width           =   1065
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Código:                                      U-M:           Uso:                       D:"
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
               Left            =   380
               TabIndex        =   82
               Top             =   525
               Width           =   6105
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
               TabIndex        =   81
               Top             =   945
               Width           =   5580
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   210
            ScaleHeight     =   375
            ScaleWidth      =   8025
            TabIndex        =   69
            Top             =   420
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
               TabIndex        =   70
               Top             =   105
               Width           =   7995
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
            TabIndex        =   64
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
               ToolTipText     =   "Lista de Compras"
               Top             =   1100
               Width           =   1170
            End
         End
         Begin VB.CommandButton Command41 
            Height          =   435
            Left            =   8520
            Picture         =   "ANACOSTO.frx":5594
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Sube al Nivel Anterior"
            Top             =   380
            Width           =   690
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
            TabIndex        =   62
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
               TabIndex        =   270
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
               TabIndex        =   269
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
            TabIndex        =   60
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
               TabIndex        =   61
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
         TabIndex        =   19
         Top             =   525
         Width           =   10215
         Begin VB.Frame Frame8 
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
            TabIndex        =   250
            Top             =   3150
            Width           =   4740
            Begin VB.TextBox Text40 
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
               Left            =   210
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   251
               Top             =   315
               Width           =   4320
            End
         End
         Begin VB.Frame Frame9 
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
            Height          =   2430
            Left            =   4935
            TabIndex        =   28
            Top             =   3150
            Width           =   5160
            Begin VB.TextBox Text46 
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
               Left            =   1260
               TabIndex        =   35
               Top             =   795
               Width           =   1410
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
               Height          =   285
               Left            =   3720
               TabIndex        =   36
               Top             =   795
               Width           =   1150
            End
            Begin VB.TextBox Text21 
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
               Left            =   3465
               TabIndex        =   49
               Top             =   1575
               Width           =   1410
            End
            Begin VB.TextBox Text13 
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
               TabIndex        =   40
               Top             =   1170
               Width           =   360
            End
            Begin VB.CommandButton Command15 
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
               Left            =   1620
               TabIndex        =   39
               Top             =   1170
               Width           =   175
            End
            Begin VB.TextBox Text14 
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
               Left            =   3720
               TabIndex        =   33
               Top             =   420
               Width           =   1005
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
               Left            =   4725
               TabIndex        =   34
               Top             =   420
               Width           =   175
            End
            Begin VB.TextBox Text12 
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
               Left            =   1260
               TabIndex        =   32
               Top             =   420
               Width           =   1410
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
               Left            =   1260
               MaxLength       =   5
               TabIndex        =   30
               Top             =   1575
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
               Left            =   2040
               TabIndex        =   29
               Top             =   1575
               Width           =   175
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo / Kg:"
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
               Left            =   -105
               TabIndex        =   285
               Top             =   840
               Width           =   1290
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% G.Nac:"
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
               Left            =   2415
               TabIndex        =   279
               Top             =   885
               Width           =   1290
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "$ / Unidad:"
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
               Index           =   5
               Left            =   2100
               TabIndex        =   50
               Top             =   1650
               Width           =   1290
            End
            Begin VB.Label Label35 
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
               TabIndex        =   43
               Top             =   1170
               Width           =   2985
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
               Left            =   1260
               TabIndex        =   42
               Top             =   1950
               Width           =   3615
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda:"
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
               Index           =   3
               Left            =   0
               TabIndex        =   41
               Top             =   1220
               Width           =   1185
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha de Costeo:"
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
               Index           =   1
               Left            =   2460
               TabIndex        =   38
               Top             =   315
               Width           =   1185
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Costo / UM:"
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
               Left            =   -105
               TabIndex        =   37
               Top             =   470
               Width           =   1290
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Prov.Habit:"
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
               Left            =   -315
               TabIndex        =   31
               ToolTipText     =   "Doble-Click para Control Cambios Proveedor Habitual"
               Top             =   1620
               Width           =   1500
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Ultimas Compras"
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
            Left            =   4935
            TabIndex        =   252
            Top             =   105
            Width           =   5160
            Begin VB.ListBox List4 
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
               Left            =   210
               TabIndex        =   253
               ToolTipText     =   "Doble-Click para Registrar Precio de Compra como Costo de Reposición"
               Top             =   685
               Width           =   4740
            End
            Begin VB.Label Label92 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha     Proveedor           Precio  "
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   210
               TabIndex        =   254
               Top             =   420
               Width           =   4740
            End
            Begin VB.Line Line4 
               X1              =   210
               X2              =   4960
               Y1              =   2785
               Y2              =   2785
            End
            Begin VB.Line Line3 
               X1              =   4965
               X2              =   4965
               Y1              =   400
               Y2              =   2800
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
            Height          =   2955
            Left            =   105
            TabIndex        =   20
            Top             =   105
            Width           =   4740
            Begin VB.TextBox Text23 
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
               Left            =   1155
               TabIndex        =   247
               Top             =   1470
               Width           =   3375
            End
            Begin VB.TextBox Text22 
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
               Left            =   1155
               TabIndex        =   246
               Top             =   1890
               Width           =   3375
            End
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
               Height          =   330
               Left            =   1155
               TabIndex        =   236
               Top             =   2415
               Width           =   1305
            End
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
               Height          =   330
               Left            =   3210
               TabIndex        =   27
               Top             =   2415
               Width           =   1305
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
               Height          =   330
               Left            =   1700
               TabIndex        =   24
               Top             =   945
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
               Height          =   330
               Left            =   1155
               MaxLength       =   2
               TabIndex        =   23
               Text            =   " "
               Top             =   945
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
               Height          =   330
               Left            =   1700
               TabIndex        =   22
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
               Height          =   330
               Left            =   1155
               MaxLength       =   3
               TabIndex        =   21
               Text            =   " "
               Top             =   525
               Width           =   570
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Color:"
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
               TabIndex        =   249
               Top             =   2010
               Width           =   1290
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Talle ó Medida:"
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
               TabIndex        =   248
               Top             =   1410
               Width           =   1185
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lote Económico:"
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
               TabIndex        =   237
               Top             =   2350
               Width           =   1200
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
               Index           =   0
               Left            =   -315
               TabIndex        =   48
               Top             =   450
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
               TabIndex        =   47
               Top             =   890
               Width           =   1185
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Kilos/ Unid:"
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
               Left            =   2310
               TabIndex        =   46
               Top             =   2350
               Width           =   870
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
               Height          =   330
               Left            =   1995
               TabIndex        =   26
               Top             =   945
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
               Height          =   330
               Left            =   1995
               TabIndex        =   25
               Top             =   525
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "No Usados"
         Height          =   4215
         Left            =   -74475
         TabIndex        =   221
         Top             =   630
         Width           =   4320
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
            Height          =   330
            Left            =   1810
            TabIndex        =   241
            Top             =   105
            Width           =   175
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
            Height          =   330
            Left            =   3675
            TabIndex        =   240
            Top             =   945
            Width           =   1305
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
            Height          =   330
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   239
            Text            =   " "
            Top             =   105
            Width           =   360
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
            Height          =   330
            Left            =   525
            TabIndex        =   238
            Top             =   945
            Width           =   1305
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
            Height          =   330
            Left            =   1500
            TabIndex        =   233
            Top             =   1785
            Width           =   175
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
            Height          =   330
            Left            =   1155
            TabIndex        =   232
            Text            =   " "
            Top             =   1785
            Width           =   360
         End
         Begin VB.CommandButton Command34 
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
            Left            =   5010
            TabIndex        =   231
            Top             =   1200
            Width           =   175
         End
         Begin VB.CommandButton Command35 
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
            Left            =   5010
            TabIndex        =   230
            Top             =   1710
            Width           =   175
         End
         Begin VB.CommandButton Command36 
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
            Left            =   1485
            TabIndex        =   227
            Top             =   735
            Width           =   175
         End
         Begin VB.TextBox Text35 
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
            Left            =   945
            MaxLength       =   3
            TabIndex        =   226
            Text            =   "  "
            Top             =   735
            Width           =   570
         End
         Begin VB.TextBox Text16 
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
            Left            =   945
            MaxLength       =   3
            TabIndex        =   223
            Text            =   "  "
            Top             =   315
            Width           =   570
         End
         Begin VB.CommandButton Command13 
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
            Left            =   1485
            TabIndex        =   222
            Top             =   315
            Width           =   175
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
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3465
            TabIndex        =   245
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Minimo:"
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
            Left            =   210
            TabIndex        =   244
            Top             =   630
            Width           =   1395
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
            Left            =   210
            TabIndex        =   243
            Top             =   0
            Width           =   1380
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
            Height          =   330
            Left            =   2100
            TabIndex        =   242
            Top             =   105
            Width           =   2900
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "F.Des:"
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
            Left            =   210
            TabIndex        =   235
            Top             =   1890
            Width           =   855
         End
         Begin VB.Label Label39 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1155
            TabIndex        =   234
            Top             =   2205
            Width           =   3375
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo I.V.A.:"
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
            TabIndex        =   229
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label57 
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
            Left            =   1785
            TabIndex        =   228
            Top             =   735
            Width           =   2535
         End
         Begin VB.Label Label31 
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
            Left            =   1785
            TabIndex        =   225
            Top             =   315
            Width           =   2535
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Contable:"
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
            Left            =   -420
            TabIndex        =   224
            Top             =   210
            Width           =   1290
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   17
         Top             =   4305
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   16
         Top             =   3360
         Width           =   15
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Pl.Nro:"
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
      Index           =   41
      Left            =   4200
      TabIndex        =   283
      Top             =   420
      Width           =   795
   End
   Begin VB.Label COMMLAB 
      Height          =   225
      Left            =   105
      TabIndex        =   276
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label58 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   0
      TabIndex        =   58
      Top             =   315
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   6090
      TabIndex        =   45
      Top             =   840
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
      Left            =   6825
      TabIndex        =   44
      Top             =   1090
      Width           =   540
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
      Left            =   420
      TabIndex        =   18
      Top             =   1090
      Width           =   750
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   105
      TabIndex        =   9
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
      Index           =   0
      Left            =   375
      TabIndex        =   8
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "ANACOSTO"
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
'
Dim PROVINT%(32767)
Dim FOCOREP%
'
Dim ESPORLOSTFOCUS%

Private Sub Command101_Click()
 '\\\OT-05-0830-WAR-21/11/05///
    'ARMO EL ARCHIVO
    If List1.ListCount <= 0 Then
      Call COMUNI("No Existen Elementos en la Fórmula / Estructura")
      Exit Sub
    End If
    '
    Call BLANARCH("!ECOMAST")
    Call FRESHECHO("!ECOMAST")
    JJ& = 0
    '
    For J1% = 0 To List1.ListCount - 1
      CODI1$ = Left$(List1.List(J1%), 16)
      CODI2% = CODINT%(CODI1$)
      '
      Y$ = REGBLAN("ECOMAST")
      Call REPLA(Y$, CODI1$, 1, 16)
      Call REPLA(Y$, DESCRIT0$(CODI2%), 17, 46)
      Call REPLA(Y$, MKI$(CODI2%), 63, 2)
      JJ& = JJ& + 1
      Call GRAREG("!ECOMAST", Y$, JJ&)
    Next J1%
    '
    Call SETULTREG("!ECOMAST", JJ&)
    Call CIERRARCH("!ECOMAST")
    '
    'ARRIBA SELECHO
    Call SELECHO("!ECOMAST")
    Y1$ = VALACT1$("!ECOMAST")
    If TRIM$(Y1$) <> "" Then
      X$ = RECFILT("!ECOMAST", 1, Y1$)
      R1 = CVS(X$)
      'POSICIONO EN EL LIST SEGUN SELECCION EN SELECHO
      List1.ListIndex = R1 - 1
      List1.Refresh
    End If
    '\\\OT-05-0830-WAR-FIN///
End Sub

'
Private Sub Command6_Click()
   Command6.Enabled = False
   If DERACCE%("IMPORMAS", "Importación de Archivos Maestros desde Excel") = 2 Then
     IMPORMAS.Show 1
   End If
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
TTXX$ = ""
FIN& = ULTREG&("INVERT")
For U& = 1 To FIN&
XX$ = REGLEIDO$("INVERT", U&)
TTXX$ = TTXX$ + TRIM$(Left$(XX$, 16)) + "  "
If Len(TTXX$) > 512 Then
  MsgBox TTXX$
  TTXX$ = ""
End If
Next U&
MsgBox TTXX$
'Call CONECRUN("*LICARFO", "")
End Sub


Private Sub COMMLAB_Change()
   If COMMLAB = "PRIMOD" Then
      COMMLAB.Caption = ""
      Call PRIANAMODI
    ElseIf COMMLAB = "PRISUMI" Then
      COMMLAB.Caption = ""
      Call PRIANASUMI
   End If
End Sub

Private Sub Form_Load()
    '
    SSTab1.Tab = 3
    Call CENTRAFORM(Me)
    Show
    Refresh
    ' Call VERJOBID
    '
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Command45.Enabled = False
    If EXISTE%("AMAPRO.EXE") > 0 Then Command45.Enabled = True
    If EXISTE%("AMAPRO01.EXE") > 0 Then Command45.Enabled = True
    '
    Command48.Enabled = False
    If EXISTE%("ADMIPRE.EXE") > 0 Then Command48.Enabled = True
    If EXISTE%("ADMPRE01.EXE") > 0 Then Command48.Enabled = True
    '
    Call CARTAVALISTO
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "DECICAL"))
    If DCC$ = "" Then
        DECICALCOS% = 4
        Call GRACONTROL("ANACOSTO", "DECICAL", TRIM$(Str$(DECICALCOS%)))
      Else
        DECICALCOS% = Val(DCC$)
    End If
    If DECICALCOS% < 0 Or DECICALCOS% > 5 Then DECICALCOS% = 4
    '
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENMAT"))
      Text41 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENFAB"))
      Text36 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "GAGENADM"))
      Text38 = CSTRING$(MKS$(Val(DCC$)), 3, 5, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORUTILI"))
      Text39 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORCOMIS"))
      Text42 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORIMPUE"))
      Text43 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL$("ANACOSTO", "PORDESCU"))
      Text44 = CSTRING$(MKS$(Val(DCC$)), 3, 4, 1, 2)
    DCC$ = TRIM$(CONTROL("ANACOSTO", "FECALCOS"))
      Label34(4) = DCC$
    DCC$ = TRIM$(CONTROL("ANACOSTO", "HOCALCOS"))
      Label95 = DCC$
      '
    Call ACVEPROINT
    '
    Call RECLAPAG   ' RECLAMO DE PAGO
    '
    Text1.SetFocus
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Resize()
   IMPLOSI.Hide
   DoEvents
   If WindowState = 1 Then
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
   Call GENECOMAS
   Call FINAL("")
End Sub

Private Sub Command1_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
End Sub

Private Sub command10_Click()
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
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Command13_Click()
   Call SELECHO("GRUCOVEN")
   VALONUE$ = VALACT1$("GRUCOVEN")
   If VALONUE$ <> "" Then
      Text16.TEXT = VALONUE$
   End If
End Sub

Private Sub Command14_Click()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text10.TEXT))))
   End If
End Sub

Private Sub COMMAND15_CLICK()
   Call SELECHO("TAMONED")
   VALONUE$ = VALACT1$("TAMONED")
   If VALONUE$ <> "" Then
      Text13.TEXT = VALONUE$
   End If
End Sub

Private Sub Command16_Click()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text14.TEXT = VALONUE$
   End If
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
    MODIFIC%(2) = 1
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
    MODIFIC%(2) = 1
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
   Call SELECHO("TACOMPO")
   VALONUE$ = VALACT1$("TACOMPO")
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
   HAYACMAS% = 1
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
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
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
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
    Command25.Enabled = False
    Call BLOQARCH("AMASTO")
    Call REVIMASTER
    Call REPUNCODIG
    Command25.Enabled = True
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
   MODIFIC%(2) = 1
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
   MODIFIC%(2) = 1
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
   Call GENECOMAS
   Call FINAL("")
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
   MODIFIC(3) = 1
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
   RESECO$(List3.ListCount) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(List3.ListCount) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Agregada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(3) = 1
   List3.AddItem SECO$
   List3.ListIndex = List3.ListCount - 1
   Text5.SetFocus
End Sub

Private Sub Command34_Click()
   Call SELECHO("TATALLES")
   TALLENUE$ = VALACT1$("TATALLES")
   If TALLENUE$ <> "" Then
      Text23.TEXT = TALLENUE$
   End If
End Sub

Private Sub Command35_Click()
   Call SELECHO("TACOLOR")
   COLONUE$ = VALACT1$("TACOLOR")
   If COLONUE$ <> "" Then
      Text22.TEXT = COLONUE$
   End If
End Sub

Private Sub Command36_Click()
   Call SELECHO("GRUPIVA")
   VALONUE$ = VALACT1$("GRUPIVA")
   If VALONUE$ <> "" Then
      Text35.TEXT = VALONUE$
   End If
End Sub

Private Sub Command37_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   For KKI% = 0 To 15: YACAR%(KKI%) = 0: Next KKI%
   '
   If DERACCE%("STCOSTO", "Configuración Procesos de Costos") >= 2 Then
     Call CONECRUN("FLSETUP/ANACOSTO", "")
     Call FINAL("")
   End If
   '
   Call ACVEPROINT
   Call CARDATITEM
   '
   CIXI% = CODINT%(Text1)
   MONEI% = Val(Text13.TEXT)
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIXI%) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + PORNACIO(CI%) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   If SSTab1.Tab = 2 Then
     If List3.ListIndex >= 0 Then
       VDEF$ = RESECO$(List3.ListIndex)
       Mid$(VDEF$, 151, 8) = MKD(Val(Text37))
       RESECO$(List3.ListIndex) = VDEF$
       MONEXI% = Val(Mid$(Label46, 49, 6))
       TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
       Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
       MODIFIC%(3) = 1
     End If
   End If
   '
End Sub

Private Sub Command38_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command38.Enabled = False
   OPLIMAS.Show 1
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   If SSTab1.Tab <> 3 Then
        PRINTFORM
      Else
        Call PRIFICHCOS
   End If
   Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   'EPAC$ = TRIM$(EMPREAC$)
   'If InStr(EPAC$, "DOSIVAC") > 0 Then
   '    Call IMPORFORMUL
   '  ElseIf InStr(EPAC$, "ARCON") > 0 Then
   '    Call IMPORARCON
   'End If
End Sub

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

Private Sub Command45_Click()
   '
   Command45.Enabled = False
   If EXISTE%("AMAPRO01.EXE") > 0 Then
       Call CONECRUN("AMAPRO01", "Padron Maestro de Proveedores")
     ElseIf EXISTE%("AMAPRO.EXE") > 0 Then
       Call CONECRUN("AMAPRO", "Padron Maestro de Proveedores")
     Else
       Call OPNOIN("")
   End If
   Command45.Enabled = True
   '
End Sub

Private Sub Command48_Click()
   '
   Command48.Enabled = False
   If EXISTE%("LISPRE02.EXE") > 0 Then
       Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
     ElseIf EXISTE%("ADMIPRE.EXE") > 0 Then
       Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
     Else
       Call OPNOIN("")
   End If
   Command48.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
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
   RESECO$(ILISTAX&) = Left$(MKI$(CI%) + MKI$(NOPY%) + Mid$(REGBLAN$("SECOPER"), 5) + String$(128, 0), 128) + Space$(16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0) 'para acomodar hoja de procesos DE FABRICACION
   DESOPE$(ILISTAX&) = ""
   NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
   DOPE$ = AJUSTI$("Operación Insertada (" + TRIM$(NOPEX$) + ")", 48)
   REGO& = 0
   REGOX$ = AJUSTD$(Str$(REGO&), 6)
   SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
   MODIFIC%(3) = 1
   List3.AddItem SECO$, ILISTAX&

   List3.ListIndex = ILISTAX&
   Text5.SetFocus
   '
End Sub

Private Sub Command51_Click()

End Sub

Private Sub Command50_Click()
   Command50.Enabled = False
   If ULTREG&("RECOSCAL") < 1 Then
       Call MENSERR(24, "Falta Generar Archivo Resumen")
     Else
       FILISCOS.Show 1
   End If
   Command50.Enabled = True
End Sub

Private Sub Command54_Click()
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
   Command54.Enabled = False
   CIXI% = CODINT%(Text1.TEXT)
   If CIXI% < 1 Then GoTo 99
   RFF$ = RECFILT$("CARFORLI", 1, MKI$(CIXI%))
   If Len(RFF$) >= 4 Then
     OPX% = COMALTER%("Elija Formato de Impresión:\\Planilla\Analisis\Ficha\Cancelar")
     If OPX% = 1 Then
         Call LIFORCOS(1)
         GoTo 99
       ElseIf OPX% = 2 Then
         Call LIFORCOS(2)
         GoTo 99
       ElseIf OPX% = 3 Then
         GoTo 50
     End If
     GoTo 99
   End If
   '
50 Call PRIFICHCOS
   Call PRIANASUMI
   Call PRIANAMODI
   '
99 Command54.Enabled = True
   '
End Sub

Sub LIFORCOS(OPX%)
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    '
    JJ& = 0: GRUPO% = 0: J2& = 0
    REBLA$ = REGBLAN$("LIFORLI")
    '
    JJ& = JJ& + 1
    YY$ = REBLA$
      JI% = JJ&
    Call REPLA(YY$, MKI$(JI%), 1, 2)
      CANTI = 1
    Call REPLA(YY$, MKS$(1), 3, 4)
      NIVIEL% = 0: GRUPO% = 0
      CODIEL% = CIACT%
      NPLA$ = NUPLANO$(CIACT%)
      NPOS$ = "CONJ"
      NDES$ = DESCRIT0$(CIACT%)
    Call REPLA(YY$, MKI$(CODIEL%), 7, 2)
    Call REPLA(YY$, NPLA$, 9, 16)
    Call REPLA(YY$, NPOS$, 25, 8)
    Call REPLA(YY$, NDES$, 33, 40)
    Call REPLA(YY$, MKI$(GRUPO%), 73, 2)
    Call REPLA(YY$, Space$(48), 81, 48)
    Call REPLA(YY$, MKI$(0), 129, 2)
    Call REPLA(YY$, "CONJ", 157, 36)
      '
      TSTA = TIESTAN(CODIEL%, 0, CDbl(CANTI), 0)   ' EN MINUTOS
    Call REPLA(YY$, MKS$(TSTA / 60), 197, 4)       ' HORAS DE MANO DE OBRA 197,4
      COSMOD = COSTAN(CODIEL%, 0, CDbl(CANTI), 1)
    Call REPLA(YY$, MKS$(COSMOD), 201, 4)     ' COSTO DE MANO DE OBRA 201,4
      COSMAQ = COSTAN(CODIEL%, 0, CDbl(CANTI), 2)
    Call REPLA(YY$, MKS$(COSMAQ), 205, 4)     ' COSTO DE MAQUINAS 205,4
      COSELA = COSTAN(CODIEL%, 0, CDbl(CANTI), 0)
    Call REPLA(YY$, MKS$(COSELA), 209, 4)     ' COSTO DE ELABORACION 209,4
      COS3S = COSOPTER(CODIEL%, 0, CDbl(CANTI), 1)
    Call REPLA(YY$, MKS$(COS3S), 213, 4)      ' TRABAJOS DE 3ROS EN 213,4
      COSTT = COSOPTER(CODIEL%, 0, CDbl(CANTI), 2)
    Call REPLA(YY$, MKS$(COSTT), 217, 4)      ' TRATAMIENTO TERMICO EN 217,4
      '
    Call GRAREG("LIFORLI", YY$, JJ&)
    '
    '
    RNC$ = RECFILT$("COSOPTER", 1, MKI$(CIACT%))
    URN& = ULTREG&("COSOPTER")
    For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
        If RENOTA& <= URN& Then
          TTXX$ = REGLEIDO$("COSOPTER", RENOTA&)
          NOPEI% = CVI(Mid$(TTXX$, 3, 2))
          If TIESTAN(CIACT%, NOPEI%, 1, 0) < 1 Then
            PREUN = CVD(Mid$(TTXX$, 7, 8))
            MONEM% = CVI(Mid$(TTXX$, 15, 2))
            DESOP$ = UCase$(DESCRIOP$(CODIEL%, NOPEI%))
            RUBG% = 91
            '
            Y2$ = REGBLAN$("LISGRUN")
            Call REPLA(Y2$, MKI$(RUBG%), 1, 2)
            Call REPLA(Y2$, "EXTERNOS", 3, 24)
            Call REPLA(Y2$, MKS$(CANTI), 27, 4)
            Call REPLA(Y2$, "UNID", 31, 4)
            Call REPLA(Y2$, MKI$(0), 35, 2)
            Call REPLA(Y2$, DESOP$, 37, 48)
            Call REPLA(Y2$, MKS$(PREUN * TICAMONE(MONEM%)), 109, 4)
            Call REPLA(Y2$, MKD$(CANTI * PREUN * TICAMONE(MONEM%)), 113, 8)
            J2& = J2& + 1
            Call GRAREG("LISGRUN", Y2$, J2&)
            '
          End If
        End If
      End If
    Next U&
    '
    LOTESTA = LOTESTAN(CIACT%): If LOTESTA < 1 Then LOTESTA = 1
    RNC$ = RECFILT$("SECOPER", 1, MKI$(CIACT%))
    URN& = ULTREG&("SECOPER")
    '
    For U& = 1 To Len(RNC$) Step 4
      RENOTA& = CVS(Mid$(RNC$, U&, 4))
      If RENOTA& > 0 Then
        If RENOTA& <= URN& Then
          TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
          NOPEI% = CVI(Mid$(TTXX$, 3, 2))
          '
          'CALOTES = 1
          CALOTES = CANTI / LOTESTA
          CAOPS = CVS(Mid$(TTXX$, 83, 4))
          If CAOPS < 0.01 Then CAOPS = 1
          CATOP% = Asc(Mid$(TTXX$, 87, 1))
          EQUIPO$ = TRIM$(Mid$(TTXX$, 53, 6))
          DESOP$ = UCase$(DESCRIOP$(CIACT%, NOPEI%))
          '
          MINPREP = 60 * CVS(Mid$(TTXX$, 59, 4)) * CALOTES
          If MODO% > 0 Then MINPREP = 0
          '
          MINFIJO = 60 * CVS(Mid$(TTXX$, 63, 4)) * CALOTES
          MINVARI = CANTI * CVS(Mid$(TTXX$, 67, 4))
          MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(TTXX$, 75, 4)) / 100
          MINLIMP = 60 * CVS(Mid$(TTXX$, 79, 4)) * CALOTES
          MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
          '
          HORAPROD = MINLOTE / 60
          COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
          COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
          '
          If COSMOD > 0.005 Then
            Y2$ = REGBLAN$("LISGRUN")
            Call REPLA(Y2$, MKI$(95), 1, 2)
            Call REPLA(Y2$, "MANO DE OBRA", 3, 24)
            'Call REPLA(Y2$, MKS$(CANTI * HORAPROD), 27, 4)
            Call REPLA(Y2$, MKS$(HORAPROD), 27, 4)
            Call REPLA(Y2$, "HS", 31, 4)
            Call REPLA(Y2$, MKI$(0), 35, 2)
            Call REPLA(Y2$, DESOP$, 37, 48)
            Call REPLA(Y2$, MKS$(VALHORA(CATOP%)), 109, 4)
            'Call REPLA(Y2$, MKD$(CANTI * COSMOD), 113, 8)
            Call REPLA(Y2$, MKD$(CDbl(COSMOD)), 113, 8)
            J2& = J2& + 1
            Call GRAREG("LISGRUN", Y2$, J2&)
          End If
          '
          If COSMAQ > 0.005 Then
            Y2$ = REGBLAN$("LISGRUN")
            Call REPLA(Y2$, MKI$(95), 1, 2)
            Call REPLA(Y2$, "EQUIPOS", 3, 24)
            'Call REPLA(Y2$, MKS$(CANTI * HORAPROD), 27, 4)
            Call REPLA(Y2$, MKS$(HORAPROD), 27, 4)
            Call REPLA(Y2$, "HS", 31, 4)
            Call REPLA(Y2$, MKI$(0), 35, 2)
            Call REPLA(Y2$, DESOP$, 37, 48)
            Call REPLA(Y2$, MKS$(CUHORMAQ(EQUIPO$)), 109, 4)
            'Call REPLA(Y2$, MKD$(CANTI * COSMAQ), 113, 8)
            Call REPLA(Y2$, MKD$(CDbl(COSMAQ)), 113, 8)
            J2& = J2& + 1
            Call GRAREG("LISGRUN", Y2$, J2&)
          End If
        End If
      End If
    Next U&
    '
    '
    RFF$ = RECFILT$("CARFORLI", 1, MKI$(CIACT%))
    For KU& = 1 To Len(RFF$) Step 4
      RECA& = CVS(Mid$(RFF$, KU&, 4))
      ZZ$ = REGLEIDO$("CARFORLI", RECA&)
      '
      JJ& = JJ& + 1
      YY$ = REBLA$
        JI% = JJ&
      Call REPLA(YY$, MKI$(JI%), 1, 2)
        CANTI = CVS(Mid$(ZZ$, 93, 4))
      Call REPLA(YY$, MKS$(CANTI), 3, 4)
        NIVIEL% = Asc(Mid$(ZZ$, 18, 1))
        If NIVIEL% = 1 Then GRUPO% = GRUPO% + 1
        CODIEL% = CVI(Mid$(ZZ$, 19, 2))
        NPLA$ = TRIM$(Mid$(ZZ$, 21, 16))
        NPOS$ = TRIM$(Mid$(ZZ$, 37, 8))
        NDES$ = Left$(".                     ", 2 * NIVIEL%) + TRIM$(Mid$(ZZ$, 45, 48))
      Call REPLA(YY$, MKI$(CODIEL%), 7, 2)
      Call REPLA(YY$, NPLA$, 9, 16)
      Call REPLA(YY$, NPOS$, 25, 8)
      Call REPLA(YY$, NDES$, 33, 40)
      Call REPLA(YY$, MKI$(GRUPO%), 73, 2)
        CODIMP% = CVI(Mid$(ZZ$, 97, 2))
        If CODIMP% < 1 Then
          If CODIEL% > 0 Then
            If COSUNI(CODIEL%) > 0 Then
              CODIMP% = CODIEL%
            End If
          End If
        End If
      Call REPLA(YY$, DESCRIT0$(CODIMP%), 81, 48)
      Call REPLA(YY$, MKI$(CODIMP%), 129, 2)
        CANTID = CANTI
        If CODIMP% < 1 Then CANTID = 0
        SONKILOS% = 0
        MEDI2 = CVS(Mid$(ZZ$, 103, 4))
          If MEDI2 > 0.05 Then CANTID = CANTID * (MEDI2 + 5) / 1000: SONKILOS% = 1
        MEDI3 = CVS(Mid$(ZZ$, 107, 4))
          If MEDI3 > 0.05 Then CANTID = CANTID * (MEDI3 + 5) / 1000: SONKILOS% = 1
        MEDI1 = CVS(Mid$(ZZ$, 99, 4))
          If MEDI1 > 0.05 Then CANTID = CANTID * (MEDI1 + 3) / 1000: SONKILOS% = 1
        KILOUN = PESUNI(CODIMP%)
          If SONKILOS% > 0 Then
            If KILOUN > 0.0005 Then
              CANTID = CANTID * KILOUN
            End If
          End If
        PORSCRA = CVS(Mid$(ZZ$, 115, 4))
          CANTID = CANTID * (1 + PORSCRA / 100)
      Call REPLA(YY$, MKS$(MEDI1), 131, 4)
      Call REPLA(YY$, MKS$(MEDI2), 135, 4)
      Call REPLA(YY$, MKS$(MEDI3), 139, 4)
      If SONKILOS% > 0 Then
        Call REPLA(YY$, MKS$(KILOUN), 143, 4)
        Call REPLA(YY$, MKS$(PORSCRA), 147, 4)
        Call REPLA(YY$, MKS$(CANTID), 151, 4)
      End If
      Call REPLA(YY$, MKI$(DEPOPRE%(CODIMP%)), 155, 2)
        OBSER$ = TRIM$(Mid$(ZZ$, 125, 36))
      Call REPLA(YY$, OBSER$, 157, 36)
      '
      COSUN = COSPESOS(CODIMP%)
      If SONKILOS% = 1 Then
        If KILOUN > 0.0005 Then
          COSUN = COSUN / KILOUN
        End If
      End If
      Call REPLA(YY$, MKS$(CANTID * COSUN), 193, 4)
        TSTA = TIESTAN(CODIEL%, 0, CDbl(CANTI), 0)   ' EN MINUTOS
      Call REPLA(YY$, MKS$(TSTA / 60), 197, 4)       ' HORAS DE MANO DE OBRA 197,4
        COSMOD = COSTAN(CODIEL%, 0, CDbl(CANTI), 1)
      Call REPLA(YY$, MKS$(COSMOD), 201, 4)     ' COSTO DE MANO DE OBRA 201,4
        COSMAQ = COSTAN(CODIEL%, 0, CDbl(CANTI), 2)
      Call REPLA(YY$, MKS$(COSMAQ), 205, 4)     ' COSTO DE MAQUINAS 205,4
        COSELA = COSTAN(CODIEL%, 0, CDbl(CANTI), 0)
      Call REPLA(YY$, MKS$(COSELA), 209, 4)     ' COSTO DE ELABORACION 209,4
        COS3S = COSOPTER(CODIEL%, 0, CDbl(CANTI), 1)
      Call REPLA(YY$, MKS$(COS3S), 213, 4)      ' TRABAJOS DE 3ROS EN 213,4
        COSTT = COSOPTER(CODIEL%, 0, CDbl(CANTI), 2)
      Call REPLA(YY$, MKS$(COSTT), 217, 4)      ' TRATAMIENTO TERMICO EN 217,4
        '
      Call GRAREG("LIFORLI", YY$, JJ&)
      '
      '
      If CODIMP% > 0 Then
        Y2$ = REGBLAN$("LISGRUN")
        Call REPLA(Y2$, MKI$(TIMATE%(CODIMP%)), 1, 2)
        Call REPLA(Y2$, UCase$(ECOARCH$("TABTIMA", MKI$(TIMATE%(CODIMP%)))), 3, 24)
        Call REPLA(Y2$, MKS$(CANTI), 27, 4)
        Call REPLA(Y2$, UNIMED$(CODIMP%), 31, 4)
        Call REPLA(Y2$, MKI$(CODIMP%), 35, 2)
          DEDE$ = TRIM$(DESCRIT$(CODIMP%))
        Call REPLA(Y2$, DEDE$, 37, 48)
        Call REPLA(Y2$, MKS$(MEDI1), 85, 4)
        Call REPLA(Y2$, MKS$(MEDI2), 89, 4)
        Call REPLA(Y2$, MKS$(MEDI3), 93, 4)
        If SONKILOS% > 0 Then
          Call REPLA(Y2$, MKS$(KILOUN), 97, 4)
          Call REPLA(Y2$, MKS$(PORSCRA), 101, 4)
          Call REPLA(Y2$, MKS$(CANTID), 105, 4)
        End If
        Call REPLA(Y2$, MKS$(COSUN), 109, 4)
        Call REPLA(Y2$, MKD$(CANTID * COSUN), 113, 8)
        J2& = J2& + 1
        Call GRAREG("LISGRUN", Y2$, J2&)
      End If
      '
      RNC$ = RECFILT$("COSOPTER", 1, MKI$(CODIEL%))
      URN& = ULTREG&("COSOPTER")
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("COSOPTER", RENOTA&)
            NOPEI% = CVI(Mid$(TTXX$, 3, 2))
            If TIESTAN(CODIEL%, NOPEI%, 1, 0) < 1 Then
              PREUN = CVD(Mid$(TTXX$, 7, 8))
              MONEM% = CVI(Mid$(TTXX$, 15, 2))
              DESOP$ = UCase$(DESCRIOP$(CODIEL%, NOPEI%))
              RUBG% = 91
              '
              Y2$ = REGBLAN$("LISGRUN")
              Call REPLA(Y2$, MKI$(RUBG%), 1, 2)
              Call REPLA(Y2$, "EXTERNOS", 3, 24)
              Call REPLA(Y2$, MKS$(CANTI), 27, 4)
              Call REPLA(Y2$, "UNID", 31, 4)
              Call REPLA(Y2$, MKI$(0), 35, 2)
              Call REPLA(Y2$, DESOP$, 37, 48)
              Call REPLA(Y2$, MKS$(PREUN * TICAMONE(MONEM%)), 109, 4)
              Call REPLA(Y2$, MKD$(CANTI * PREUN * TICAMONE(MONEM%)), 113, 8)
              J2& = J2& + 1
              Call GRAREG("LISGRUN", Y2$, J2&)
              '
            End If
          End If
        End If
      Next U&
      '
      LOTESTA = LOTESTAN(CODIEL%): If LOTESTA < 1 Then LOTESTA = 1
      RNC$ = RECFILT$("SECOPER", 1, MKI$(CODIEL%))
      URN& = ULTREG&("SECOPER")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
            NOPEI% = CVI(Mid$(TTXX$, 3, 2))
            '
            'CALOTES = 1
            CALOTES = CANTI / LOTESTA
            CAOPS = CVS(Mid$(TTXX$, 83, 4))
            If CAOPS < 0.01 Then CAOPS = 1
            CATOP% = Asc(Mid$(TTXX$, 87, 1))
            EQUIPO$ = TRIM$(Mid$(TTXX$, 53, 6))
            DESOP$ = UCase$(DESCRIOP$(CODIEL%, NOPEI%))
            '
            MINPREP = 60 * CVS(Mid$(TTXX$, 59, 4)) * CALOTES
            If MODO% > 0 Then MINPREP = 0
            '
            MINFIJO = 60 * CVS(Mid$(TTXX$, 63, 4)) * CALOTES
            MINVARI = CANTI * CVS(Mid$(TTXX$, 67, 4))
            MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(TTXX$, 75, 4)) / 100
            MINLIMP = 60 * CVS(Mid$(TTXX$, 79, 4)) * CALOTES
            MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
            '
            HORAPROD = MINLOTE / 60
            COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
            COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
            '
            If COSMOD > 0.005 Then
              Y2$ = REGBLAN$("LISGRUN")
              Call REPLA(Y2$, MKI$(95), 1, 2)
              Call REPLA(Y2$, "MANO DE OBRA", 3, 24)
              'Call REPLA(Y2$, MKS$(CANTI * HORAPROD), 27, 4)
              Call REPLA(Y2$, MKS$(HORAPROD), 27, 4)
              Call REPLA(Y2$, "HS", 31, 4)
              Call REPLA(Y2$, MKI$(0), 35, 2)
              Call REPLA(Y2$, DESOP$, 37, 48)
              Call REPLA(Y2$, MKS$(VALHORA(CATOP%)), 109, 4)
              'Call REPLA(Y2$, MKD$(CANTI * COSMOD), 113, 8)
              Call REPLA(Y2$, MKD$(CDbl(COSMOD)), 113, 8)
              J2& = J2& + 1
              Call GRAREG("LISGRUN", Y2$, J2&)
            End If
            '
            If COSMAQ > 0.005 Then
              Y2$ = REGBLAN$("LISGRUN")
              Call REPLA(Y2$, MKI$(95), 1, 2)
              Call REPLA(Y2$, "EQUIPOS", 3, 24)
              'Call REPLA(Y2$, MKS$(CANTI * HORAPROD), 27, 4)
              Call REPLA(Y2$, MKS$(HORAPROD), 27, 4)
              Call REPLA(Y2$, "HS", 31, 4)
              Call REPLA(Y2$, MKI$(0), 35, 2)
              Call REPLA(Y2$, DESOP$, 37, 48)
              Call REPLA(Y2$, MKS$(CUHORMAQ(EQUIPO$)), 109, 4)
              'Call REPLA(Y2$, MKD$(CANTI * COSMAQ), 113, 8)
              Call REPLA(Y2$, MKD$(CDbl(COSMAQ)), 113, 8)
              J2& = J2& + 1
              Call GRAREG("LISGRUN", Y2$, J2&)
            End If
          End If
        End If
      Next U&
      '
    Next KU&
    '
    Call SETULTREG("LIFORLI", JJ&)
    Call CIERRARCH("LIFORLI")
    Call SETULTREG("LISGRUN", J2&)
    Call CIERRARCH("LISGRUN")
    '
    If OPX% = 1 Then
        Call HEADERS("LIFORLI", "")
        Call HEADERS("LIFORLI", "Conjunto: " + TRIM$(Text3))
        Call FINAL("*LIFORCOS")
      Else
        Call FILESORT("LISGRUN", "", 5, 5, "ASC")
        Call FILESORT("LISGRUN", "", 1, 1, "ASC")
        '
        SUMATO# = 0
        For KU& = 1 To ULTREG&("LISGRUN")
          XX$ = REGLEIDO$("LISGRUN", KU&)
          SUMATO# = SUMATO# + CVD(Mid$(XX$, 113, 8))
        Next KU&
        '
        If SUMATO# > 0.1 Then
          SUMAPAR# = 0
          FIN& = ULTREG&("LISGRUN")
          For KU& = 1 To FIN&
            XX$ = REGLEIDO$("LISGRUN", KU&)
            RUACT% = CVI(Left$(XX$, 2))
            SUMAPAR# = SUMAPAR# + CVD(Mid$(XX$, 113, 8))
            RUSIG% = (-1)
            If KU& < FIN& Then
              YY$ = REGLEIDO$("LISGRUN", KU& + 1)
              RUSIG% = CVI(Left$(YY$, 2))
            End If
            If RUSIG% <> RUACT% Then
              Call REPLA(XX$, MKD$(100 * SUMAPAR# / SUMATO#), 129, 8)
              Call GRAREG("LISGRUN", XX$, KU&)
              SUMAPAR# = 0
            End If
          Next KU&
        End If
        '
        Call HEADERS("LISGRUN", "")
        Call HEADERS("LISGRUN", "Conjunto: " + TRIM$(Text3))
        Call FINAL("*LISGRUN")
    End If
    '
End Sub

Private Sub Command55_Click()
   Command55.Enabled = False
   RECALCOSTO% = 1
   '
   On Error Resume Next
   RECALCOS.Show
   On Error GoTo 0
   Command55.Enabled = True
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

Private Sub Label103_Change()
   Label107 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text39) / 100), 3, 12, DECICALCOS%, 2))
   Label110 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text42) / 100), 3, 12, DECICALCOS%, 2))
   Label104 = TRIM$(CSTRING$(MKD$(Val(Label103) * Val(Text43) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label104_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label105_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label110_Change()
   Label111 = TRIM$(CSTRING$(MKD$(Val(Label105) + Val(Label110) + Val(Label104)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label111_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) - Val(Label111)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label112_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
   '
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
   Label118 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * (Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - (Val(Label85) + Val(Label93))) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label118 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label118.ForeColor = RGB(0, 0, 0)
   If Val(Label118) < 0 Then Label118.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label114_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Text44) / 100), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label115_Change()
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
   Label112 = TRIM$(CSTRING$(MKD$(Val(Label114) - Val(Label115)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label116_Change()
   Label117 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * Val(Label116) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label117 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label117.ForeColor = RGB(0, 0, 0)
   If Val(Label117) < 0 Then Label117.ForeColor = RGB(255, 0, 0)
   '
   Label118 = ""
   If Abs(Val(Label112)) > 0.005 Then
     PORMA# = 100 * (Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - (Val(Label85) + Val(Label93))) / Val(Label112)
     If Abs(PORMA#) <= 99.9 Then
       Label118 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label118.ForeColor = RGB(0, 0, 0)
   If Val(Label118) < 0 Then Label118.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label117_Change()
   Label119 = ""
   If Val(Text39) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Text39)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label13_DblClick()
   CI% = CODINT%(Text1)
   If CI% > 0 Then
     Call TRALOFIL("RECAMPH", 1, MKI$(CI%))
     If ULTREG&("!RECAMPH") < 1 Then
          Call MENSERR(24, "Sin Cambios Registrados\para Código " + TRIM$(CODEXT$(CI%)))
        Else
          VTB% = VENTABU%("RECAMPH/TITUPAN=Cambios Proveedor Habitual Código " + TRIM$(CODEXT$(CI%)))
     End If
   End If
End Sub

Private Sub Label27_DblClick()
   Call Picture5_DblClick
End Sub

Private Sub Label38_DblClick()
   Screen.MousePointer = 11
   INI& = ULTREG&("MOVISTO") - 8192
   If INI& < 1 Then INI& = 1
   JJ& = 0
   Call COPYSTRU("MOVISTO", "MOVISTF")
   For U& = INI& To ULTREG&("MOVISTO")
     XX$ = REGLEIDO$("MOVISTO", U&)
     JJ& = JJ& + 1
     Call GRAREG("MOVISTF", XX$, JJ&)
   Next U&
   Call SETULTREG("MOVISTF", JJ&)
   Call CIERRARCH("MOVISTF")
   Call CIERRARCH("MOVISTO")
   Screen.MousePointer = 1
End Sub

Private Sub Label46_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 159, 2) = MKI(Val(Mid$(Label46, 49, 6)))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = Val(Mid$(Label46, 49, 6))
   TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   MODIFIC%(3) = 1
End Sub

Private Sub Label46_DblClick()
   '
   Call SELECHO("TAMONED")
   MONEI% = Val(TRIM$(VALACT1$("TAMONED")))
   If MONEI% > 0 Then
     MONEX$ = AJUSTI$(VALACT2$("TAMONED"), 48) + Str$(MONEI%)
     Label46.Caption = MONEX$
   End If
   '
End Sub

Private Sub Label48_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 149, 2) = MKI(Val(Mid$(Label48, 49, 6)))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Label48_DblClick()
   '
   Call SELECHO("ACREDOR")
   NPRO% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NPRO% > 0 Then
     PROVEX$ = AJUSTI$(VALACT2$("ACREDOR"), 48) + Str$(NPRO%)
     Label48.Caption = PROVEX$
   End If
   '
End Sub

Private Sub Label63_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 161, 2) = MKI(FECHANUM(Label63))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Label63_DblClick()
   Call SELECHO("SELFECHA")
   FECHU$ = TRIM$(VALACT1$("SELFECHA"))
   If FECHU$ <> "" Then
     Label63.Caption = FECHU$
   End If
End Sub

Private Sub Label65_Change()
   Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Text38) / 100), 3, 12, DECICALCOS%, 2))
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label68_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label70_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label77_Change()
   Label96 = TRIM$(CSTRING$(MKD$(Val(Label65) + Val(Label77)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label80_Change()
   Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Text36) / 100), 3, 12, DECICALCOS%, 2))
   Label68.Refresh
End Sub

Private Sub Label80_DBLClick()
   Label80.Enabled = False
   If ANACOSMOD.List1.ListCount > 0 Then
     ANACOSMOD.Show 1
   End If
   Label80.Enabled = True
End Sub

Private Sub Label85_Change()
   Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Text41) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label85_dblClick()
   Label85.Enabled = False
   If ANACOSMAT.List1.ListCount > 0 Then
     ANACOSMAT.Show 1
   End If
   Label85.Enabled = True
End Sub

Private Sub Label93_Change()
   Label65 = TRIM$(CSTRING$(MKD$(Val(Label70) + Val(Label68) + Val(Label93)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Label96_Change()
   Label105 = Label96
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
   Call BLANFORCOS
   
   Text3.TEXT = ""
   Label37.Caption = ""
   Label28.Caption = ""
   Text45 = ""
   '
   Text4.TEXT = ""
   Label30.Caption = ""
   Text6.TEXT = ""
   Label29.Caption = ""
   Text8.TEXT = ""
   Label32.Caption = ""
   Text13.TEXT = ""
   Label35.Caption = ""
   Text16.TEXT = ""
   Label31.Caption = ""
   Text35.TEXT = ""
   Label57.Caption = ""
   Text10.TEXT = ""
   Label33.Caption = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   Text20.TEXT = ""
   Text12.TEXT = ""
   Text14.TEXT = ""
   Text21.TEXT = ""
   Text22.TEXT = ""
   Text23.TEXT = ""
   Text24.TEXT = ""
   Text25.TEXT = ""
   Text46 = ""
   Text40 = ""
   Label39.Caption = ""
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
    Text45 = NUPLANO$(CI%)
    MODIFIC(0) = 0
    FOCOREP% = 1
    '
    Select Case SSTab1.Tab
      Case 0
        Call CARGENITEM
      Case 1
        Call CARFORMULA
      Case 2
        Call CARHORUTA
      Case 3
        Call CARCOSTO
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
      Label32.Caption = ECOARCH$("TACRIRE", Chr$(CRIT%))
         MONEI% = Asc(Mid$(REMAS$, 95, 1))
      Text13.TEXT = TRIM$(Str$(MONEI%))
      Label35.Caption = ECOARCH$("TAMONED", Chr$(MONEI%))
         GRUCO% = CVI(Mid$(REMAS$, 125, 2))
      Text16.TEXT = TRIM$(Str$(GRUCO%))
      Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(GRUCO%))
         GRUIVA% = Asc(Mid$(REMAS$, 3, 1))
      Text35.TEXT = TRIM$(Str$(GRUIVA%))
      Label57.Caption = ECOARCH$("GRUPIVA", Chr$(GRUIVA%))
      Label57.Refresh
         NPRO% = CVI(Mid$(REMAS$, 123, 2))
      Text10.TEXT = TRIM$(Str$(NPRO%))
      Label33.Caption = ECOARCH$("ACREDOR", MKI$(NPRO%))
      Text18.TEXT = CSTRING$(MKS$(STOMIN(CI%)), 3, 10, 1, 2)
      Text19.TEXT = CSTRING$(MKS$(LOTESTAN(CI%)), 3, 10, 1, 2)
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(CI%)), 3, 10, 1, 2)
         CP = CVS(Mid$(REMAS$, 99, 4))
      Text24.TEXT = LTITPE$(CP)
      Text12.TEXT = CSTRING$(MKS$(COSUNI(CI%)), 3, 11, 4, 2)
         FEX = CVI(Mid$(REMAS$, 115, 2))
      Text14.TEXT = FECHATEX$(FEX)
         PORNA = CVI(Mid$(REMAS$, 127, 2)) / 100
      Text28.TEXT = AJUSTD$(Format$(PORNA, "###0.00"), 8)
      MONEI% = Val(Text13.TEXT)
      COSUDOL = Val(Text12)
      '
      If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNA / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
      End If
      Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
         '
         CC% = Asc(Mid$(REMAS$, 96, 1))
      Text22.TEXT = LCOLOR$(CC%)
         CT% = Asc(Mid$(REMAS$, 97, 1))
      Text23.TEXT = LTALLE$(CT%)
         CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
      Text25.TEXT = TRIM$(Str$(CCOMPO%))
      Label39.Caption = TRIM$(ECOARCH$("TACOMPO", Chr$(CCOMPO%)))
      '
      List4.Clear
      UAGRE$ = ""
      URE& = ULTREG&("OCOMDETA")
      RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
      For KKI% = 1 To Len(RFF$) Step 4
        REOCO& = CVS(Mid$(RFF$, KKI%, 4))
        If REOCO& > 0 Then
          If REOCO& <= URE& Then
            XX$ = REGLEIDO$("OCOMDETA", REOCO&)
            STAT% = Asc(Mid$(XX$, 124, 1))
            If STAT% <> 9 Then    ' Or CVS(Mid$(XX$, 99, 4)) >= 0.05 Then       ' no esta anulada
              NUORSE& = CVS(Mid$(XX$, 117, 4))
              If NUORSE& < 1 Then ' excluye orden de servicios
                Z$ = Space$(64)
                FEX = CVI(Mid$(XX$, 5, 2))
                Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
                NPXX% = CVI(Mid$(XX$, 7, 2))
                RPP$ = ECOARCH$("ACREDOR", MKI$(NPXX%))
                If RPP$ <> "" Then
                  Call REPLA(Z$, RPP$, 11, 12)
                End If
                COEFUNI = CVS(Mid$(XX$, 63, 4))
                If COEFUNI < 0.0001 Then COEFUNI = 1
                IUNI# = CVD(Mid$(XX$, 79, 8)) / COEFUNI
                Call REPLA(Z$, "$", 25, 1)
                Call REPLA(Z$, CSTRING$(MKD$(IUNI#), 3, 10, 4, 2), 27, 10)
                If TRIM$(Mid$(XX$, 43, 12)) <> "" Then
                  Call REPLA(Z$, Mid$(XX$, 43, 12), 25, 12)
                End If
                If Z$ <> UAGRE$ Then '  para no repetir si la O/C tiene varias entregas programadas
                   List4.AddItem Z$
                   UAGRE$ = Z$
                End If
              End If
            End If
          End If
        End If
      Next KKI%
      '
      If EXISTE%(LUDAT$ + "OBSECOIT.RFS") < 1 Then
        Call COPYSTRU("OBSEOPER", "OBSECOIT")
      End If
      '
      NPLX$ = ""
      RNC$ = RECFILT$("OBSECOIT", 1, MKI$(CI%))
      URN& = ULTREG&("OBSECOIT")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("OBSECOIT", RENOTA&)
            NPLX$ = NPLX$ + Mid$(TTZZ$, 5)
          End If
        End If
      Next U&
      Text40 = RTrim$(NPLX$)
      '
      MODIFIC%(0) = 0
      YACAR%(0) = 1
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

Sub CARESPECI()
   '
   If YACAR%(1) < 1 Then
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
      YACAR%(1) = 1
      MODIFIC%(1) = 0
      '
   End If
   '
End Sub
'
Sub CARFORMULA()
   '
   If YACAR%(2) < 1 Then
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
      YACAR%(2) = 1
      MODIFIC%(2) = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub CARHORUTA()
   '
   If YACAR%(3) < 1 Then
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
        RESECO$(List3.ListCount - 1) = Left$(X$ + String$(128, 0), 128) + Left$(NUHOPROF$(CI%, NOPY%) + Space$(16), 16) + MKI$(CI%) + MKI$(NOPY%) + String$(28, 0)
        DESOPE$(List3.ListCount - 1) = ""
      Next I1&
      '
      RNC$ = RECFILT$("COSOPTER", 1, MKI$(CI%))
      URN& = ULTREG&("COSOPTER")
      '
      For U& = 1 To Len(RNC$) Step 4
        RENOTA& = CVS(Mid$(RNC$, U&, 4))
        If RENOTA& > 0 Then
          If RENOTA& <= URN& Then
            TTZZ$ = REGLEIDO$("COSOPTER", RENOTA&)
            NOPY% = CVI(Mid$(TTZZ$, 3, 2))
            For KKI% = 1 To List3.ListCount
              If CVI(Mid$(RESECO$(KKI% - 1), 3, 2)) = NOPY% Then
                Mid$(RESECO$(KKI% - 1), 145, 32) = TTZZ$
              End If
            Next KKI%
          End If
        End If
      Next U&
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
      YACAR%(3) = 1
      MODIFIC%(3) = 0
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
   RETEMODI3% = MODIFIC%(3)
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
   NPRO% = CVI(Mid$(RESECO$(List3.ListIndex), 149, 2))
   PREC# = CVD(Mid$(RESECO$(List3.ListIndex), 151, 8))
   MONEI% = CVI(Mid$(RESECO$(List3.ListIndex), 159, 2))
   If MONEI% < 1 Then MONEI% = 1
   FECOI = CVI(Mid$(RESECO$(List3.ListIndex), 161, 2))
   '
   Label48.Caption = ""
   If NPRO% > 0 Then
     PROVEX$ = AJUSTI$(ECOARCH$("ACREDOR", MKI$(NPRO%)), 48) + Str$(NPRO%)
     Label48.Caption = PROVEX$
   End If
   '
   Text37.TEXT = CSTRING$(MKD$(PREC#), 3, 11, 4, 2)
   '
   If MONEI% > 0 Then
     MONEX$ = AJUSTI$(ECOARCH$("TAMONED", Chr$(MONEI%)), 48) + Str$(MONEI%)
     Label46.Caption = MONEX$
   End If
   '
   Label63 = FECHATEX$(FECOI)
   '
   MODIFIC%(3) = RETEMODI3%
   '
End Sub
'
Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   '
   If MODIFIC%(0) > 0 Then GoSub 90: Call GRAGENITEM
   If MODIFIC%(1) > 0 Then GoSub 90: Call GRAESPECI
   If MODIFIC%(2) > 0 Or (SSTab1.Tab = 2 And GRABATODO% = 1) Then Command9.SetFocus: Call GRAFORMULA
   If MODIFIC%(3) > 0 Then GoSub 90: Call GRAHORUTA
   '
   If GRABATODO% = 1 Then
     If SSTab1.Tab = 3 Then Call GRACOSTAN
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Exit Sub
   '
90 On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
Return
   '
End Sub
'
Sub GRACOSTAN()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If DERACCE%("GRACOSTA", "Grabar Costos Standard por Articulo") < 2 Then
      Exit Sub
   End If
   '
   RETRA& = CI%
   If RETRA& <= ULTREG&("RECOSCAL") Then
      XX$ = REGLEIDO$("MASTER", RETRA& + 1)
      TIMA% = Asc(Mid$(XX$, 93, 1))
      UMED$ = AJUSTI$(Mid$(XX$, 109, 2), 2)
      FECO% = HOY(HOS$)
      YY$ = String$(128, 0)
      Call REPLA(YY$, MKI$(CI%), 1, 2)
      Call REPLA(YY$, MKI$(TIMA%), 3, 2)
      Call REPLA(YY$, UMED$, 5, 2)
      Call REPLA(YY$, MKI$(FECO%), 7, 2)
      '
      Call REPLA(YY$, MKD$(Val(Label90) + Val(Label82)), 9, 8)
      Call REPLA(YY$, MKD$(Val(Label88) + Val(Label83)), 17, 8)
      Call REPLA(YY$, MKD$(Val(Label86) + Val(Label84)), 25, 8)
      Call REPLA(YY$, MKD$(Val(Label85)), 33, 8)
      Call REPLA(YY$, MKD$(Val(Label75) + Val(Label72)), 41, 8)
      Call REPLA(YY$, MKD$(Val(Label79) + Val(Label74)), 49, 8)
      Call REPLA(YY$, MKD$(0), 57, 8)
      Call REPLA(YY$, MKD$(Val(Label80)), 65, 8)
      Call REPLA(YY$, MKD$(Val(Label70)), 73, 8)
      Call REPLA(YY$, MKD$(Val(Label65)), 81, 8)
      Call REPLA(YY$, MKD$(Val(Label96)), 89, 8)
      Call REPLA(YY$, MKD$(Val(Label103)), 97, 8)
      Call REPLA(YY$, MKD$(Val(Label112)), 105, 8)
      Call REPLA(YY$, MKD$(Val(Label116)), 113, 8)
      Call REPLA(YY$, MKS$(Val(Label117)), 121, 4)
      Call REPLA(YY$, MKS$(Val(Label119)), 125, 4)
      Call GRAREG("RECOSCAL", YY$, RETRA&)
   End If
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
   PLDE$ = Text45
   REPLAN$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CI%))
   Call REPLA(REPLAN$, MKI$(CI%), 1, 2)
   Call REPLA(REPLAN$, PLDE$, 3, 16)
   Call FILTERPUTRECORD("INDIPLAN", 1, MKI$(CI%), REPLAN$)
   Call CIERRARCH("INDIPLAN")
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
      HAYACMAS% = 1
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
      HAYACMAS% = 1
   End If
   Call REPLA(REMAS$, DE$, 5, 64)
   '
   TIMA% = Val(Text6.TEXT)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   '
   GRUIVA% = Val(Text35.TEXT)
   Call REPLA(REMAS$, Chr$(GRUIVA%), 3, 1)
   '
   CRIT% = Val(Text8.TEXT)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   '
   MONEI% = Val(Text13.TEXT)
   Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
   '
   TCOLOR$ = TRIM$(Text22.TEXT)
   Call REPLA(REMAS$, Chr$(CCOLOR%(TCOLOR$)), 96, 1)
   '
   TTALLE$ = TRIM$(Text23.TEXT)
   Call REPLA(REMAS$, Chr$(CTALLE%(TTALLE$)), 97, 1)
   '
   CCOMPO% = Val(Text25.TEXT)
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
   COU = Val(Text12.TEXT)
   Call REPLA(REMAS$, MKS$(COU), 111, 4)
   '
   FECO% = FECHANUM(Text14.TEXT)
   Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
   '
   NPRO% = Val(Text10.TEXT)
   NPROA% = CVI(Mid$(REMAS$, 123, 2))
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   PHABIT%(CI%) = (-1)
   If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
     PHABIT%(CI%) = NPRO%
   End If
   '
   GRUCO% = Val(Text16.TEXT)
   Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   '
   If Abs(Val(Text28)) > 320 Then
       Call MENSERR(24, "% Gastos de Nacionalización No Válido.")
     Else
       FCPI% = Int(100 * Val(Text28.TEXT) + 0.5)
       Call REPLA(REMAS$, MKI$(FCPI%), 127, 2)
   End If
   '
30 Call GRAREG("MASTER", REMAS$, REG&)
   Call CIERRARCH("MASTER")
   '
   If NPRO% <> NPROA% Then  ' cambio de proveedor habitual
     ZZX$ = REGBLAN$("RECAMPH")
     Call REPLA(ZZX$, MKI$(CI%), 1, 2)
     HOII% = HOY(HOS$)
     Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
     Call REPLA(ZZX$, MKI$(NPROA%), 5, 2)
     Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
     Call REPLA(ZZX$, USERTER$, 9, 24)
     Call REPLA(ZZX$, "Costos Standard", 33, 32)
     Call REGAPP("RECAMPH", ZZX$)
     Call CIERRARCH("RECAMPH")
   End If
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
      Call FRESHECHO("ECOMAST")
      '
   End If
   '
   If CI% > 1 Then     ' fuerza actualizacion de descripcion
      DDD = DESCRIT$(CI% - 1)
   End If
   Call CIERRARCH("*.*")
   '
   If EXISTE%(LUDAT$ + "OBSECOIT.RFS") < 1 Then
     Call COPYSTRU("OBSEOPER", "OBSECOIT")
   End If
   '
   URN& = ULTREG&("OBSECOIT")
   RNC$ = RECFILT$("OBSECOIT", 1, MKI$(CI%))
   KK& = (-3)
   '
   TEXTO$ = RTrim$(Text40)
   While Len(TEXTO$) > 0
     TTYY$ = MKI$(CI%) + MKI$(0) + Left$(TEXTO$, 252)
     KK& = KK& + 4
     If KK& <= Len(RNC$) - 3 Then
          RENOTA& = CVS(Mid$(RNC$, KK&, 4))
        Else
          URN& = URN& + 1
          RENOTA& = URN&
     End If
     Call GRAREG("OBSECOIT", TTYY$, RENOTA&)
     TEXTO$ = Mid$(TEXTO$, 253)
   Wend
   '
   While KK& < Len(RNC$) - 3
     KK& = KK& + 4
     RENOTA& = CVS(Mid$(RNC$, KK&, 4))
     Call GRAREG("OBSECOIT", String$(256, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH("OBSECOIT")
   '
   Call BAJALDISCO
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
   If YACAR%(1) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("MODISPEC", "Modificación de Especificaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       MODIFIC(1) = 0
       Call CARGENITEM
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
      TTYY$ = MKI$(CI%) + RETEX$(LL%)
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
   MODIFIC(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAFORMULA()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(2) < 1 Then Exit Sub
   If MODIFIC%(2) < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MODISTRU", "Modificación Fórmula de Conjunto") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Fórmula del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(2) = 0
       MODIFIC(2) = 0
       Call CARGENITEM
       GoTo 99
     End If
   End If
   '
   RFF1$ = RECFILT$("CARFORLI", 1, MKI$(CI%))
   RFF2$ = RECFILT$("CARFORLI", 4, MKI$(CI%))
   If RFF1$ <> "" Or RFF2$ <> "" Then
     If MODIFIC%(2) > 0 Then
       Call MENSERR(24, "Artículo Componente de Fórmula por Lista.\  \Esta Fórmula solo Puede Modificarse\a Partir de la Lista de Componentes, Materiales y Medidas.")
     End If
     YACAR%(2) = 0
     MODIFIC(2) = 0
     Call CARGENITEM
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
         MODIFIC(2) = 0
         SSTab1.Tab = 2
         Call CARGENITEM
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
      MODIFIC%(2) = 0
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
   If YACAR%(3) < 1 Then MODIFIC%(3) = 0
   If MODIFIC%(3) < 1 Then Exit Sub
   '
   If DERACCE%("MODIRUTA", "Modificación Hojas de Ruta") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Hoja de Ruta del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(3) = 0
       MODIFIC%(3) = 0
       Call CARGENITEM
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
   URN& = ULTREG&("COSOPTER")
   RNC$ = RECFILT$("COSOPTER", 1, MKI$(CI%))
   KK& = (-3)
   '
   For LL% = 1 To List3.ListCount
      TTYY$ = MKI$(CI%) + Mid$(RESECO$(LL% - 1), 3, 2) + Mid$(RESECO$(LL% - 1), 149, 28)
      If CVD(Mid$(TTYY$, 7, 8)) >= 0.00005 Then
        KK& = KK& + 4
        If KK& <= Len(RNC$) - 3 Then
             RENOTA& = CVS(Mid$(RNC$, KK&, 4))
             XY$ = REGLEIDO$("COSOPTER", RENOTA&)
             If CVI(Left$(XY$, 2)) <> CI% Then
               Call MENSERR(24, "Error de Grabación de Costos de Tercerización.")
               URN& = URN& + 1
               RENOTA& = URN&
             End If
           Else
             URN& = URN& + 1
             RENOTA& = URN&
        End If
        Call GRAREG("COSOPTER", TTYY$, RENOTA&)
      End If
      '
   Next LL%
   '
   While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("COSOPTER", String$(32, 0), RENOTA&)
   Wend
   '
   MODIFIC(3) = 0
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

Private Sub List4_DblClick()
   FCRP% = FOCOREP%
   FOCOREP% = 2
   CMONE$ = Mid$(List4.TEXT, 25, 2)
   If CMONE$ = "  " Or CMONE$ = "$ " Then
       Text21 = Mid$(List4.TEXT, 27, 10)
       Text14 = Left$(List4.TEXT, 8)
       MONEI% = Val(Text13.TEXT)
       If MONEI% <= 1 Then
         Text12 = ""
         FOCOREP% = 1
         Text12 = Mid$(List4.TEXT, 27, 10)
       End If
     Else
       KMONE% = 2
       For UV& = 1 To ULTREG&("TAMONED")
         XXV$ = REGLEIDO$("TAMONED", UV&)
         If Mid$(XXV$, 2, 2) = CMONE$ Then
           KMONE% = Asc(Left$(XXV$, 1))
           GoTo 19
         End If
       Next UV&
19     Text13 = TRIM$(Str$(KMONE%))
       Text12 = ""
       FOCOREP% = 1
       Text12 = Mid$(List4.TEXT, 27, 10)
       Text14 = Left$(List4.TEXT, 8)
   End If
   FOCOREP% = FCRP%
End Sub

Private Sub Picture5_DblClick()
   If CODINT%(Text1) > 0 Then
     CIXI% = CODINT%(Text1)
     Call INDEXARCH("CARFORLI", 1)
     Call CIERRARCH("CARFORLI")
     Call INDEXARCH("CARFORLI", 4)
     Call CIERRARCH("CARFORLI")
     Call FILTERFILE("CARFORLI", "", 1, MKI$(CIXI%))
     '
     If ULTREG("!CARFORLI") < 1 Then
       REBLA$ = REGBLAN$("CARFORLI")
       J& = 0
       Call EXPLOMUL(CIXI%)
       For KI% = 2 To CONTAMUL%
         ZZ$ = REBLA$
         Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
         Call REPLA(ZZ$, CODEXT$(CIK%(KI%)), 3, 15)
         Call REPLA(ZZ$, Chr$(NIVE%(KI%)), 18, 1)
         Call REPLA(ZZ$, MKI$(CIK%(KI%)), 19, 2)
         Call REPLA(ZZ$, NUPLANO$(CIK%(KI%)), 21, 16)
         Call REPLA(ZZ$, POSPLANO$(CIK%(KI%)), 37, 8)
         Call REPLA(ZZ$, DESCRIT0$(CIK%(KI%)), 45, 48)
         Call REPLA(ZZ$, MKS$(CANT(KI%)), 93, 4)
         Call REPLA(ZZ$, MKI$(0), 97, 2)
         Call REPLA(ZZ$, MKS$(0), 99, 4)
         Call REPLA(ZZ$, MKS$(0), 103, 4)
         Call REPLA(ZZ$, MKS$(0), 107, 4)
         Call REPLA(ZZ$, MKS$(1), 111, 4)   ' KILOS POR UNIDAD
         Call REPLA(ZZ$, MKS$(0), 115, 4)   ' DESPERDICIO
         Call REPLA(ZZ$, MKS$(CANT(KI%)), 119, 4)
         Call REPLA(ZZ$, MKI$(PAD%(KI%)), 123, 2)
         Call REPLA(ZZ$, Space$(36), 125, 36)
         JJ& = JJ& + 1
         Call GRAREG("!CARFORLI", ZZ$, JJ&)
       Next KI%
       Call SETULTREG("!CARFORLI", JJ&)
       Call CIERRARCH("!CARFORLI")
     Else
       For KU& = 1 To ULTREG&("!CARFORLI")
         ZZ$ = REGLEIDO$("!CARFORLI", KU&)
         CIKI% = CVI(Mid$(ZZ$, 19, 2))
         Call REPLA(ZZ$, CODEXT$(CIKI%), 3, 15)
         If TRIM$(NUPLANO$(CIKI%)) <> "" Then Call REPLA(ZZ$, NUPLANO$(CIKI%), 21, 16)
         If TRIM$(POSPLANO$(CIKI%)) <> "" Then Call REPLA(ZZ$, POSPLANO$(CIKI%), 37, 8)
         Call REPLA(ZZ$, DESCRIT0$(CIKI%), 45, 48)
         Call GRAREG("!CARFORLI", ZZ$, KU&)
       Next KU&
       Call CIERRARCH("!CARFORLI")
     End If
     '
     ATRI$ = "/NC"
     ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + SSTab1.Top - 30))
     ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + SSTab1.Left - 30))
     '
10   VTB% = VENTABU%("CARFORLI" + ATRI$)
     If VTB% >= 0 Then
       '
       HAYACMAS% = 1
       '
       VECJX$ = ""
       ' VERIFICAR NIVELES
       NVANT% = 0
       CODICREA% = 0
       For KU& = 1 To ULTREG&("!CARFORLI")
         ZZ$ = REGLEIDO$("!CARFORLI", KU&)
         NVACT% = Asc(Mid$(ZZ$, 18, 1))
         If NVACT% < 1 Then NVACT% = NVANT%
         'If NVACT% < 1 Or NVACT% > NVANT% + 1 Then
         If NVACT% > NVANT% + 1 Then
           Call MENSERR(24, "Nivel No Procesable")
           GoTo 10
         End If
         If CVS(Mid$(ZZ$, 93, 4)) < 1 Then
           Call MENSERR(24, "Falta Cantidad en Renglon " + TRIM$(Str$(KU&)))
           GoTo 10
         End If
         NVANT% = NVACT%
         '
         If Len(TRIM$(Mid$(ZZ$, 3, 15))) <= 1 Then
           If TRIM$(Mid$(ZZ$, 21, 16)) <> "" Then
             If TRIM$(Mid$(ZZ$, 37, 8)) <> "" Then
               If Len(TRIM$(Mid$(ZZ$, 21, 16))) + Len(TRIM$(Mid$(ZZ$, 37, 8))) <= 14 Then
                 NCOD$ = TRIM$(Mid$(ZZ$, 21, 16)) + "-" + TRIM$(Mid$(ZZ$, 37, 8))
                 Call REPLA(ZZ$, NCOD$, 3, 15)
                 Call GRAREG("!CARFORLI", ZZ$, KU&)
                 CODICREA% = 1
               End If
             End If
           End If
         End If
         '
       Next KU&
       '
       If CODICREA% > 0 Then
         Call COMUNI("Codigos Creados a Partir\de Plano y Posicion. Revise !")
         GoTo 10
       End If
       '
       ' ALTA PLANO Y POSICION COMPONENTES
       ATXX% = 0: CAIT% = 0: JJ& = 0: UCOCON% = 0
       For KU& = 1 To ULTREG&("!CARFORLI")
         ZZ$ = REGLEIDO$("!CARFORLI", KU&)
         CODE$ = TRIM$(Mid$(ZZ$, 3, 15))
         '
         If CODE$ = "" Then
           Call REPLA(ZZ$, String$(3, 0) + Space$(72), 18, 75)
           GoTo 19
         End If
         '
         If Len(CODE$) < 3 Then
           Call MENSERR(24, "Codigo '" + CODE$ + "'no Valido")
           GoTo 10
         End If
         '
         If CODE$ <> "" Then
           If Asc(Mid$(ZZ$, 18, 1)) < 1 Then
             Call MENSERR(24, "Nivel no Procesable para Código '" + CODE$ + "'.")
             GoTo 10
           End If
         End If
         '
         PLDE$ = TRIM$(Mid$(ZZ$, 21, 16))
         PSDE$ = TRIM$(Mid$(ZZ$, 37, 8))
         DEDE$ = TRIM$(Mid$(ZZ$, 45, 48))
         UMDE$ = "U."
         CANTID = CVS(Mid$(ZZ$, 93, 4))
         '
         Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
         Call GRAREG("!CARFORLI", ZZ$, KU&)
         '
         CODIEL% = CODINT%(CODE$)
         If CODIEL% < 1 Then
             If DEDE$ = "" Then
               Call MENSERR(24, "Imposible Alta de Codigo '" + CODE$ + "'. Falta Descripcion.")
               GoTo 10
             End If
             If UMDE$ = "" Then UMDE$ = "U."
             Call AGRECOD(CODE$, DEDE$, 1, UMDE$, OKX%)
             If OKX% < 1 Then
               Call MENSERR(24, "Error al Agregar Codigo " + CODE$ + ".")
             End If
             CODIEL% = NUMAS%
           Else
             If DEDE$ <> "" Then
               REMAS& = CODIEL% + 1
               REMAX$ = REGLEIDO$("MASTER", REMAS&)
               Call REPLA(REMAX$, DEDE$, 5, 64)
               Call GRAREG("MASTER", REMAX$, REMAS&)
             End If
         End If
         '
         Call REPLA(ZZ$, MKI$(CODIEL%), 19, 2)
         '
         RFQ$ = RECFILT$("INDIPLAN", 1, MKI$(CODIEL%))
         If RFQ$ <> "" Or PLDE$ <> "" Or PSDE$ <> "" Then
           REPLAN$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CODIEL%))
           Call REPLA(REPLAN$, MKI$(CODIEL%), 1, 2)
           Call REPLA(REPLAN$, PLDE$, 3, 16)
           Call REPLA(REPLAN$, PSDE$, 25, 8)
           Call FILTERPUTRECORD("INDIPLAN", 1, MKI$(CODIEL%), REPLAN$)
           Call CIERRARCH("INDIPLAN")
         End If
         '
19       JJ& = JJ& + 1
         Call GRAREG("!CARFORLI", ZZ$, JJ&)
         '
         If CODIEL% > 0 Then
           If CODIEL% = CIXI% Then
             If ATXX% < 1 Then
               Call MENSERR(24, "Error de Recursividad !!!\  \El Código del Subconjunto\Debe ser Distinto del Código\del Conjunto Completo.")
               ATXX% = 1
             End If
           End If
         End If
         '
         If Asc(Mid$(ZZ$, 18, 1)) = 1 Then
           CAIT% = CAIT% + 1
           TZX$ = Space$(64)
           Call REPLA(TZX$, CODEXT$(CODIEL%), 1, 16)
           Call REPLA(TZX$, FORMATNUM$(CANTID, "F12.5"), 49, 12)
           Call REPLA(TZX$, CSTRING$(Mid$(ZZ$, 123, 2), 1, 4, 0, 2), 61, 4)
           EXPLOLIN$(CAIT%) = TZX$
         End If
         '
         If CVI(Mid$(ZZ$, 97, 2)) = CODIEL% Then
           If ATXX% < 1 Then
             Call MENSERR(24, "Error de Recursividad !!!\  \El Código de la Materia Prima\Debe ser Distinto del Código\de la Pieza que Compone.")
             ATXX% = 1
           End If
         End If
         '
       Next KU&
       Call SETULTREG("!CARFORLI", JJ&)
       Call FILTERUPDATE("CARFORLI", "", 1, MKI$(CIXI%))
       Call CIERRARCH("CARFORLI")
       '
       If ATXX% > 0 Then
         Call COMUNI("Corrija Errores de Recursividad.")
         GoTo 10
       End If
       '
       For KUI% = 1 To Len(VECJX$) Step 2
         If CVI(Mid$(VECJX$, KUI%, 2)) = CIXI% Then
           Call MENSERR(24, "Imposible Grabar Segunda Fórmula\de '" + TRIM$(CODEXT$(CIXI%)) + "'.")
           GoTo 22
         End If
       Next KUI%
       '
       Call AREXPLO(CIXI%, 1, 2)
       VECJX$ = VECJX$ + MKI$(CIXI%)
       ' HASTA AQUI GRABO LA ESTRUCTURA BASICA
       '
       ' GRABA LAS ESTRUCTURAS SECUNDARIAS
22     For KU1& = 1 To ULTREG&("!CARFORLI")
         ZZ$ = REGLEIDO$("!CARFORLI", KU1&)
         NIVEBA% = Asc(Mid$(ZZ$, 18, 1))
         CODIBA% = CVI(Mid$(ZZ$, 19, 2))
         CANTBA = CVS(Mid$(ZZ$, 93, 4))
         '
         CAIT% = 0
         For KU2& = KU1& + 1 To ULTREG&("!CARFORLI")
          ZZ$ = REGLEIDO$("!CARFORLI", KU2&)
          NIVELE% = Asc(Mid$(ZZ$, 18, 1))
           If NIVELE% > 0 Then
             If NIVELE% <= NIVEBA% Then GoTo 20
             If NIVELE% = NIVEBA% + 1 Then
               CODELE% = CVI(Mid$(ZZ$, 19, 2))
               If CODELE% = CODIBA% Then
                 Call MENSERR(24, "Error de Recursividad !!!\  \El Código del Elemento\Debe ser Distinto del Código\del Conjunto que Compone.")
                 GoTo 10
               End If
               '
               CANELE = CVS(Mid$(ZZ$, 93, 4)) / CANTBA
               '
               CAIT% = CAIT% + 1
               TZX$ = Space$(64)
               Call REPLA(TZX$, CODEXT$(CODELE%), 1, 16)
               Call REPLA(TZX$, FORMATNUM$(CANELE, "F12.5"), 49, 12)
               Call REPLA(TZX$, CSTRING$(Mid$(ZZ$, 123, 2), 1, 4, 0, 2), 61, 4)
               EXPLOLIN$(CAIT%) = TZX$
             End If
           End If
         Next KU2&
         '
         For KUI% = 1 To Len(VECJX$) Step 2
           If CVI(Mid$(VECJX$, KUI%, 2)) = CODIBA% Then
             Call MENSERR(24, "Imposible Grabar Segunda Fórmula\de '" + TRIM$(CODEXT$(CODIBA%)) + "'.")
             GoTo 23
           End If
         Next KUI%
         '
20       Call AREXPLO(CODIBA%, 1, 2)
         VECJX$ = VECJX$ + MKI$(CODIBA%)
       Next KU1&
       '
23     For KU& = 1 To ULTREG&("!CARFORLI")
         ZZ$ = REGLEIDO$("!CARFORLI", KU&)
         CODIPI% = CVI(Mid$(ZZ$, 19, 2))
         NIVEPI% = Asc(Mid$(ZZ$, 18, 1))
         CODIMP% = CVI(Mid$(ZZ$, 97, 2))
         CABASI = CVS(Mid$(ZZ$, 93, 4)): If CABASI < 0.1 Then CABASI = 1
         '
         If CODIMP% > 0 Then
           If CODIPI% > 0 Then
             CAIT% = 0
             For KU2& = KU& To ULTREG&("!CARFORLI")
               ZZ$ = REGLEIDO$("!CARFORLI", KU2&)
               If KU2& > KU& And CVI(Mid$(ZZ$, 19, 2)) > 0 And Asc(Mid$(ZZ$, 18, 1)) = NIVEPI% + 1 Then
                 UNIDS = CVS(Mid$(ZZ$, 93, 4))
                 CODIMP% = CVI(Mid$(ZZ$, 19, 2))
                 GoTo 24
               End If
               If KU2& > KU& And CVI(Mid$(ZZ$, 19, 2)) > 0 Then GoTo 29
               '
               UNIDS = CVS(Mid$(ZZ$, 93, 4))
               If KU2& > KU& Then UNIDS = UNIDS * CABASI
               CODIMP% = CVI(Mid$(ZZ$, 97, 2))
               '
               If CODIMP% = CODIPI% Then
                 Call MENSERR(24, "Error de Recursividad !!!\  \El Código del Elemento\Debe ser Distinto del Código\del Conjunto que Compone.")
                 GoTo 10
               End If
               '
24             If CODIMP% > 0 Then
                 If CODIMP% <= NUMAS% Then
                   CANTID = UNIDS
                   MEDI2 = CVS(Mid$(ZZ$, 103, 4))
                     If MEDI2 > 0.05 Then CANTID = CANTID * (MEDI2 + 5) / 1000
                   MEDI3 = CVS(Mid$(ZZ$, 107, 4))
                     If MEDI3 > 0.05 Then CANTID = CANTID * (MEDI3 + 5) / 1000
                   MEDI1 = CVS(Mid$(ZZ$, 99, 4))
                   If MEDI1 > 0.05 Then
                     If MEDI2 < 1 Then
                       If MEDI3 < 1 Then
                         MEDI1 = MEDI1 + 3
                       End If
                     End If
                   End If
                     If MEDI1 > 0.05 Then CANTID = CANTID * MEDI1 / 1000
                   KILOUN = PESUNI(CODIMP%)
                     If TRIM$(UCase$(UNIMED$(CODIMP%))) = "KG" Then
                       If KILOUN > 0.0005 Then CANTID = CANTID * KILOUN
                     End If
                   PORSCRA = CVS(Mid$(ZZ$, 115, 4))
                     CANTID = CANTID * (1 + PORSCRA / 100)
                   Call REPLA(ZZ$, MKS$(KILOUN), 111, 4)
                   Call REPLA(ZZ$, MKS$(CANTID), 119, 4)
                   Call GRAREG("!CARFORLI", ZZ$, KU&)
                   '
                   If UNIDS > 0.5 Then
                     CAIT% = CAIT% + 1
                     TZX$ = Space$(64)
                     Call REPLA(TZX$, CODEXT$(CODIMP%), 1, 16)
                     Call REPLA(TZX$, FORMATNUM$(CANTID, "F12.5"), 49, 12)
                     Call REPLA(TZX$, CSTRING$(Mid$(ZZ$, 123, 2), 1, 4, 0, 2), 61, 4)
                     EXPLOLIN$(CAIT%) = TZX$
                   End If
                 End If
               End If
             Next KU2&
             '
             For KUI% = 1 To Len(VECJX$) Step 2
               If CVI(Mid$(VECJX$, KUI%, 2)) = CODIPI% Then
                 'Call MENSERR(24, "Imposible Grabar Segunda Fórmula\de '" + TRIM$(CODEXT$(CODIPI%)) + "'.")
                 GoTo 29
               End If
             Next KUI%
             '
29           Call AREXPLO(CODIPI%, CABASI, 2)
             VECJX$ = VECJX$ + MKI$(CODIPI%)
           End If
         End If
31     Next KU&
       '
       Call BAJALDISCO
       '
     End If
   End If
   '
   YACAR%(2) = 0
   Call CARFORMULA
   '
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab <> PreviousTab Then
      If InStr(SSTab1.Caption, ". . .") > 0 Then
        SSTab1.Tab = PreviousTab
        Exit Sub
      End If
      '
      If MODOEDIT% = 1 Then
        If PreviousTab = 1 Then
          SSTab1.Tab = PreviousTab
        End If
        Call MENSAEDIT
        Exit Sub
      End If
      Call GRADATITEM
      If SSTab1.Tab = 3 Then
        If LLAMACOSTO$ <> "" Then
          If Text1.TEXT = LLAMACOSTO Then Text1.TEXT = ""
          Text1.TEXT = LLAMACOSTO$
          'LLAMACOSTO$ = ""
          Exit Sub
        End If
      End If
      '
      If SSTab1.Tab = 2 Then
        If LLAMACOSTO$ <> "" Then
          If PreviousTab = 3 Then
            List3.Clear
            YACAR%(3) = 0
          End If
        End If
      End If
      '
      Call CARDATITEM
      '
   End If
End Sub

Private Sub Text1_Change()
   '
   For Index = 0 To 15: YACAR%(Index) = 0: Next Index
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
   MODIFIC%(0) = 1
End Sub

Private Sub Text10_DblClick()
   Call SELECHO("ACREDOR")
   VALONUE$ = VALACT1$("ACREDOR")
   If VALONUE$ <> "" Then
      Text10.TEXT = VALONUE$
      Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(Val(Text10.TEXT))))
   End If
End Sub


Private Sub Text10_LostFocus()
   PROVEHABI = Val(Text10.TEXT)
   If PROVEHABI < 1 Or PROVEHABI > 32767 Then
       Text10.TEXT = ""
     ElseIf ECOARCH$("ACREDOR", MKI$(Int(PROVEHABI))) = "" Then
       Text10.TEXT = ""
     Else
       Label33.Caption = ECOARCH("ACREDOR", MKI$(Int(PROVEHABI)))
   End If
End Sub

Private Sub Text11_Change()
   If List3.ListIndex >= 0 Then
     Mid$(RESECO$(List3.ListIndex), 129, 16) = AJUSTI$(Text11.TEXT, 16)
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11.TEXT)
End Sub

Private Sub Text12_Change()
   If FOCOREP% <= 1 Then
     MONEI% = Val(Text13.TEXT)
     COSUDOL = Val(Text12)
     If MONEI% > 1 Then
         VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
       Else
         FEBAS% = FECHANUM(Text14)
         FEACT% = HOY(HOS$)
         VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
     End If
     Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     '
     Text46 = ""
     KILOUN = Val(Text24)
     If KILOUN >= 0.005 Then
       VALKILO = COSUDOL / KILOUN
       Text46.TEXT = CSTRING$(MKS$(VALKILO), 4, 11, 4, 2)
     End If
     '
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text12_GotFocus()
   FOCOREP% = 1
End Sub

Private Sub Text12_LostFocus()
   Text12.TEXT = CSTRING$(MKS$(Val(Text12.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text46_Change()
   If FOCOREP% = 3 Then
     VALKILO = Val(Text46)
     KILOUN = Val(Text24)
     If KILOUN >= 0.005 Then
       COSUDO = VALKILO * KILOUN
       Text12.TEXT = CSTRING$(MKS$(COSUDO), 4, 11, 4, 2)
     End If
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text46_GotFocus()
   KILOUN = Val(Text24)
   If KILOUN < 0.005 Then
     Text46 = ""
     Text12.SetFocus
     Exit Sub
   End If
   FOCOREP% = 3
End Sub

Private Sub Text46_LostFocus()
   Text46.TEXT = CSTRING$(MKS$(Val(Text46.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text13_Change()
   VALONU = Val(Text13.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text13.TEXT = "": Exit Sub
   End If
   Label35.Caption = ECOARCH$("TAMONED", Chr$(VALONU))
   '
   MONEI% = Val(Text13.TEXT)
   If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text13_DBLCLICK()
   Call SELECHO("TAMONED")
   VALONUE$ = VALACT1$("TAMONED")
   If VALONUE$ <> "" Then
      Text13.TEXT = VALONUE$
   End If
End Sub

Private Sub Text14_Change()
   MODIFIC%(0) = 1
   If FECHANUM(Text14) > 0 Then
      CIXI% = CODINT%(Text1)
      MONEI% = Val(Text13)
      COSUDOL = Val(Text12)
      If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIXI%) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
      End If
      Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
   End If
End Sub

Private Sub Text14_DblClick()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text14.TEXT = VALONUE$
   End If
End Sub

Private Sub Text15_Change(Index As Integer)
   If Val(Text15(Index).TEXT) > 999 Then Text15(Index).TEXT = "999"
   If Val(Text15(Index).TEXT) < 0 Then Text15(Index).TEXT = "  0"
   If ESPORLOSTFOCUS% > 0 Then Exit Sub
   '
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
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text15_GotFocus(Index As Integer)
   Text15(Index).SelStart = 0
   Text15(Index).SelLength = Len(Text15(Index).TEXT)
End Sub

Private Sub Text15_LostFocus(Index As Integer)
   ESPORLOSTFOCUS% = 1
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text15(Index).TEXT)), 3, 3, 0, 2))
   If TTTT$ = "" Then TTTT$ = "0"
   While Len(TTTT$) < 3: TTTT$ = " " + TTTT$: Wend
   Text15(Index).TEXT = TTTT$
   ESPORLOSTFOCUS% = 0
End Sub

Private Sub Text17_Change(Index As Integer)
   If Val(Text17(Index).TEXT) > 59 Then Text17(Index).TEXT = "59"
   If Val(Text17(Index).TEXT) < 0 Then Text17(Index).TEXT = " 0"
   If ESPORLOSTFOCUS% > 0 Then Exit Sub
   '
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
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text17_GotFocus(Index As Integer)
   Text17(Index).SelStart = 0
   Text17(Index).SelLength = Len(Text17(Index).TEXT)
End Sub

Private Sub Text17_LostFocus(Index As Integer)
   ESPORLOSTFOCUS% = 1
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text17(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text17(Index).TEXT = TTTT$
   ESPORLOSTFOCUS% = 0
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
   MODIFIC%(3) = 1
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_LostFocus()
   Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text21_Change()
   If FOCOREP% = 2 Then
     MONEI% = Val(Text13.TEXT)
     On Error Resume Next
     COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
     On Error GoTo 0
     Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   MODIFIC%(0) = 1
End Sub

Private Sub Text21_GotFocus()
   FOCOREP% = 2
End Sub

Private Sub Text21_LostFocus()
   Text21.TEXT = CSTRING$(MKS$(Val(Text21.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text22_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text22_DblClick()
   Call SELECHO("TACOLOR")
   COLONUE$ = VALACT1$("TACOLOR")
   If COLONUE$ <> "" Then
      Text22.TEXT = COLONUE$
   End If
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_DblClick()
   Call SELECHO("TATALLES")
   TALLENUE$ = VALACT1$("TATALLES")
   If TALLENUE$ <> "" Then
      Text23.TEXT = TALLENUE$
   End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
   FOCOREP% = 1
   Call Text12_Change
End Sub

Private Sub Text24_LostFocus()
   Text24.TEXT = CSTRING$(MKS$(Val(Text24.TEXT)), 3, 10, 5, 2)
End Sub

Private Sub Text25_Change()
   VALONU = Val(Text25.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text25.TEXT = "": Exit Sub
   End If
   Label39.Caption = ECOARCH$("TACOMPO", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_DblClick()
   Call SELECHO("TACOMPO")
   VALONUE$ = VALACT1$("TACOMPO")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Text26_Change(Index As Integer)
   If Val(Text26(Index).TEXT) > 59 Then Text26(Index).TEXT = "59"
   If Val(Text26(Index).TEXT) < 0 Then Text26(Index).TEXT = " 0"
   If ESPORLOSTFOCUS% > 0 Then Exit Sub
   '
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
     If RESECO$(List3.ListIndex) <> RESECOA$ Then MODIFIC%(3) = 1
   End If
   Call CALCOSOPE
End Sub

Private Sub Text26_GotFocus(Index As Integer)
   Text26(Index).SelStart = 0
   Text26(Index).SelLength = Len(Text26(Index).TEXT)
End Sub

Private Sub Text26_LostFocus(Index As Integer)
   ESPORLOSTFOCUS% = 1
   TTTT$ = TRIM$(CSTRING$(MKS$(Val(Text26(Index).TEXT)), 3, 2, 0, 2))
   While Len(TTTT$) < 2: TTTT$ = "0" + TTTT$: Wend
   Text26(Index).TEXT = TTTT$
   ESPORLOSTFOCUS% = 1
End Sub

Private Sub Text28_Change()
   MONEI% = Val(Text13.TEXT)
   If MONEI% <= 1 Then Text28 = ""
   If FOCOREP% <= 1 Then
       COSUDOL = Val(Text12)
       If MONEI% > 1 Then
           VALPESI = COSUDOL * TICAMONE(MONEI%) * (1 + Val(Text28) / 100)
         Else
           FEBAS% = FECHANUM(Text14)
           FEACT% = HOY(HOS$)
           VALPESI = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
       End If
       Text21.TEXT = CSTRING$(MKS$(VALPESI), 4, 11, 4, 2)
     ElseIf FOCOREP% = 2 Then
       On Error Resume Next
       COSUDOL = Val(Text21) / (TICAMONE(MONEI%) * (1 + Val(Text28) / 100))
       On Error GoTo 0
       Text12.TEXT = CSTRING$(MKS$(COSUDOL), 4, 11, 4, 2)
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_LostFocus()
  If Val(Text13) <= 1 Then Text28 = ""
  Text28 = AJUSTD$(Format$(Val(Text28), "###0.00"), 8)
End Sub

Private Sub Text29_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 75, 4) = MKS(Val(Text29.TEXT))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
   Call CALCOSOPE
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
   MODIFIC%(3) = 1
   Call CALCOSOPE
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
   MODIFIC%(3) = 1
   Call CALCOSOPE
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
   MODIFIC%(3) = 1
   Label51.Caption = ECOARCH$("CATOPER", Chr$(VALONU))
   Call CALCOSOPE
End Sub

Private Sub Text33_DblClick()
   Call SELECHO("CATOPER")
   VALONUE$ = VALACT1$("CATOPER")
   If VALONUE$ <> "" Then
      Text33.TEXT = VALONUE$
   End If
End Sub

Private Sub Text35_Change()
   VALONU = Val(Text35.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text35.TEXT = "": Exit Sub
   End If
   Label57.Caption = ECOARCH$("GRUPIVA", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text35_DblClick()
   Call SELECHO("GRUPIVA")
   VALONUE$ = VALACT1$("GRUPIVA")
   If VALONUE$ <> "" Then
      Text35.TEXT = VALONUE$
   End If
End Sub

Private Sub Text36_Change()
   Label68 = TRIM$(CSTRING$(MKD$(Val(Label80) * Val(Text36) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text36_GotFocus()
   Text36.SelStart = 0
   Text36.SelLength = Len(Text36.TEXT)
End Sub

Private Sub Text36_LostFocus()
   If Val(Text36) > 999.9 Then Text36 = "999.9"
   Text36 = CSTRING$(MKS$(Val(Text36)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENFAB", Text36.TEXT)
End Sub

Private Sub Text37_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 151, 8) = MKD(Val(Text37))
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MONEXI% = Val(Mid$(Label46, 49, 6))
   TERPESOS = Val(Text37) * TICAMONE(MONEXI%)
   Label52 = CSTRING$(MKS$(TERPESOS), 3, 11, 4, 2)
   If FECHANUM(Label63) < 1 Then
     Label63 = FECHATEX$(HOY(HOS$))
   End If
   MODIFIC%(3) = 1
End Sub

Private Sub Text37_LostFocus()
   Text37.TEXT = CSTRING$(MKS$(Val(Text37.TEXT)), 3, 11, 4, 2)
End Sub

Private Sub Text38_Change()
   Label77 = TRIM$(CSTRING$(MKD$(Val(Label65) * Val(Text38) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text38_GotFocus()
   Text38.SelStart = 0
   Text38.SelLength = Len(Text38.TEXT)
End Sub

Private Sub Text38_LostFocus()
   If Val(Text38) > 999.9 Then Text38 = "999.9"
   Text38 = CSTRING$(MKS$(Val(Text38)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENADM", Text38.TEXT)
End Sub

Private Sub Text39_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label107 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text39) / 100), 3, 12, DECICALCOS%, 2))
   '
   Label119 = ""
   If Val(Text39) > 0.005 Then
     PORMA# = 100 * Val(Label117) / Val(Text39)
     If Abs(PORMA#) <= 999.9 Then
       Label119 = TRIM$(CSTRING$(MKD$(PORMA#), 3, 5, 1, 2))
     End If
   End If
   Label119.ForeColor = RGB(0, 0, 0)
   If Val(Label119) < 100 Then Label119.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Text39_GotFocus()
   Text39.SelStart = 0
   Text39.SelLength = Len(Text39.TEXT)
End Sub

Private Sub Text39_LostFocus()
   If Val(Text39) > 99.9 Then Text39 = "99.9"
   Text39 = CSTRING$(MKS$(Val(Text39)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORUTILI", Text39.TEXT)
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

Private Sub Text40_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text41_Change()
   Label93 = TRIM$(CSTRING$(MKD$(Val(Label85) * Val(Text41) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text41_GotFocus()
   Text41.SelStart = 0
   Text41.SelLength = Len(Text41.TEXT)
End Sub

Private Sub Text41_LostFocus()
   If Val(Text41) > 999.9 Then Text41 = "999.9"
   Text41 = CSTRING$(MKS$(Val(Text41)), 3, 5, 1, 2)
   Call GRACONTROL("ANACOSTO", "GAGENMAT", Text41.TEXT)
End Sub

Private Sub Text42_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label110 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text42) / 100), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text42_GotFocus()
   Text42.SelStart = 0
   Text42.SelLength = Len(Text42.TEXT)
End Sub

Private Sub Text42_LostFocus()
   If Val(Text42) > 99.9 Then Text42 = "99.9"
   Text42 = CSTRING$(MKS$(Val(Text42)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORCOMIS", Text42.TEXT)
End Sub

Private Sub Text43_Change()
   PREVEN# = Val(Label105) / (1 - (Val(Text39) + Val(Text42) + Val(Text43)) / 100)
   Label103 = TRIM$(CSTRING$(MKD$(PREVEN#), 3, 12, DECICALCOS%, 2))
   Label104 = TRIM$(CSTRING$(MKD$(PREVEN# * Val(Text43) / 100), 3, 12, DECICALCOS%, 2))
   Label116 = ""
   If Val(Label114) >= 0.00005 Then Label116 = TRIM$(CSTRING$(MKD$(Val(Label112) * (1 - Val(Text42) / 100 - Val(Text43) / 100) - Val(Label105)), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text43_GotFocus()
   Text43.SelStart = 0
   Text43.SelLength = Len(Text43.TEXT)
End Sub

Private Sub Text43_LostFocus()
   If Val(Text43) > 99.9 Then Text43 = "99.9"
   Text43 = CSTRING$(MKS$(Val(Text43)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORIMPUE", Text43.TEXT)
End Sub

Private Sub Text44_Change()
   Label115 = TRIM$(CSTRING$(MKD$(Val(Label114) * Val(Text44) / 100), 3, 12, DECICALCOS%, 2))
End Sub

Private Sub Text44_GotFocus()
   Text44.SelStart = 0
   Text44.SelLength = Len(Text41.TEXT)
End Sub

Private Sub Text44_LostFocus()
   If Val(Text44) > 99.9 Then Text42 = "99.9"
   Text44 = CSTRING$(MKS$(Val(Text44)), 3, 4, 1, 2)
   Call GRACONTROL("ANACOSTO", "PORDESCU", Text44.TEXT)
End Sub

Private Sub Text45_Change()
   MODIFIC(0) = 1
End Sub

Private Sub Text5_Change()
   If List3.ListIndex >= 0 Then
     VDEF$ = RESECO$(List3.ListIndex)
     Mid$(VDEF$, 5, 48) = AJUSTI$(Text5.TEXT, 48)
     RESECO$(List3.ListIndex) = VDEF$
   End If
   MODIFIC%(3) = 1
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

Private Sub Text16_Change()
   VALONU = Val(Text16.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text16.TEXT = "": Exit Sub
   End If
   Label31.Caption = ECOARCH$("GRUCOVEN", MKI$(Int(VALONU)))
   MODIFIC%(0) = 1
End Sub

Private Sub Text16_DblClick()
   Call SELECHO("GRUCOVEN")
   VALONUE$ = VALACT1$("GRUCOVEN")
   If VALONUE$ <> "" Then
      Text16.TEXT = VALONUE$
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
   Call CALCOSOPE
   MODIFIC%(3) = 1
   '
End Sub

Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Len(Text7.TEXT)
End Sub

Private Sub Text8_Change()
   VALONU = Val(Text8.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text8.TEXT = "": Exit Sub
   End If
   Label32.Caption = ECOARCH$("TACRIRE", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text8_DblClick()
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
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
   MODIFIC(3) = 1
   Call CALCOSOPE
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
   HAYACMAS% = 1
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then Exit Sub
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
   DESCRINUE$ = TRIM$(Mid$(OBX$, 5, 64))
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
   Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
End Sub

Sub AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
   '
   Screen.MousePointer = 11
   '
   NMSI% = NUMAS% + 1
   '
   OBX$ = REGBLAN$("MASTER")
   Call REPLA(OBX$, MKI$(NMSI%), 1, 2)
   Call REPLA(OBX$, CODINUE$, 77, 16)
   Call REPLA(OBX$, DESCRINUE$, 5, 64)
   Call REPLA(OBX$, Chr$(TIMA%), 93, 1)
   Call REPLA(OBX$, UNID$, 109, 2)
   '
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
   Call FILTERDELETE("ESTRULAR", 1, MKI$(CODEST%))
   J& = 0
   For I& = 1 To ULTREG&("ESTRULAR")
      EST3$ = REGLEIDO$("ESTRULAR", I&)
      CIV% = CVI(Left$(EST3$, 2))
      If CIV% = CODORI% Then
        Call REPLA(EST3$, MKI$(CODEST%), 1, 2)
        Call REGAPP("ESTRULAR", EST3$)
      End If
   Next I&
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
   Screen.MousePointer = 1
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
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Indice Restaurado.")
              HAYINCO% = 1
          Else
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Control O.K.")
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
                Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                Screen.MousePointer = 11
                cix% = II&
                Call REPLA(X$, MKI$(cix%), 1, 2)
                Call GRAREG("MASTER", X$, II& + 1)
            End If
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
                Call COMUNI("Codigo '" + TRIM$(CODL$) + "' Duplicado o no Válido.\  \Reemplazado por '" + TRIM$(CODNUE$) + "'.")
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
            Call COMUNI("No se han Detectado Inconsistencias\en la Revisión del Maestro de Items.")
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
        Call CONECRUN("FLSETUP/ANACOSTO", "")
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
        Call REGAPP("TAMONED", Chr$(3) + AJUSTI$("Coef.Est.Ref.", 27) + MKS$(1))
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
   CI12% = CODINT%(Text1)
   If CI12% > 0 Then
       CI1% = CI12%: CI2% = CI12%
     Else
       XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
       XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   End If
   '
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
   cod1$ = TRIM$(CODEXT$(CI1%)): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2))
   cod2$ = TRIM$(CODEXT$(CI2%)): If TRIM$(cod2$) = "" Then Exit Sub
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
Sub LIMPLOSI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Implosión " + DESCRIT$(CI%)
   IMPLOSI.Show
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
             If STATUIT%(CICONJU%) >= 0 Then   '\\\OT-05-0838-WAR-21/11/05///
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
     End If
   Next IK&
   '
99 Screen.MousePointer = 1
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
    If CI1% = CI2% Then
      If CI1% = CIACT% Then
        RFF$ = RECFILT$("CARFORLI", 1, MKI$(CI1%))
        If Len(RFF$) >= 4 Then
          Call LIFORLI
          Exit Sub
        End If
      End If
    End If
    
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
                       Call REPLA(XX$, MKS$(COSCALCU(CIK%(I%), MODOCO%)), 93, 4)
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
Sub LIFORLI()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    '
    JJ& = 0
    REBLA$ = REGBLAN$("LIFORLI")
    '
    JJ& = JJ& + 1
    YY$ = REBLA$
      JI% = JJ&
    Call REPLA(YY$, MKI$(JI%), 1, 2)
      CANTI = 1
    Call REPLA(YY$, MKS$(1), 3, 4)
      NIVIEL% = 0: GRUPO% = 0
      CODIEL% = CIACT%
      NPLA$ = NUPLANO$(CIACT%)
      NPOS$ = "CONJ"
      NDES$ = DESCRIT0$(CIACT%)
    Call REPLA(YY$, MKI$(CODIEL%), 7, 2)
    Call REPLA(YY$, NPLA$, 9, 16)
    Call REPLA(YY$, NPOS$, 25, 8)
    Call REPLA(YY$, NDES$, 33, 40)
    Call REPLA(YY$, MKI$(GRUPO%), 73, 2)
    Call REPLA(YY$, Space$(48), 81, 48)
    Call REPLA(YY$, MKI$(0), 129, 2)
    Call REPLA(YY$, "CONJ", 157, 36)
    Call GRAREG("LIFORLI", YY$, JJ&)
    '
    RFF$ = RECFILT$("CARFORLI", 1, MKI$(CIACT%))
    For KU& = 1 To Len(RFF$) Step 4
      RECA& = CVS(Mid$(RFF$, KU&, 4))
      ZZ$ = REGLEIDO$("CARFORLI", RECA&)
      '
      JJ& = JJ& + 1
      YY$ = REBLA$
        JI% = JJ&
      Call REPLA(YY$, MKI$(JI%), 1, 2)
        CANTI = CVS(Mid$(ZZ$, 93, 4))
      Call REPLA(YY$, MKS$(CANTI), 3, 4)
        NIVIEL% = Asc(Mid$(ZZ$, 18, 1))
        CODIEL% = CVI(Mid$(ZZ$, 19, 2))
        NPLA$ = TRIM$(Mid$(ZZ$, 21, 16))
        NPOS$ = TRIM$(Mid$(ZZ$, 37, 8))
        NDES$ = Left$(".                     ", 2 * NIVIEL%) + TRIM$(Mid$(ZZ$, 45, 48))
      Call REPLA(YY$, MKI$(CODIEL%), 7, 2)
      Call REPLA(YY$, NPLA$, 9, 16)
      Call REPLA(YY$, NPOS$, 25, 8)
      Call REPLA(YY$, NDES$, 33, 40)
        CODIMP% = CVI(Mid$(ZZ$, 97, 2))
      Call REPLA(YY$, DESCRIT0$(CODIMP%), 81, 48)
      Call REPLA(YY$, MKI$(CODIMP%), 129, 2)
        CANTID = CANTI
        If CODIMP% < 1 Then CANTID = 0
        SONKILOS% = 0
        MEDI2 = CVS(Mid$(ZZ$, 103, 4))
          If MEDI2 > 0.05 Then CANTID = CANTID * (MEDI2 + 5) / 1000: SONKILOS% = 1
        MEDI3 = CVS(Mid$(ZZ$, 107, 4))
          If MEDI3 > 0.05 Then CANTID = CANTID * (MEDI3 + 5) / 1000: SONKILOS% = 1
        MEDI1 = CVS(Mid$(ZZ$, 99, 4))
          If MEDI1 > 0.05 Then CANTID = CANTID * (MEDI1 + 3) / 1000: SONKILOS% = 1
        KILOUN = PESUNI(CODIMP%)
          If SONKILOS% > 0 Then
            If KILOUN > 0.0005 Then
              CANTID = CANTID * KILOUN
            End If
          End If
        PORSCRA = CVS(Mid$(ZZ$, 115, 4))
          CANTID = CANTID * (1 + PORSCRA / 100)
      Call REPLA(YY$, MKS$(MEDI1), 131, 4)
      Call REPLA(YY$, MKS$(MEDI2), 135, 4)
      Call REPLA(YY$, MKS$(MEDI3), 139, 4)
      Call REPLA(YY$, MKS$(KILOUN), 143, 4)
      Call REPLA(YY$, MKS$(PORSCRA), 147, 4)
      If SONKILOS% > 0 Then
        Call REPLA(YY$, MKS$(CANTID), 151, 4)
      End If
      Call REPLA(YY$, MKI$(DEPOPRE%(CODIMP%)), 155, 2)
        OBSER$ = TRIM$(Mid$(ZZ$, 125, 36))
      Call REPLA(YY$, OBSER$, 157, 36)
      '
      Call GRAREG("LIFORLI", YY$, JJ&)
    Next KU&
    '
    Call SETULTREG("LIFORLI", JJ&)
    Call CIERRARCH("LIFORLI")
    '
    Call HEADERS("LIFORLI", "")
    Call HEADERS("LIFORLI", "Conjunto: " + TRIM$(Text3))
    '
    Call FINAL("*LIFORLI")
    '
End Sub

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

Sub IMPORFORMUL()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   If COMALTER%("Forma de Importación:\\Completa\Solo Fórmulas") = 2 Then GoTo 51
   '
   Label58.Visible = True
   IKK& = 0
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 15
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
10   If EOF(N1%) Then
       Close #N1%
       GoTo 15
     End If
     '
     Line Input #N1%, AA$
     IKK& = IKK& + 1: Label58.Caption = TRIM$(Str$(IKK&))
     Label58.Refresh
     COD$ = TRIM$(Mid$(AA$, 1, 15))
     DEX$ = TRIM$(Mid$(AA$, 21, 55))
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
11   COD$ = TRIM$(Mid$(AA$, 76, 15))
     DEX$ = TRIM$(Mid$(AA$, 92, 51))
     UNI$ = "U."
     COSUNIX$ = TRIM$(Mid$(AA$, 149, 10))
     Call REPLACAR(COSUNIX$, ",", ".")
     COU = Val(COSUNIX$)
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 12
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
12   REGMA& = KKI% + 1
     REMAS$ = REGLEIDO$("MASTER", REGMA&)
     If CVS(Mid$(REMAS$, 111, 4)) < 0.0001 Then
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
     End If
     '
     GoTo 10
     '
15 Next UI%
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 65
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 15))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR$ = TRIM$(Mid$(AA$, 76, 15))
     If TRIM$(CODPAR$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 143, 6))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 Next UI%
   '
   PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub
'

Private Sub Text9_GotFocus(Index As Integer)
   Text9(Index).SelStart = 0
   Text9(Index).SelLength = Len(Text9(Index).TEXT)
End Sub

Sub IMPORARCON()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   If COMALTER%("Forma de Importación:\\Completa\Solo Fórmulas") = 2 Then GoTo 51
   '
   Label58.Visible = True
   IKK& = 0
   '
     ARIMPO$ = "A:\MAESTRO.PRN"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
10   If EOF(N1%) Then
       Close #N1%
       GoTo 19
     End If
     '
     Line Input #N1%, AA$
     IKK& = IKK& + 1: Label58.Caption = TRIM$(Str$(IKK&))
     Label58.Refresh
     COD$ = TRIM$(Mid$(AA$, 1, 12))
     DEX$ = TRIM$(Mid$(AA$, 13, 41))
     UNI$ = TRIM$(Mid$(AA$, 64, 4))
     COSUNIX$ = TRIM$(Mid$(AA$, 83, 10))
     COSUN$ = REPLACAR(COSUNIX$, ",", ".")
     COU = Val(COSUN$)
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = String$(128, 0)
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
11 GoTo 10
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   ARIMPO$ = "A:\FORMULAS.PRN"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 13))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR$ = TRIM$(Mid$(AA$, 14, 15))
     If TRIM$(CODPAR$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 34, 10))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub

 Sub RECACOSVIE()
     '
     Screen.MousePointer = 11
     QQQ% = 1024
     Dim DP%(16384), DQ%(16384)
     Dim COMAQ$(1024), CUHOR(1024)
     '
     CAMAQ% = ULTREG&("PADMAQ")
     For UI& = 1 To ULTREG&("PADMAQ")
        R0$ = REGLEIDO$("PADMAQ", UI&)
        COMAQ$(UI&) = TRIM$(Left$(R0$, 6))
        CUHOR(UI&) = CVS(Mid$(R0$, 93, 4))
     Next UI&
     Call CIERRARCH("PADMAQ")
     '
62   HOI% = HOY(HO$)
AAA = CODINT%("12710001")
     NVI% = NUINV%
     NMM% = NUMAS%
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     '
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        R0$ = REGLEIDO$("ESTRUNUE", I&)
        R2$ = Mid$(R0$, 3, 120)
        CI% = CVI(Left$(R0$, 2))
        If CI% > 0 Then
            If CI% <= NMM% Then
                For K% = 1 To 120 Step 8
                    DPSX$ = Mid$(R2$, K%, 8)
                    If CVI(Left$(DPSX$, 2)) < 1 Or CVI(Left$(DPSX$, 2)) > NMM% Then GoTo 105
                    If CVS(Mid$(DPSX$, 3)) >= 0.0005 Then
                        DP%(CI%) = 1
                        DQ%(CI%) = 1
                        GoTo 105
                    End If
                Next K%
            End If
        End If
105  Next I&
     '
110  N9% = FILEOPEN%(LU$ + "COCALCU.DAT", 0, 128)
     'OPEN "R",#N9%,LU$+"COCALCU.DAT",128
     'FIELD #N9%,2 AS Q1$,2 AS Q3$,2 AS Q4$,4 AS Q11$,4 AS Q12$,4 AS Q13$,4 AS Q21$,4 AS Q22$,4 AS Q23$,4 AS Q31$,4 AS Q32$,4 AS Q33$,4 AS Q41$,4 AS Q42$,4 AS Q43$,74 AS Q99$
     Dim Q0 As String * 128
     N19% = FILEOPEN%(LU$ + "MODIREC.DAT", 0, 128)
     Dim Q9 As String * 128
     For II% = 1 To NMM%
        Get #N9%, II%, Q0$
        LSet Q0$ = String$(128, 0)
        Put #N9%, II%, Q0$
        Get #N19%, II%, Q9$
        LSet Q9$ = String$(127, 0) + Chr$(1)
        Put #N19%, II%, Q9$
     Next II%
115  CANA& = 0

     ' CARGANDO COSTOS DE SERVICIOS SUMINISTRADOS POR TERCEROS
     N6% = FILEOPEN(LU$ + "LPREPRO.DAT", 0, 48)
     Dim P0 As String * 48
     FIN& = ULTREG&("LPREPRO")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        P0$ = REGLEIDO$("LPREPRO", I&)
125     CI% = CVI(Mid$(P0$, 3, 2))
        COU = CVS(Mid$(P0$, 37, 4))
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N9%, CI%, Q0$
                RQ0$ = Q0$
                V31 = CVS(Mid$(Q0$, 31, 4)) + COU
                V33 = CVS(Mid$(Q0$, 39, 4)) + COU
                V41 = CVS(Mid$(Q0$, 43, 4)) + COU
                V43 = CVS(Mid$(Q0$, 51, 4)) + COU
                Call REPLA(RQ0$, MKS$(V31), 31, 4)
                Call REPLA(RQ0$, MKS$(V33), 39, 4)
                Call REPLA(RQ0$, MKS$(V41), 43, 4)
                Call REPLA(RQ0$, MKS$(V43), 51, 4)
                '
                Q99$ = Mid$(RQ0$, 55, 74)
                POIN$ = ""
                For U% = 1 To 74 Step 2
                    U1% = CVI(Mid$(Q99$, U%, 2))
                    If U1% > 0 Then
                        POIN$ = POIN$ + MKI$(U1%)
                    End If
                Next U%
                POIN$ = POIN$ + MKI$(KI%) + String$(72, 0)
                Call REPLA(RQ0$, POIN$, 55, 74)
                '
                LSet Q0$ = RQ0$
                Put #N9%, CI%, Q0$
            End If
        End If
     Next I&
     '
     'CARGANDO COSTOS DE MANO DE OBRA DIRECTA
     FIN& = ULTREG&("SECOPER")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        SCP$ = REGLEIDO$("SECOPER", I&)
        CI% = CVI(Left$(SCP$, 2))
        MAQUI$ = TRIM$(Mid$(SCP$, 53, 6))
        CUOTHOR = 0
        'If TRIM$(MAQUI$) <> "" Then
           For KK% = 1 To CAMAQ%
              If COMAQ$(KK%) = MAQUI$ Then
                 CUOTHOR = CUHOR(KK%)
              End If
           Next KK%
        'End If
        '
        LOTESTA = LOTESTAN(CI%)
        If LOTESTA < 1 Then LOTESTA = 1
        CALOTES = 1
        MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
        MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
        MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
        MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
        MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
        MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
        HORAPROD = MINLOTE / 60
        CAOPS = CVS(Mid$(SCP$, 83, 4))
        CATOP% = Asc(Mid$(SCP$, 87, 1))
        CUOTOPE = VALHORA(CATOP%)
        COU = (CUOTHOR + CUOTOPE * CAOPS) * HORAPROD / LOTESTA
        '
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N19%, CI%, Q9$
                RQ0$ = Q9$
                V61 = CVS(Mid$(Q9$, 3, 4)) + COU
                V63 = CVS(Mid$(Q9$, 11, 4)) + COU
                Call REPLA(RQ0$, MKS$(V61), 3, 4)
                Call REPLA(RQ0$, MKS$(V63), 11, 4)
                LSet Q9$ = RQ0$
                Put #N19%, CI%, Q9$
            End If
        End If
     Next I&
     '
     ' TRANSFIRIENDO COSTOS DE REPOSICION DE ITEMS SIN EXPLOSION
     FIN& = NUMAS%
     Screen.MousePointer = 11
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        M0$ = REGLEIDO$("MASTER", I& + 1)
        CI% = CVI(Left$(M0$, 2))
        UNI$ = Mid$(M0$, 109, 2)
        Get #N9%, CI%, Q0$
        RQ0$ = Q0$
        Call REPLA(RQ0$, MKI$(CI%), 1, 2)
        Call REPLA(RQ0$, UNI$, 3, 2)
        Call REPLA(RQ0$, "$ ", 5, 2)
        LSet Q0$ = RQ0$
        Put #N9%, CI%, Q0$
        '
        If DP%(CI%) = 0 Then
            COU = CVS(Mid$(M0$, 111, 4))
            Get #N19%, CI%, Q9$
            If CVS(Mid$(Q9$, 3, 4)) > 0.000005 Then
               COU = 0 ' si tiene M.Obra no puede tener Costo de Materiales
            End If
            MONE$ = "$"
            If Asc(Mid$(M0$, 95, 1)) = 2 Then MONE$ = "U$"
            If Asc(Mid$(M0$, 95, 1)) = 3 Then MONE$ = "DM"
            '
            Call REPLA(RQ0$, MONE$, 5, 2)
            If UNI$ = "U." Then
                 Call REPLA(RQ0$, MKS$(COU), 19, 4)
                 Call REPLA(RQ0$, MKS$(COU), 27, 4)
               Else
                 Call REPLA(RQ0$, MKS$(COU), 7, 4)
                 Call REPLA(RQ0$, MKS$(COU), 15, 4)
            End If
            Call REPLA(RQ0$, MKS$(COU), 43, 4)
            Call REPLA(RQ0$, MKS$(COU), 51, 4)
            LSet Q0$ = RQ0$
            Put #N9%, CI%, Q0$

            If COU > 0.00005 Then
                CANA& = CANA& + 1
                RW0$ = String$(96, 0)
                Call REPLA(RW0$, MKI$(CI%), 1, 2)
                CIXI% = CI%
                Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                CIXI% = CI%
                Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                Call REPLA(RW0$, MONE$, 67, 2)
                Call REPLA(RW0$, MKS$(COU), 69, 4)
                Call REPLA(RW0$, MKS$(1), 73, 4)
                Call REPLA(RW0$, MKS$(COU), 77, 4)
                Call REPLA(RW0$, MKS$(0), 81, 4)
                Call REPLA(RW0$, MKS$(0), 85, 4)
                Call REPLA(RW0$, MKS$(COU), 89, 4)
                Call REPLA(RW0$, MKS$(0), 93, 4)
                '
                Call GRAREG("ANACOS", RW0$, CANA&)
            End If

130         If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                POIN$ = Mid$(Q0$, 55, 74)
                For U% = 1 To Len(POIN$) Step 2
                    POINI% = CVI(Mid$(POIN$, U%, 2))
                    If POINI% > 0 Then
                        CIXI% = CI%
                        Get #N6%, POINI%, P0$
                        DE$ = TRIM$(Mid$(P0$, 5, 30))
                        MN$ = Mid$(P0$, 35, 2)
                        COU = CVS(Mid$(P0$, 37, 4))
                        CANA& = CANA& + 1
                        '
                        RW0$ = String$(96, 0)
                        Call REPLA(RW0$, MKI$(CI%), 1, 2)
                        CIXI% = CI%
                        Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                        CIXI% = CI%
                        Call REPLA(RW0$, DE$, 19, 48)
                        Call REPLA(RW0$, MN$, 67, 2)
                        Call REPLA(RW0$, MKS$(COU), 69, 4)
                        Call REPLA(RW0$, MKS$(1), 73, 4)
                        Call REPLA(RW0$, MKS$(0), 77, 4)
                        Call REPLA(RW0$, MKS$(0), 81, 4)
                        Call REPLA(RW0$, MKS$(COU), 85, 4)
                        Call REPLA(RW0$, MKS$(COU), 89, 4)
                        Call REPLA(RW0$, MKS$(0), 93, 4)
                        Call GRAREG("ANACOS", RW0$, CANA&)
                        '
                    End If
                Next U%
            End If
        End If
     Next I&
     ' RE-CALCULANDO COSTOS STANDARD DE ITEMS SEGUN EXPLOSION

140  RECA% = 0
     FIN& = ULTREG&("ESTRUNUE")
     Screen.MousePointer = 11
     For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         R0$ = Left$(REGLEIDO$("ESTRUNUE", I&), 122)
         If I& <= ULTREG&("ESTRUDOS") Then
              R0$ = R0$ + REGLEIDO$("ESTRUDOS", I&)
            Else
              R0$ = R0$ + String$(256, 0)
         End If
         CI% = CVI(Left$(R0$, 2))
         DPS$ = Mid$(R0$, 3)
         If CI% < 1 Or CI% > NUMAS% Then GoTo 199

142      If DP%(CI%) = 0 Then GoTo 199                          ' ya fue calculado

         CIXI% = CI%
         CABAS = 1: SUMP = 0: SUMF = 0: SUS3 = 0: SUMO = 0
145      For K% = 1 To 376 Step 8
             DPSX$ = Mid$(DPS$, K%, 8)
             CI1% = CVI(DPSX$)
             If CI1% < 1 Or CI1% > NUMAS% Then GoTo 149    ' terminar
             '
             If CI1% = CI% Then
                If COMALTER%("Error de Recursividad en la Fórmula\Correspondiente a " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".\\Cancelar\Continuar") <> 2 Then
                   Call GRACONTROL("COSTOS", "FECALCOS", "00/00/00")
                   Call GRACONTROL("COSTOS", "HOCALCOS", "00:00:00")
                   FORCOSTO.Label34.Caption = " " + HOS$
                   FORCOSTO.Label34.Refresh
                   FORCOSTO.Label35.Caption = " " + Time$
                   FORCOSTO.Label35.Refresh
                   GoTo 3800
                End If
             End If
             '
             If DP%(CI1%) > 0 Then GoTo 199                ' aun no fue calculado
             CANT7 = CABAS * CVS(Mid$(DPSX$, 3))
             Get #N9%, CI1%, Q0$
             Get #N19%, CI1%, Q9$
             SUMP = SUMP + CANT7 * CVS(Mid$(Q0$, 15, 4))
             SUMF = SUMF + CANT7 * CVS(Mid$(Q0$, 27, 4))
             SUS3 = SUS3 + CANT7 * CVS(Mid$(Q0$, 39, 4))
             MOOX = CVS(Mid$(Q9$, 11, 4))
             SUMO = SUMO + CANT7 * CVS(Mid$(Q9$, 11, 4))
'BBBB$ = CODEXT$(CI1%)
         Next K%
         '
149      Get #N9%, CI%, Q0$
         RQ0$ = Q0$
         Get #N19%, CI%, Q9$
         RQ9$ = Q9$
         '
         V12 = CVS(Mid$(Q0$, 11, 4)) + SUMP
         V13 = CVS(Mid$(Q0$, 15, 4)) + SUMP
         V22 = CVS(Mid$(Q0$, 23, 4)) + SUMF
         V23 = CVS(Mid$(Q0$, 27, 4)) + SUMF
         V32 = CVS(Mid$(Q0$, 35, 4)) + SUS3
         V33 = CVS(Mid$(Q0$, 39, 4)) + SUS3
         V42 = CVS(Mid$(Q0$, 47, 4)) + SUMP + SUMF + SUS3
         V43 = CVS(Mid$(Q0$, 51, 4)) + SUMP + SUMF + SUS3
         V62 = CVS(Mid$(Q9$, 7, 4)) + SUMO
         V63 = CVS(Mid$(Q9$, 11, 4)) + SUMO
         '
         Call REPLA(RQ0$, MKS$(V12), 11, 4)
         Call REPLA(RQ0$, MKS$(V13), 15, 4)
         Call REPLA(RQ0$, MKS$(V22), 23, 4)
         Call REPLA(RQ0$, MKS$(V23), 27, 4)
         Call REPLA(RQ0$, MKS$(V32), 35, 4)
         Call REPLA(RQ0$, MKS$(V33), 39, 4)
         Call REPLA(RQ0$, MKS$(V42), 47, 4)
         Call REPLA(RQ0$, MKS$(V43), 51, 4)
         '
         Call REPLA(RQ9$, MKS$(V62), 7, 4)
         Call REPLA(RQ9$, MKS$(V63), 11, 4)
         '
         LSet Q0$ = RQ0$
         Put #N9%, CI%, Q0$
         '
         LSet Q9$ = RQ9$
         Put #N19%, CI%, Q9$
         QBAS = 1: CIX0% = CI%
         Call EXPLOMUL(CI%)
         Screen.MousePointer = 11
         For J% = 1 To CONTAMUL%
             CII% = CIK%(J%)
             If CII% > 0 Then
                If CII% <= NUMAS% Then
                   Get #N9%, CII%, Q0$
                   If DQ%(CII%) = 0 Then
                      COU = CVS(Mid$(Q0$, 7, 4)) + CVS(Mid$(Q0$, 19, 4))
                      CANA& = CANA& + 1
                      RW0$ = String$(96, 0)
                      Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                      CIXI% = CII%
                      Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                      CIXI% = CII%
                      Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                      Call REPLA(RW0$, "$ ", 67, 2)
                      Call REPLA(RW0$, MKS$(COU), 69, 4)
                      Call REPLA(RW0$, MKS$(CANT(J%)), 73, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(J%)), 77, 4)
                      Call REPLA(RW0$, MKS$(0), 81, 4)
                      Call REPLA(RW0$, MKS$(0), 85, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(J%)), 89, 4)
                      Call REPLA(RW0$, MKS$(0), 93, 4)
                      Call GRAREG("ANACOS", RW0$, CANA&)
                   End If
                   '
                   If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                      POIN$ = Mid$(Q0$, 55, 74)
                      For U% = 1 To Len(POIN$) Step 2
                         POINI% = CVI(Mid$(POIN$, U%, 2))
                         If POINI% > 0 Then
                            CIXI% = CI%
                            Get #N6%, POINI%, P0$
                            DE$ = TRIM$(Mid$(P0$, 5, 30))
                            MN$ = Mid$(P0$, 35, 2)
                            COU = CVS(Mid$(P0$, 37, 4))
                            CANA& = CANA& + 1
                            '
                            RW0$ = String$(96, 0)
                            Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                            CIXI% = CII%
                            Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                            CIXI% = CII%
                            Call REPLA(RW0$, DE$, 19, 48)
                            Call REPLA(RW0$, MN$, 67, 2)
                            Call REPLA(RW0$, MKS$(COU), 69, 4)
                            Call REPLA(RW0$, MKS$(CANT(J%)), 73, 4)
                            Call REPLA(RW0$, MKS$(0), 77, 4)
                            Call REPLA(RW0$, MKS$(0), 81, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(J%)), 85, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(J%)), 89, 4)
                            Call REPLA(RW0$, MKS$(0), 93, 4)
                            Call GRAREG("ANACOS", RW0$, CANA&)
                            '
                         End If
                      Next U%
                   End If
                End If
             End If
         Next J%
         '
         DP%(CIX0%) = 0
         RECA% = 1
         '
199  Next I&
     If RECA% > 0 Then GoTo 140

210  JJ& = 0
     For I& = 1 To NUINV%
        I0$ = REGLEIDO$("INVERT", I& + 1)
        COD$ = Left$(I0$, 16): CI% = CVI(Mid$(I0$, 17, 2))
        If Asc(Mid$(COD$, 16, 1)) <> 96 Then
            If CI% > 0 Then
                If CI% <= NUMAS% Then
                    Get #N9%, CI%, Q0$
                    Z$ = Left$(Q0$, 64)
                    JJ& = JJ& + 1
                    Call GRAREG("LICOCAL", Z$, JJ&)
                End If
            End If
        End If
     Next I&
     '
     Call SETULTREG("LICOCAL", JJ&)
     Call SETULTREG("ANACOS", CANA&)
     '
     FEX = HOY(HOS$)
     Call GRACONTROL("COSTOS", "FECALCOS", HOS$)
     Call GRACONTROL("COSTOS", "HOCALCOS", Time$)
     '
     FORCOSTO.Label34.Caption = " " + HOS$
     FORCOSTO.Label34.Refresh
     FORCOSTO.Label35.Caption = " " + Time$
     FORCOSTO.Label35.Refresh
     '
3800 Call CIERRARCH("*.*")
     Close #N9%
     Close #N19%
3900 Screen.MousePointer = 1
     '
End Sub

Sub MUECOSTO()
   '
   RFF$ = RECFILT$("COCALCU", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
      I& = CVS(RFF$)
      Q0$ = REGLEIDO$("COCALCU", I&)
      Q9$ = REGLEIDO$("MODIREC", I&)
      If CVI(Left$(Q0$, 2)) = CI% Then
         Label90.Caption = TRIM$(CSTRING$(Mid$(Q0$, 7, 4), 3, 12, 4, 2))
         Label90.Refresh
         Label82.Caption = TRIM$(CSTRING$(Mid$(Q0$, 11, 4), 3, 12, 4, 2))
         Label82.Refresh
         Label88.Caption = TRIM$(CSTRING$(Mid$(Q0$, 19, 4), 3, 12, 4, 2))
         Label88.Refresh
         Label83.Caption = TRIM$(CSTRING$(Mid$(Q0$, 23, 4), 3, 12, 4, 2))
         Label83.Refresh
         Label86.Caption = TRIM$(CSTRING$(Mid$(Q0$, 31, 4), 3, 12, 4, 2))
         Label86.Refresh
         Label84.Caption = TRIM$(CSTRING$(Mid$(Q0$, 35, 4), 3, 12, 4, 2))
         Label84.Refresh
         Label75.Caption = TRIM$(CSTRING$(Mid$(Q9$, 3, 4), 3, 12, 4, 2))
         Label75.Refresh
         Label72.Caption = TRIM$(CSTRING$(Mid$(Q9$, 7, 4), 3, 12, 4, 2))
         Label72.Refresh
         '
         SUMA1 = 0
         SUMA1 = SUMA1 + Val(Label90.Caption)
         SUMA1 = SUMA1 + Val(Label88.Caption)
         SUMA1 = SUMA1 + Val(Label86.Caption)
         SUMA1 = SUMA1 + Val(Label82.Caption)
         SUMA1 = SUMA1 + Val(Label83.Caption)
         SUMA1 = SUMA1 + Val(Label84.Caption)
         Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, 4, 2))
         Label85.Refresh
         '
         SUMA2 = 0
         SUMA2 = SUMA2 + Val(Label75.Caption)
         SUMA2 = SUMA2 + Val(Label72.Caption)
         Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, 4, 2))
         Label80.Refresh
         '
         SUMA3 = SUMA1
         SUMA3 = SUMA3 + Val(Label80.Caption)
         Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, 4, 2))
         Label70.Refresh
         '
         SUMA4 = SUMA3 * (Val(Text4.TEXT)) / 100
         Label68.Caption = TRIM$(CSTRING$(MKS$(SUMA4), 3, 12, 4, 2))
         Label68.Refresh
         '
         SUMA5 = SUMA3 + SUMA4
         Label65.Caption = TRIM$(CSTRING$(MKS$(SUMA5), 3, 12, 4, 2))
         Label65.Refresh
         '
      End If
   End If
End Sub

Sub CARCOSTO()
   '
   Static DP%(32767), FESTRU#
   '
   Screen.MousePointer = 11
   Call FRESHALL
   Call BLANFORCOS
   GoSub 90
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   CICON% = CI%
   '
   NVI% = NUINV%
   NMM% = NUMAS%
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
     FESTRU# = FESTRUC#
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = REGLEIDO$("ESTRUNUE", I&)
       CI1% = CVI(Left$(R0$, 2))
       If CI1% > 0 Then
         If CI1% <= NMM% Then
           If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
             DP%(CI1%) = 1
           End If
         End If
       End If
     Next I&
   End If
   '
   Call EXPLOMUL(CI%)
   SUMPPROP = 0: SUMPREC = 0
   SUPIPROP = 0: SUPIREC = 0
   SUSEPROP = 0: SUSEREC = 0
   '
   ' costo de materias primas y piezas de terceros
   '
   ANACOSMAT.List1.Clear
   SUPRI = 0: SUPIE = 0: SUSER = 0
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
'     II& = KKI%: Call TRACE(20, II&, FIN&)
     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
        MONEI% = MONECOSTO%(CIK%(KKI%))
        DEMONE$ = DEMONECO$(MONEI%)
        COSUDOL = COSUNI(CIK%(KKI%))
        If MONEI% > 1 Then
            COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
          Else
            FEBAS% = FECOSTO%(CIK%(KKI%))
            FEACT% = HOY(HOS$)
            COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
        End If
        '
        Z$ = Space$(256)
        Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
        DESCRI$ = DESCRIT0$(CIK%(KKI%))
        Call REPLA(Z$, DESCRI$, 17, 23)
        Call REPLA(Z$, DESCRI$, 129, 40)
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
        Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
        Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
        Call REPLA(Z$, DEMONE$, 179, 10)
        Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
        UNIME$ = UCase$(UNIMED$(CIK%(KKI%)))
        If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
             If KKI% = 1 Then
                  SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                Else
                  SUPIREC = SUPIREC + CANT(KKI%) * COSUN
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 65, 10)
             SUPIE = SUPIE + CANT(KKI%) * COSUN
          Else
             If KKI% = 1 Then
                  SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                Else
                  SUMPREC = SUMPREC + COSUN * CANT(KKI%)
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 55, 10)
             SUPRI = SUPRI + CANT(KKI%) * COSUN
        End If
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
        ANACOSMAT.List1.AddItem Z$
        '
      Else
        '
        If COSUNI(CIK%(KKI%)) > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
             If PROVINT%(PHB%) = 0 Then
               MONEI% = MONECOSTO%(CIK%(KKI%))
               DEMONE$ = DEMONECO$(MONEI%)
               COSUDOL = COSUNI(CIK%(KKI%))
               If MONEI% > 1 Then
                   COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                 Else
                   FEBAS% = FECOSTO%(CIK%(KKI%))
                   FEACT% = HOY(HOS$)
                   COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
               End If
               '
               Z$ = Space$(256)
               Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
               DESCRI$ = DESCRIT0$(CIK%(KKI%))
               Call REPLA(Z$, DESCRI$, 17, 23)
               Call REPLA(Z$, DESCRI$, 129, 40)
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
               Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
               Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
               Call REPLA(Z$, DEMONE$, 179, 10)
               Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
               If KKI% = 1 Then
                    SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                  Else
                    SUSEREC = SUSEREC + CANT(KKI%) * COSUN
               End If
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 75, 10)
               SUSER = SUSER + CANT(KKI%) * COSUN
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
               ANACOSMAT.List1.AddItem Z$
             End If
           End If
        End If
        '
     End If
   Next KKI%
   '
   If ANACOSMAT.List1.ListCount > 0 Then
      ANACOSMAT.List1.AddItem String$(128, "=")
      ANACOSMAT.Label1.Caption = TRIM$(CSTRING$(MKS$(SUPRI), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label4.Caption = TRIM$(CSTRING$(MKS$(SUPIE), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label3.Caption = TRIM$(CSTRING$(MKS$(SUSER), 3, 10, DECICALCOS%, 2))
      ANACOSMAT.Label5.Caption = TRIM$(CSTRING$(MKS$(SUPRI + SUPIE + SUSER), 3, 10, DECICALCOS%, 2))
   End If
   '
   GoSub 90
   '
   ' costo de mano de obra directa
   '
   FIN& = CONTAMUL%
   URE& = ULTREG&("SECOPER")
   UCT& = ULTREG&("COSOPTER")
   SUMODPRO = 0: SUMODREC = 0
   SUMAQPRO = 0: SUMAQREC = 0
   '
   ANACOSMOD.List1.Clear
   SUTIE = 0: SUMOD = 0: SUMAQ = 0
   ICX% = 0
   '
   For UI& = 1 To FIN&
      '
      CIXI% = CIK%(UI&)
      CODEX$ = CODEXT$(CIXI%)
      CACOM = CANT(UI&)
      '
      RFF$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
      RCT$ = RECFILT$("COSOPTER", 1, MKI$(CIXI%))
      IOP% = 0
      For UJ& = 1 To Len(RFF$) Step 4
        I& = CVS(Mid$(RFF$, UJ&, 4))
        If I& > 0 Then
          If I& <= URE& Then
            SCP$ = REGLEIDO$("SECOPER", I&)
            NOPY% = CVI(Mid$(SCP$, 3, 2))
            DOPE$ = Mid$(SCP$, 5, 48)
            LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
            CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
            MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
            MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
            MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
            MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
            MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
            MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
            HORAPROD = CACOM * MINLOTE / 60 / LOTESTA
            '
            CAOPS = CVS(Mid$(SCP$, 83, 4))
            CATOP% = Asc(Mid$(SCP$, 87, 1))
            EQUIPO$ = TRIM$(Mid$(SCP$, 53, 6))
            COSMOD = CAOPS * HORAPROD * VALHORA(CATOP%)
            COSMAQ = HORAPROD * CUHORMAQ(EQUIPO$)
            '
            TERCOP% = 0: COSTER = 0
            For UT& = 1 To Len(RCT$) Step 4
              IT& = CVS(Mid$(RCT$, UT&, 4))
              If IT& > 0 Then
                If IT& <= UCT& Then
                  SCT$ = REGLEIDO$("COSOPTER", IT&)
                  If CVI(Mid$(SCT$, 3, 2)) = NOPY% Then
                    MONEI% = CVI(Mid$(SCT$, 15, 2))
                    COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * TICAMONE(MONEI%)
                    If MONEI% <= 1 Then
                       FEBAS% = CVI(Mid$(SCT$, 17, 2))
                       FEACT% = HOY(HOS$)
                       COSTER = CACOM * CVD(Mid$(SCT$, 7, 8)) * IAJUINF#(FEBAS%, FEACT%)
                    End If
                    If COSTER >= 0.00005 Then
                      If (COSTER < COSMOD + COSMAQ) Or ((COSMOD + COSMAQ) < 0.00005) Then
                        COSMOD = COSTER
                        COSMAQ = 0
                        HORAPROD = 0
                        TERCOP% = 1
                      End If
                    End If
                  End If
                End If
              End If
            Next UT&
            '
            SUTIE = SUTIE + HORAPROD
            SUMOD = SUMOD + COSMOD
            SUMAQ = SUMAQ + COSMAQ
            '
            If UI& = 1 Then
                SUMODPRO = SUMODPRO + COSMOD
                SUMAQPRO = SUMAQPRO + COSMAQ
              Else
                SUMODREC = SUMODREC + COSMOD
                SUMAQREC = SUMAQREC + COSMAQ
            End If
            '
            IOP% = IOP% + 1: ICX% = ICX% + 1
            Z$ = Space$(256)
            If IOP% = 1 Then
              If ICX% > 1 Then
                ANACOSMOD.List1.AddItem String$(256, "-")
              End If
              Call REPLA(Z$, CODEXT$(CIXI%), 1, 15)
              DESCRI$ = DESCRIT0$(CIXI%)
              Call REPLA(Z$, DESCRI$, 17, 21)
              Call REPLA(Z$, DESCRI$, 129, 40)
            End If
            Call REPLA(Z$, DOPE$, 39, 15)
            Call REPLA(Z$, DOPE$, 169, 30)
            Call REPLA(Z$, CSTRING$(MKS$(CAOPS), 3, 5, 2, 2), 199, 5)
            Call REPLA(Z$, HORMINSE$(3600 * HORAPROD), 55, 10)
            If TERCOP% = 1 Then
              Call REPLA(Z$, "  (3ros)", 55, 10)
            End If
            Call REPLA(Z$, CSTRING$(MKS$(COSMOD), 3, 10, DECICALCOS%, 2), 65, 10)
            Call REPLA(Z$, CSTRING$(MKS$(COSMAQ), 3, 10, DECICALCOS%, 2), 75, 10)
            Call REPLA(Z$, CSTRING$(MKS$(COSMOD + COSMAQ), 3, 10, DECICALCOS%, 2), 85, 10)
            If COSTER * TERCOP% > 0 Then
              Call REPLA(Z$, CSTRING$(MKS$(COSTER / CACOM), 3, 10, DECICALCOS%, 2), 204, 10)
            End If
            ANACOSMOD.List1.AddItem Z$
          End If
        End If
      Next UJ&
   Next UI&
   '
   If ANACOSMOD.List1.ListCount > 0 Then
      ANACOSMOD.List1.AddItem String$(256, "=")
      ANACOSMOD.Label3.Caption = HORMINSE$(3600 * SUTIE)
      ANACOSMOD.Label1.Caption = TRIM$(CSTRING$(MKS$(SUMOD), 3, 10, DECICALCOS%, 2))
      ANACOSMOD.Label4.Caption = TRIM$(CSTRING$(MKS$(SUMAQ), 3, 10, DECICALCOS%, 2))
      ANACOSMOD.Label5.Caption = TRIM$(CSTRING$(MKS$(SUMOD + SUMAQ), 3, 10, DECICALCOS%, 2))
   End If
   '
   GoSub 90
   '
   PRELISX# = PRELISTA#("LIPRE001", CICON%) * TICAMONE(MONELISTA%(1))
   Label114.Caption = TRIM$(CSTRING$(MKD$(PRELISX#), 3, 12, DECICALCOS%, 2))
   DoEvents
   '
   YACAR%(4) = 1
   MODIFIC%(4) = 0
   Screen.MousePointer = 1
   '
   If ORILLAMA$ = "MAT" Then
     If REGILLAMA% >= 0 Then
       If REGILLAMA% < ANACOSMAT.List1.ListCount Then
         ANACOSMAT.List1.ListIndex = REGILLAMA%
         On Error Resume Next
           ANACOSMAT.List1.TopIndex = TOPELLAMA%
         On Error GoTo 0
       End If
     End If
     ORILLAMA$ = ""
     REGILLAMA% = -1
     TOPELLAMA% = 0
     LLAMACOSTO$ = ""
     ANACOSMAT.Show 1
   End If
   '
   If ORILLAMA$ = "MOD" Then
     If REGILLAMA% >= 0 Then
       If REGILLAMA% < ANACOSMOD.List1.ListCount Then
         ANACOSMOD.List1.ListIndex = REGILLAMA%
         On Error Resume Next
           ANACOSMOD.List1.TopIndex = TOPELLAMA%
         On Error GoTo 0
       End If
     End If
     ORILLAMA$ = ""
     REGILLAMA% = -1
     TOPELLAMA% = 0
     LLAMACOSTO$ = ""
     ANACOSMOD.Show 1
   End If
   '
Exit Sub
   '
   '
   '
90 Label90.Caption = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   Label90.Refresh
   Label82.Caption = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   Label82.Refresh
   Label88.Caption = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   Label88.Refresh
   Label83.Caption = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   Label83.Refresh
   Label86.Caption = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   Label86.Refresh
   Label84.Caption = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   Label84.Refresh
   '
   Label75.Caption = TRIM$(CSTRING$(MKS$(SUMODPRO), 3, 12, DECICALCOS%, 2))
   Label75.Refresh
   Label72.Caption = TRIM$(CSTRING$(MKS$(SUMODREC), 3, 12, DECICALCOS%, 2))
   Label72.Refresh
   Label79.Caption = TRIM$(CSTRING$(MKS$(SUMAQPRO), 3, 12, DECICALCOS%, 2))
   Label79.Refresh
   Label74.Caption = TRIM$(CSTRING$(MKS$(SUMAQREC), 3, 12, DECICALCOS%, 2))
   Label74.Refresh
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(Label90.Caption)
   SUMA1 = SUMA1 + Val(Label88.Caption)
   SUMA1 = SUMA1 + Val(Label86.Caption)
   SUMA1 = SUMA1 + Val(Label82.Caption)
   SUMA1 = SUMA1 + Val(Label83.Caption)
   SUMA1 = SUMA1 + Val(Label84.Caption)
   Label85.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, DECICALCOS%, 2))
   Label85.Refresh
   '
   SUMA2 = 0
   SUMA2 = SUMA2 + Val(Label75.Caption)
   SUMA2 = SUMA2 + Val(Label72.Caption)
   SUMA2 = SUMA2 + Val(Label79.Caption)
   SUMA2 = SUMA2 + Val(Label74.Caption)
   Label80.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, DECICALCOS%, 2))
   Label80.Refresh
   '
   SUMA3 = SUMA1 + SUMA2
   Label70.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, DECICALCOS%, 2))
   Label70.Refresh
   '
Return
'
End Sub

Sub CANACOSMAT(CIXX%)
   '
'Exit Sub
   Static ENCALI$, LINSEP$
   Static DP%(16384), FESTRU#
   '
   Screen.MousePointer = 11
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
     NVI% = NUINV%
     NMM% = NUMAS%
     FESTRU# = FESTRUC#
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = REGLEIDO$("ESTRUNUE", I&)
       CI1% = CVI(Left$(R0$, 2))
       If CI1% > 0 Then
         If CI1% <= NMM% Then
           DP%(CI1%) = 1
         End If
       End If
     Next I&
   End If
   '
   If ENCALI$ = "" Then
      ENCALI$ = "Codigo          Descripcion                                    Operacion                             $/Unid      Cant     Tot.Mat    Tot.Serv    Tot.Comp"
      LINSEP$ = String$(154, "-")
   End If
   '
   UREANA& = ULTREG&("ANACOS")
   Dim RECOSMAT$(4096)
   CACOM% = 0
   JJ& = 0
   '
   FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMAT.TXT", 2, 0)
   Call EXPLOMUL(CIXX%)
   For KKI% = 2 To CONTAMUL
     If DP%(CIK%(KKI%)) < 1 Then
       CODD$ = CODEXT$(CIK%(KKI%))
       DECO$ = DESCRIT$(CIK%(KKI%))
       DEOP$ = ""
       VUNI = COSUNI(CIK%(KKI%))
       MONEI% = MONECOSTO%(CIK%(KKI%))
       If MONEI% = 2 Then VUNI = VUNI * TICAMBIO
       CANX = CANT(KKI%)
       CMAT = CANX * VUNI
       '
       TEBUSCA$ = Space$(75)
       Call REPLA(TEBUSCA$, CODD$, 1, 16)
       Call REPLA(TEBUSCA$, DECO$, 17, 47)
       'Call REPLA(TEBUSCA$, DEOP$, 64, 32)
       Call REPLA(TEBUSCA$, CSTRING$(MKS$(VUNI), 3, 12, 4, 2), 64, 12)
       '
       For UJ% = 1 To CACOM%
         If Left$(RECOSMAT$(UJ%), 75) = TEBUSCA$ Then
           CANX = CANX + Val(Mid$(RECOSMAT$(UJ%), 76, 10))
           CMAT = CMAT + Val(Mid$(RECOSMAT$(UJ%), 86, 12))
           CSER = CSER + Val(Mid$(RECOSMAT$(UJ%), 98, 12))
           TENCU$ = RECOSMAT$(UJ%)
           Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 76, 10)
           Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 86, 12)
           Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 98, 12)
           RECOSMAT(UJ%) = TENCU$
           GoTo 9
         End If
       Next UJ%
       CACOM% = CACOM% + 1
       TENCU$ = Space$(154)
       Call REPLA(TENCU$, TEBUSCA$, 1, 75)
       Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 76, 10)
       Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 86, 12)
       Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 98, 12)
       RECOSMAT(CACOM%) = TENCU$
     End If
9  Next KKI%
   '
   For UI% = 1 To CACOM%
     For UJ% = UI% + 1 To CACOM%
       If Left$(RECOSMAT$(UI%), 16) > Left$(RECOSMAT$(UJ%), 16) Then
         TENCU$ = RECOSMAT$(UJ%)
         RECOSMAT$(UJ%) = RECOSMAT$(UI%)
         RECOSMAT$(UI%) = TENCU$
       End If
     Next UJ%
   Next UI%
   '
   SUMAPAR = 0: SUMATO = 0: SUMATER = 0: SUSERVI = 0
   Print #FISAL%, ENCALI$
   Print #FISAL%, LINSEP$
   For UI% = 1 To CACOM%
     Z$ = RECOSMAT$(UI%)
     SUMAPAR = SUMAPAR + Val(Mid$(Z$, 86, 12)) + Val(Mid$(Z$, 98, 12))
     SUMATO = SUMATO + Val(Mid$(Z$, 86, 12)) + Val(Mid$(Z$, 98, 12))
     SUMATER = SUMATER + Val(Mid$(Z$, 86, 12))
     SUSERVI = SUSERVI + Val(Mid$(Z$, 98, 12))
     If UI% = CACOM% Or Left$(RECOSMAT$(UI%), 16) <> Left$(RECOSMAT$(UI% + 1), 16) Then
       Call REPLA(Z$, CSTRING$(MKS$(SUMAPAR), 3, 12, 4, 2), 110, 12)
       SUMAPAR = 0
       ATXX% = 1
     End If
     Print #FISAL%, Z$
     If ATXX% = 1 Then Print #FISAL%, LINSEP$
   Next UI%
   '
   Z$ = Space$(154)
   Call REPLA(Z$, "S U B T O T A L   Materiales y Servicios", 1, 40)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATER), 3, 12, 4, 2), 86, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUSERVI), 3, 12, 4, 2), 98, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATO), 3, 12, 4, 2), 110, 12)
   Print #FISAL%, Z$
   Print #FISAL%, LINSEP$
   Close #FISAL%
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CALCOSOPE()
   '
   LOTESTA = Val(Text7.TEXT)
   If LOTESTA < 1 Then LOTESTA = 1
   CICLOHE = Val(Text34.TEXT)
   If CICLOHE < 1 Then CICLOHE = 1
   '
   CALOTES = 1
   MINPREP = 60 * Val(Text15(0)) + Val(Text17(0)) + Val(Text26(0)) / 60
   MINFIJO = 60 * Val(Text15(1)) + Val(Text17(1)) + Val(Text26(1)) / 60
   MINVARI = 60 * Val(Text15(2)) + Val(Text17(2)) + Val(Text26(2)) / 60
   MINHERR = LOTESTA * (60 * Val(Text15(4)) + Val(Text17(4)) + Val(Text26(4)) / 60) / CICLOHE
   MINDEMO = (MINFIJO + MINVARI + MINHERR) * Val(Text29.TEXT) / 100
   MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
   HORAPROD = MINLOTE / 60 / LOTESTA
   '
   CAOPS = Val(Text30.TEXT)
   CATOP% = Val(Text33.TEXT)
   EQUIPO$ = Text9(0).TEXT
   CUOTOPE = VALHORA(CATOP%)
   CUOTMAQ = CUHORMAQ(EQUIPO$)
   '
   COSMOD = CUOTOPE * CAOPS * HORAPROD
   COSMAQ = CUOTMAQ * HORAPROD
   '
   Label73.Caption = TRIM$(CSTRING$(MKS$(COSMOD + COSMAQ), 3, 12, 4, 2))
   '
End Sub

Sub BLANFORCOS()
   '
   Label80 = ""
   Label85 = ""
   Label70 = ""
   Label90 = ""
   Label82 = ""
   Label88 = ""
   Label83 = ""
   Label86 = ""
   Label84 = ""
   Label75 = ""
   Label72 = ""
   Label79 = ""
   Label74 = ""
   Label114 = ""
   Label116 = ""
   ANACOSMAT.List1.Clear
   ANACOSMOD.List1.Clear
   DoEvents
   '
End Sub

Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   Command26.SetFocus
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
   IMPLOSI.Top = Top + (Height - ScaleHeight - 50) + SSTab1.Top + MARCOFICHA(2).Top + Picture5.Top
   IMPLOSI.Left = Left + (Width - ScaleWidth) / 2 + SSTab1.Left + MARCOFICHA(2).Left + Picture5.Left
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   IMPLOSI.Show
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
End Sub

Sub ACVEPROINT()
   For UI& = 0 To 32767
     PROVINT%(UI&) = 0
   Next UI&
   '
   For U& = 1 To ULTREG&("TAPROINT")
      X$ = REGLEIDO$("TAPROINT", U&)
      NP% = CVI(Left$(X$, 2))
      If NP% > 0 Then
        PROVINT%(NP%) = 1
      End If
   Next U&
   Call CIERRARCH("TAPROINT")
   '
End Sub

Sub PRIFICHCOS()
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      Exit Sub
   End If
   '
   FORIPRE$ = "FANACOS"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        PRINTFORM
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Analisis de Costo " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        '
        CODPARAM$(5) = "USO-MPRI": DOCPARAM$(5) = Label90
        CODPARAM$(6) = "PIE-COMP": DOCPARAM$(6) = Label88
        CODPARAM$(7) = "DOMI-TRA": DOCPARAM$(7) = Label86
        CODPARAM$(8) = "CAN-MPRI": DOCPARAM$(8) = Label82
        CODPARAM$(9) = "CAN-COMP": DOCPARAM$(9) = Label83
        CODPARAM$(10) = "LOCA-TRA": DOCPARAM$(10) = Label84
        CODPARAM$(11) = "CANT-RES": DOCPARAM$(11) = Label85
        '
        CODPARAM$(12) = "TOT-PREP": DOCPARAM$(12) = Label75
        CODPARAM$(13) = "TOT-PROD": DOCPARAM$(13) = Label79
        CODPARAM$(14) = "TOT-DEBE": DOCPARAM$(14) = Label72
        CODPARAM$(15) = "TOT-HABE": DOCPARAM$(15) = Label74
        CODPARAM$(16) = "TOT-TIEM": DOCPARAM$(16) = Label80
        '
        CODPARAM$(17) = "PRE-LIST": DOCPARAM$(17) = Label114
        CODPARAM$(18) = "POR-DESC": DOCPARAM$(18) = Text44
        CODPARAM$(19) = "IMP-DESC": DOCPARAM$(19) = Label115
        CODPARAM$(20) = "PRE-NETO": DOCPARAM$(20) = Label112
        CODPARAM$(21) = "IMPNEDES": DOCPARAM$(21) = Label116
        CODPARAM$(22) = "NOP-OPER": DOCPARAM$(22) = Label117
        CODPARAM$(23) = "VALO-DEC": DOCPARAM$(23) = Label118
        CODPARAM$(24) = "DES-OPER": DOCPARAM$(24) = Label119
        CAPARDOC% = 24
        '
        CODTABU1$(1) = "POR-IVA"
        CODTABU1$(2) = "CANTIDAD"
        CODTABU1$(3) = "NET-ITEM"
        CACOLTAB1% = 3
        '
        For KK1% = 1 To 3
          For KK2% = 1 To 11
            TETABU1$(KK1%, KK2%) = ""
          Next KK2%
        Next KK1%
        '
        TETABU1$(3, 1) = Label70
        '
        TETABU1$(1, 2) = Text41
        TETABU1$(2, 2) = Label93
        '
        TETABU1$(1, 3) = Text36
        TETABU1$(2, 3) = Label68
        '
        TETABU1$(3, 4) = Label65
        '
        TETABU1$(1, 5) = Text38
        TETABU1$(2, 5) = Label77
        '
        TETABU1$(3, 6) = Label96
        '
        TETABU1$(1, 7) = Text42
        TETABU1$(2, 7) = Label110
        '
        TETABU1$(1, 8) = Text43
        TETABU1$(2, 8) = Label104
        '
        TETABU1$(3, 9) = Label111
        '
        TETABU1$(1, 10) = Text39
        TETABU1$(2, 10) = Label107
        '
        TETABU1$(3, 11) = Label103
        '
        CAITAB1% = 11
        '
        Call PRINTDOC("@FANACOS")
        Screen.MousePointer = 1
        '
      End If
      '
End Sub

Sub PRIANASUMI()
   '
   FORIPRE$ = "FANACOP"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        Exit Sub
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Costo Suministros " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = ANACOSMAT.Label3.Caption
        CODPARAM$(6) = "TOT-DEBE": DOCPARAM$(6) = ANACOSMAT.Label1.Caption
        CODPARAM$(7) = "TOT-HABE": DOCPARAM$(7) = ANACOSMAT.Label4.Caption
        CODPARAM$(8) = "TOT-PROD": DOCPARAM$(8) = ANACOSMAT.Label5.Caption
        CAPARDOC% = 8
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "DES-OPER"
        CODTABU1$(4) = "NUOPERS"
        CODTABU1$(5) = "TIE-PREP"
        CODTABU1$(6) = "PRE-LIST"
        CODTABU1$(7) = "IMP-DEBE"
        CODTABU1$(8) = "IMP-HABE"
        CODTABU1$(9) = "TIE-PROD"
        CODTABU1$(10) = "TIE-TOTA"
        CODTABU1$(11) = "IMP-TOTA"
        CACOLTAB1% = 11
        '
        CAITAB1% = 0: SUMAPAR# = 0
        For UI& = 1 To ANACOSMAT.List1.ListCount
          '
          TTXX$ = ANACOSMAT.List1.List(UI& - 1)
          If Left$(TTXX$, 64) = String$(64, "-") Then GoTo 19
          If Left$(TTXX$, 64) = String$(64, "=") Then GoTo 19
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Left$(TTXX$, 15)
          TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 129, 40)
          TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 40, 10)
          TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 52, 4)
          TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 169, 10)
          TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 179, 10)
          TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 189, 10)
          TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 55, 10)
          TETABU1$(9, CAITAB1%) = Mid$(TTXX$, 65, 10)
          TETABU1$(10, CAITAB1%) = Mid$(TTXX$, 75, 10)
          TETABU1$(11, CAITAB1%) = Mid$(TTXX$, 85, 10)
19      Next UI&
        '
        Call PRINTDOC("@FANACOP")
        Screen.MousePointer = 1
        '
      End If
      '
End Sub
   
Sub PRIANAMODI()
   '
   FORIPRE$ = "FANACOM"
   If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
        Exit Sub
      Else
        Screen.MousePointer = 11
        FEX = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = ""
        TIPODOC$ = "Costo M.O.D." + TRIM$(Text1.TEXT) + " - " + TRIM$(Text3.TEXT)
        '
        CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = TRIM$(Text1.TEXT)
        CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Text3.TEXT)
        CODPARAM$(3) = "UNI-CONJ": DOCPARAM$(3) = Label37.Caption
        CODPARAM$(4) = "NRO-LOTE": DOCPARAM$(4) = Str$(LOTESTAN(CODINT%(Text1.TEXT)))
        CODPARAM$(5) = "TOT-TIEM": DOCPARAM$(5) = ANACOSMOD.Label3.Caption
        CODPARAM$(6) = "TOT-DEBE": DOCPARAM$(6) = ANACOSMOD.Label1.Caption
        CODPARAM$(7) = "TOT-HABE": DOCPARAM$(7) = ANACOSMOD.Label4.Caption
        CODPARAM$(8) = "TOT-PROD": DOCPARAM$(8) = ANACOSMOD.Label5.Caption
        CAPARDOC% = 8
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "DES-OPER"
        CODTABU1$(4) = "NUOPERS"
        CODTABU1$(5) = "TIE-PREP"
        CODTABU1$(6) = "PRE-LIST"
        CODTABU1$(7) = "IMP-DEBE"
        CODTABU1$(8) = "IMP-HABE"
        CODTABU1$(9) = "TIE-PROD"
        CODTABU1$(10) = "TIE-TOTA"
        CACOLTAB1% = 10
        '
        CAITAB1% = 0: SUMAPAR# = 0
        For UI& = 1 To ANACOSMOD.List1.ListCount
          '
          TTXX$ = ANACOSMOD.List1.List(UI& - 1)
          If Left$(TTXX$, 64) = String$(64, "-") Then GoTo 19
          If Left$(TTXX$, 64) = String$(64, "=") Then GoTo 19
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = Left$(TTXX$, 15)
          TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 129, 40)
          TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 169, 30)
          TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 199, 5)
          TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 55, 10)
          TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 204, 10)
          TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 65, 10)
          TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 75, 10)
          TETABU1$(9, CAITAB1%) = Mid$(TTXX$, 85, 10)
          SUMAPAR# = SUMAPAR# + Val(Mid$(TTXX$, 85, 10))
          '
          TETABU1$(10, CAITAB1%) = ""
          TTYY$ = String$(64, "-")
          If UI& < ANACOSMOD.List1.ListCount Then
            TTYY$ = Left$(ANACOSMOD.List1.List(UI&), 64)
          End If
          If TTYY$ = String$(64, "-") Or TTYY$ = String$(64, "=") Then
            TTZZ$ = CSTRING$(MKD$(SUMAPAR#), 4, 10, DECICALCOS%, 2)
            TETABU1$(10, CAITAB1%) = TTZZ$
            SUMAPAR# = 0
            '
            CAITAB1% = CAITAB1% + 1
            For KKI% = 1 To 10
              TETABU1$(KKI%, CAITAB1%) = ""
            Next KKI%
          End If
          '
19      Next UI&
        '
        Call PRINTDOC("@FANACOM")
        Screen.MousePointer = 1
        '
      End If
      '

End Sub

