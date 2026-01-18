VERSION 5.00
Begin VB.Form AMAS_MAQ04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Repuestos e Insumos - A-B-M-L Maestro de Items de Stock"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "AMAS_MAQ04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
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
      Height          =   1030
      Index           =   2
      Left            =   8660
      TabIndex        =   47
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton Command20 
         Height          =   600
         Left            =   1390
         Picture         =   "AMAS_MAQ04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Búsqueda por Descripción Corta,Descripción General - Especificación, etc"
         Top             =   280
         Width           =   630
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
         Left            =   140
         Picture         =   "AMAS_MAQ04.frx":1670
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   280
         Width           =   630
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
         Height          =   600
         Left            =   770
         Picture         =   "AMAS_MAQ04.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   280
         Width           =   630
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
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
      Height          =   1030
      Index           =   1
      Left            =   7140
      TabIndex        =   46
      Top             =   240
      Width           =   1110
      Begin VB.CommandButton Cmd100 
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
         Left            =   240
         Picture         =   "AMAS_MAQ04.frx":1C84
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Listado de Implosión"
         Top             =   280
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1005
      TabIndex        =   45
      Top             =   0
      Width           =   1065
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":1F8E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Alta de Nuevo Repuesto"
         Top             =   2895
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":20D8
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   2265
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":23E2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Otros Accesos Directos"
         Top             =   3540
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":26EC
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8080
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command37 
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
         Picture         =   "AMAS_MAQ04.frx":29F6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5140
         Width           =   650
      End
      Begin VB.CommandButton Command38 
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
         Picture         =   "AMAS_MAQ04.frx":2E38
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Listado de Atributos Generales"
         Top             =   5780
         Width           =   650
      End
      Begin VB.CommandButton Command39 
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
         Picture         =   "AMAS_MAQ04.frx":3142
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   6420
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":37AC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   9345
         Visible         =   0   'False
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":3AB6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sustitución de Código de Repuesto"
         Top             =   8700
         Visible         =   0   'False
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":3C00
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Alta de Nuevo Repuesto  por Copia Completa de Datos de Uno Existente"
         Top             =   8055
         Visible         =   0   'False
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
         Height          =   630
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":3D4A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ayuda en Línea"
         Top             =   850
         Width           =   650
      End
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
         Height          =   640
         Index           =   4
         Left            =   105
         Picture         =   "AMAS_MAQ04.frx":4054
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit"
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "AMAS_MAQ04.frx":419E
         Stretch         =   -1  'True
         Top             =   7350
         Width           =   1410
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6420
      Index           =   0
      Left            =   105
      TabIndex        =   28
      Top             =   1470
      Width           =   10740
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Características"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   210
         TabIndex        =   37
         Top             =   315
         Width           =   4740
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
            Left            =   945
            MaxLength       =   20
            TabIndex        =   4
            Top             =   2310
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
            Height          =   300
            Left            =   4335
            TabIndex        =   38
            Top             =   1320
            Width           =   175
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
            Height          =   300
            Left            =   4335
            TabIndex        =   39
            Top             =   480
            Width           =   175
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
            Height          =   300
            Left            =   3795
            MaxLength       =   3
            TabIndex        =   2
            Text            =   " "
            Top             =   480
            Width           =   570
         End
         Begin VB.TextBox Text4 
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
            Left            =   3795
            MaxLength       =   2
            TabIndex        =   3
            Text            =   " "
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   3255
            TabIndex        =   55
            Top             =   2310
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "$ Ultima Compra:"
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
            TabIndex        =   54
            Top             =   2205
            Width           =   975
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
            Height          =   540
            Left            =   -105
            TabIndex        =   49
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label Label29 
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
            Height          =   300
            Left            =   315
            TabIndex        =   43
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label30 
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
            Height          =   300
            Left            =   315
            TabIndex        =   42
            Top             =   1680
            Width           =   4215
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
            Height          =   300
            Left            =   525
            TabIndex        =   41
            Top             =   1440
            Width           =   3180
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
            Height          =   300
            Left            =   525
            TabIndex        =   40
            Top             =   580
            Width           =   3180
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Reposición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   5145
         TabIndex        =   36
         Top             =   315
         Width           =   5400
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
            Height          =   1860
            Left            =   225
            TabIndex        =   17
            ToolTipText     =   "Doble-Click para Cargar Tabla  de Códigos Equivalentes"
            Top             =   735
            Width           =   4950
         End
         Begin VB.Label Label92 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proveedor                Cod. Equiv."
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
            Left            =   225
            TabIndex        =   50
            Top             =   465
            Width           =   4950
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2310
         ScaleHeight     =   375
         ScaleWidth      =   8145
         TabIndex        =   34
         Top             =   3345
         Width           =   8205
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
            TabIndex        =   35
            Top             =   105
            Width           =   7680
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   2310
         TabIndex        =   33
         Top             =   5880
         Width           =   3240
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   9180
         TabIndex        =   32
         Top             =   5880
         Width           =   1320
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7905
         TabIndex        =   31
         Top             =   5880
         Width           =   1275
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6735
         TabIndex        =   30
         Top             =   5880
         Width           =   1170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5550
         TabIndex        =   29
         Top             =   5880
         Width           =   1170
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
         Height          =   2130
         Left            =   2310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3765
         Width           =   8205
      End
      Begin VB.Line Line1 
         X1              =   10510
         X2              =   10510
         Y1              =   3360
         Y2              =   6135
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
         TabIndex        =   44
         Top             =   3765
         Width           =   1905
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   900
      Left            =   8400
      ScaleHeight     =   840
      ScaleWidth      =   60
      TabIndex        =   25
      Top             =   370
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   12000
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
      TabIndex        =   1
      Top             =   735
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
      Height          =   285
      Left            =   1575
      MaxLength       =   15
      TabIndex        =   18
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
      Height          =   285
      Left            =   3600
      Picture         =   "AMAS_MAQ04.frx":60C0
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label58 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   420
      TabIndex        =   27
      Top             =   1050
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3465
      TabIndex        =   24
      Top             =   360
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
      Height          =   285
      Left            =   4560
      TabIndex        =   23
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   22
      Top             =   360
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
      Height          =   285
      Left            =   6195
      TabIndex        =   21
      Top             =   315
      Width           =   750
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   20
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   420
      TabIndex        =   19
      Top             =   360
      Width           =   1005
   End
   Begin VB.Menu archiv 
      Caption         =   "Archivos"
      Begin VB.Menu Abri 
         Caption         =   "Abrir"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu nueite 
         Caption         =   "Nuevo Repuesto"
      End
      Begin VB.Menu altcop 
         Caption         =   "Alta por Copia"
         Visible         =   0   'False
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu guar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu impri 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu Sali 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "Edición"
      Begin VB.Menu deshcam 
         Caption         =   "Deshacer Cambios"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu SuCoAr 
         Caption         =   "Sustitución de Código de Repuesto"
         Visible         =   0   'False
      End
      Begin VB.Menu BaLoIn 
         Caption         =   "Baja Lógica Individual"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu CoyLi 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ListAtGe 
         Caption         =   "Listado de Atributos Generales"
      End
   End
   Begin VB.Menu Config 
      Caption         =   "Configuración"
      Begin VB.Menu Configral 
         Caption         =   "Configuración General"
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu TaTiMa 
         Caption         =   "Tabla de Tipos de Material"
      End
      Begin VB.Menu TaUnMe 
         Caption         =   "Tabla de Unidades de Medida"
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu ConfIndu 
         Caption         =   "Configuración Industrial"
         Begin VB.Menu MnuPlant 
            Caption         =   "Plantas Industriales"
         End
         Begin VB.Menu DepAlm 
            Caption         =   "Depositos y Almacenes"
         End
         Begin VB.Menu SeccyCen 
            Caption         =   "Secciones y Centros de Costos"
         End
         Begin VB.Menu MayEq 
            Caption         =   "Máquinas y Equipos de Producción"
         End
         Begin VB.Menu CaOpPro 
            Caption         =   "Categorias de Operarios "
         End
      End
   End
   Begin VB.Menu Ayu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "AMAS_MAQ04"
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



Private Sub Abri_Click()
Call Command9_Click
End Sub

Private Sub altcop_Click()
Call Command22_Click
End Sub

Private Sub Ayu_Click()
  '
  Call Command4_Click
  '
End Sub

Private Sub BaLoIn_Click()
  '
  Call Command24_Click
  '
End Sub

Private Sub BOTOPENFORM_Click(Index As Integer)
  '
  '
End Sub

Private Sub CaOpPro_Click()
  SETGEMA04.Option1(4).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub CoeInBdD_Click()
Call Command25_Click
End Sub

Private Sub Cmd100_Click()
  '
  Cmd100.Enabled = False
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call ABREMAQUI
    SQLX$ = "SELECT * FROM REPUMAQ"
    SQLX$ = SQLX$ + " WHERE CodInte_Rep = " & CI1% & " "
    Set RSTREP = Maquinarias.OpenRecordset(SQLX$, 4)
    With RSTREP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
     End If
     On Error GoTo 0
     '
     L1$ = Space$(256)
     T1$ = Space$(256)
     '
     IMPLOSI.Label27 = ""
     IMPLOSI.List1.Clear
     T1$ = "Codigo  Descripcion                                           Marca          Modelo"
     IMPLOSI.Label27 = T1$
     IMPLOSI.Caption = "Implosión del Repuesto: " + Text1.TEXT + " - " + Text3.TEXT
     JJ& = 0
     Do While .EOF = False
      MAQUI1$ = !Codigo_Equ
      With Equipos
        .FindFirst " Codigo_Equ = '" & MAQUI1$ & "' "
        If .NoMatch = False Then
          '
          EQ1$ = ECOARCH$("INDIEQUI", !Codigo_Equ)
          Call REPLA(L1$, !Codigo_Equ, 1, 6)
          Call REPLA(L1$, EQ1$, 8, 30)
          Call REPLA(L1$, !Marca_Equ, 40, 10)
          Call REPLA(L1$, !Mode_Equ, 51, 10)
          '
          JJ& = JJ& + 1
          IMPLOSI.List1.AddItem L1$
        End If
      End With
      .MoveNext
     Loop
     If JJ& = 0 Then
      Call MENSERR(24, "El Repuesto: '" & Text1.TEXT & "' \No se Encuentra Asociado a Ninguna Máquina")
      GoTo 99
     End If
     IMPLOSI.Show 1
    End With
  End If
99 Cmd100.Enabled = True
  '
End Sub

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
  Call Command1_Click
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call CONECRUN("PLANEL04/" + TRIM$(Text1), "Planera Electrónica")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
 Command6.Enabled = False
   Call ACCDIR("AMA_GES")
   Command6.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
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
   MODOEDIT% = 0
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Command20_Click()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Repuesto por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("ECOMAST", "SELMAS")
     '
     Call ABREMASTER
     On Error Resume Next
     With Master
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           If !STATUS <= 0 Then GoTo 15 'no presenta repuestos de baja
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               CI1% = !CODINT
               Call REPLA(XX$, MKI$(CI1%), 63, 2)
               Call REPLA(XX$, CODEXT$(CI1%), 1, 16)
               Call REPLA(XX$, DESCRIT0$(CI1%), 17, 46)
               JJ& = JJ& + 1
               Call GRAREG("!SELMAS", XX$, JJ&)
               Exit For
             End If
           Next KKI%
15       .MoveNext
         Loop
       End If
     End With
     Call SETULTREG("!SELMAS", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         CI2% = TRIM$(Str$(CVI(Mid$(REGLEIDO$("!SELMAS", 1), 63, 2))))
         Text1.TEXT = CODEXT$(CI2%)
       Else
         Call FRESHECHO("!SELMAS")
         Call SELECHO("!SELMAS")
         NCLIX$ = VALACT1$("!SELMAS")
         If NCLIX$ <> "" Then
           Text1 = NCLIX$
         End If
     End If
   End If
   Command20.Enabled = True
End Sub

Private Sub Command21_Click()
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
   End If
   '
   Call GENECOMAS
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
    '
   Dim FIELDVAL(256)
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
   VDEF$ = REGBLAN$("MASTER")
   CINUE% = 1 + NUMAS%
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, DESCRIT0$(CII%), 5, 64)
   Call REPLA(VDEF$, UNIMED$(CII%), 109, 2)
   Call REPLA(VDEF$, Chr$(TIMATE%(CII%)), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
     Call ABREMASTER
     CACAM% = Articulos.TableDefs("Master").Fields.Count
     Master.FindFirst ("CODINT = " & CII%)
     If Master.NoMatch = False Then
       Master.Edit
       For KU% = 1 To CACAM%
         FIELDVAL(KU%) = Master.Fields(KU% - 1).Value
       Next KU%
       '
       Master.FindFirst ("CODINT = " & CINUE%)
       If Master.NoMatch = False Then
         Master.Edit
         For KU% = 1 To CACAM%
           NOCAMPO$ = TRIM$(UCase$(Master.Fields(KU% - 1).Name))
           If NOCAMPO$ <> "CODIGO" Then
             If NOCAMPO$ <> "DESCRIPCION" Then
               If NOCAMPO$ <> "CODINT" Then
                 If NOCAMPO$ <> "ALIAS" Then
                   If NOCAMPO$ <> "UMEDIDA" Then
                     Master.Fields(KU% - 1).Value = FIELDVAL(KU%)
                   End If
                 End If
               End If
             End If
           End If
         Next KU%
         Master.Update
       End If
       '
       Call COPYFORMU(CII%, CINUE%)
       Call COPYRUTA(CII%, CINUE%)
     End If
     'COPIA LA LISTA DE PROVEEDORES CON LOS CODI. EQUIVAL.
     CI0% = CODINT%(CODIVIE$)
     Call FILTERFILE("REPLACOP", "!REPLACOP", 3, MKI$(CI0%))
     For J& = 1 To ULTREG&("!REPLACOP")
       X$ = REGLEIDO$("!REPLACOP", J&)
       Call REPLA(X$, MKI$(CINUE%), 3, 2)
       Call REGAPP("REPLACOP", X$)
     Next J&
     '
     CODINUE$ = Mid$(VDEF$, 77, 16)
     Text1.TEXT = CODINUE$
     '
   End If
   '
99 Call GENECOMAS
   Command22.Enabled = True
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
   CI% = CODINT%(Text1.TEXT)
   If CI% > 0 Then
     Master.Edit
     Master!Codigo = CODINUE$
     Master.Update
     UPDATMASTER% = 1
   End If
   '
   Text1.TEXT = CODINUE$
   '
99 Call GENECOMAS
   Screen.MousePointer = 1
   Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command24.Enabled = False
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
   CII% = CODINT%(Text1.TEXT)
   If CII% > 0 Then
     Master.Edit
     Master!STATUS = (-1)
     Master.Update
     UPDATMASTER% = 1
     Call BLANFORMU
     Text1 = ""
   End If
   '
99 Call GENECOMAS
   Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
99  Command25.Enabled = True
End Sub

Private Sub Command3_Click(Index As Integer)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Sub Command37_Click()
   With SETGEMA04
     .Height = .Frame2.Top + 250
     .Show 1
   End With
End Sub

Private Sub Command38_Click()
   Command38.Enabled = False
   Call LIMASTER(1)
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   PRINTFORM
   Command39.Enabled = True
End Sub


Private Sub Command4_Click()
  '
  Call HELPONLINE("AMASMAQ")
  '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub
Sub LIMASTER(MODO%)
    '
    HOO% = HOY(HO$)
    If MODO% = 0 Then
      MODO% = ALTERNA%("Listado de Códigos y Atributos\Código y Descripción para Inventarios\Stock Mínimo y Lote Reposición")
      If MODO% < 1 Or MODO% > 3 Then
        Exit Sub
      End If
    End If
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = TRIM$(ECOARCH$("TATIMAT", Chr$(TMT%)))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If MODO% = 1 Then
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
      Else
        Screen.MousePointer = 1
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    'Consulta sql de filtro
    SQLX$ = "SELECT * FROM Master "
    SQLX$ = SQLX$ + "WHERE CODIGO>='" & cod1$ & "' "
    SQLX$ = SQLX$ + "AND CODIGO<='" & cod2$ & "' "
    SQLX$ = SQLX$ + "AND STATUS>=1 "
    SQLX$ = SQLX$ + "AND (" & TMT% & " =0 OR TIMATE = " & TMT% & ") "
    SQLX$ = SQLX$ + "ORDER BY CODIGO ASC;"
    ' Seteo el recordset
    Set MasterTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
    
    J& = 0: FIN& = NUINV%
    With MasterTemp
       On Error Resume Next
       .MoveLast
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 99
       End If
       
       registros& = .RecordCount
       FIN& = registros&
       '
       .MoveFirst
       Do While (.EOF = False)
          CI% = !CODINT
          If MODO% = 1 Then
                'Genera el listado de Codigos y Atributos
                ML$ = REGBLAN$("LIMASEL")
                Call REPLA(ML$, MKI$(CI%), 1, 2)
                Call REPLA(ML$, Chr$(TIMATE%(CI%)), 3, 1)
                Call REPLA(ML$, Chr$(CRITREP%(CI%)), 4, 1)
                Call REPLA(ML$, Chr$(MONECOSTO%(CI%)), 5, 1)
                Call REPLA(ML$, MKI$(STOMIN(CI%)), 15, 2)
                Call REPLA(ML$, MKI$(LOTESTAN(CI%)), 17, 2)
                Call REPLA(ML$, UNIMED$(CI%), 19, 2)
                Call REPLA(ML$, MKS$(COSUNI(CI%)), 21, 4)
                Call REPLA(ML$, MKI$(FECOSTO%(CI%)), 25, 2)
                Call REPLA(ML$, MKI$(PROVHABI(CI%)), 33, 2)
                Call REPLA(ML$, MKI$(GRUCOFAB(CI%)), 35, 2)
                JJ& = JJ& + 1: Call GRAREG("LIMASEL", ML$, JJ&)
          End If
          '
       .MoveNext
145    Loop
    End With
    '
149 FIECO% = 0
    If MODO% = 1 Then
        Call SETULTREG("LIMASEL", JJ&)
        Call CIERRARCH("LIMASEL")
        Call FINAL("*LIMASEL")
    End If
    '
99  Screen.MousePointer = 1
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

Private Sub Configral_Click()
  Call Command37_Click
End Sub

Private Sub CoReItBa_Click()
Call OPLIMAS04.Command3_Click
End Sub

Private Sub DepAlm_Click()
  SETGEMA04.Option1(3).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub DeshCam_Click()
Call Command2_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    'Call TRANSLABELS(Name)
    'Show
    'Refresh
    Call CENTRAFORM(Me)
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    'Call GRATITFOR(Caption)
    '
    Refresh
    If EXISTE%("PLANEL04.EXE") > 0 Then
      Command13.Enabled = True
    End If
    '
    Call CARTAVALISTO
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
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
   On Error Resume Next
   Hide
   On Error GoTo 0
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
   Text18.TEXT = ""
   '
   Text28.TEXT = ""
   '
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   Label43 = ""
   Text11 = ""
   List4.Clear
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
    Call CARGENITEM
    '
    Call CARTACODPRO(CI%) 'CARGO LOS CODIGOS DEL  PROVEEDORES HABI. DEL ARTIC.
    Call COULTCOM(CI%) 'CARGO COSTO DE ULTIMA COMPRA
    '
    MODIFIC%(0) = 0
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
        UNID$ = UNIMED$(CI%)
      Text4.TEXT = UNID$
      Label30.Caption = ECOARCH$("UMEDIDA", UNID$)
         TIMA% = TIMATE%(CI%)
      Text6 = TRIM$(Str$(TIMA%))
      Label29.Caption = ECOARCH$("TATIMAT", Chr$(TIMA%))
      Text18.TEXT = CSTRING$(MKS$(STOMIN(CI%)), 3, 10, 1, 2)
      Text24 = FORMATNUM$(PESUNI(CI%), "F9.5")
         FEX = FEREVI%(CI%)
      Text36 = REVICOD$(CI%)
         '
         OBSX$ = ESPECIFIT$(CI%)
      Text28 = OBSX$
      '
      YACAR%(0) = 1
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

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
   End If
   Call GENECOMAS
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
      MODESCRI% = 0
      DE$ = DESCRIT0$(CI%)
      If TRIM$(Text3) <> "" Then
        If TRIM$(Text3) <> DE$ Then
          DE$ = TRIM$(Text3)
          MODESCRI% = 1
          Master.Edit
          Master!DESCRIPCION = Left$(DE$, 64)
          GoTo 30
        End If
      End If
      Screen.MousePointer = 1
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
   MODESCRI% = 0
   DE$ = DESCRIT0$(CI%)
   Master.Edit
   '
   If TRIM$(Text3) <> "" Then
     If TRIM$(Text3) <> DE$ Then
       DE$ = TRIM$(Text3)
       MODESCRI% = 1
       Master!DESCRIPCION = Left$(DE$, 64)
     End If
   End If
   '
'   TTITPE$ = TRIM$(Text24.TEXT)
'   PESU = Val(TTITPE$)
'   If InStr(TTITPE$, "/") > 0 Then
'      TTQP$ = TTITPE$
'      TTQP$ = REPLACAR$(TTQP$, "/", ".")
'      PESU = (-1) * Abs(Val(TTQP$))
'   End If
'   Master!PESUNI = PESU
   '
   TIMA% = Val(Text6.TEXT)
   Master!TIMATE = TIMA%
   '
   'GRUIVA% = Val(Text35.TEXT)
   'Master!GRUPIVA = GRUIVA
   '
'   CRIT% = Val(Text8.TEXT)
'   Master!CRITREP = CRIT%
   '
   'MONEI% = Val(Text13.TEXT)
   'Master!MONECOS = MONEI%
   '
   'TCOLOR$ = TRIM$(Text22.TEXT)
   'Master!COLOR = TCOLOR$
   '
   'TTALLE$ = TRIM$(Text23.TEXT)
   'Master!TALLE = TTALLE
   '
'   CCOMPO% = Val(Text25.TEXT)
'   Master!DESCOMPO = CCOMPO%
   '
   STOMINI = Val(Text18)
   Master!STOMIN = STOMINI
   '
   LOREPO = Val(Text19)
   Master!LOREPOS = LOREPO
   '
   STOMAX = Val(Text20)
   Master!STOMAX = STOMAX
   '
   UNID$ = TRIM$(Text4.TEXT)
   Master!UMEDIDA = Left$(UNID$, 4)
   '
   'COU = Val(Text12.TEXT)
   'Master!COSUNI = COU
   '
'   FEREVX = Text25.TEXT
'   On Error Resume Next
'   Master!FEREVIS = CDate(FEREVX)
'   On Error GoTo 0
   '
   'PORNA = Val(Text28.TEXT)
   'Master!PORNACIO = PORNA
   '
   'NPRO% = Val(Text10.TEXT)
   'Master!PROVHABI = NPRO%
   'NPROA% = CVI(Mid$(REMAS$, 123, 2))
   '
   'PHABIT%(CI%) = (-1)
   'If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
   '  PHABIT%(CI%) = NPRO%
   'End If
   '
   'GRUCO% = Val(Text16.TEXT)
   'Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   '
   'If Abs(Val(Text28)) > 320 Then
   '    Call MENSERR(24, "% Gastos de Nacionalización No Válido.")
   '  Else
   '    FCPI% = Int(100 * Val(Text28.TEXT) + 0.5)
'  '     Call REPLA(REMAS$, MKI$(FCPI%), 127, 2)
   'End If
   '
   ESPECIF$ = RTrim$(Text28) + " "
   Master!ESPECIFICACION = ESPECIF$
   '
30 Master.Update
   If MODESCRI% > 0 Then
     UPDATMASTER% = 1
   End If
   '
   'If NPRO% <> NPROA% Then  ' cambio de proveedor habitual
   '  ZZX$ = REGBLAN$("RECAMPH")
   '  Call REPLA(ZZX$, MKI$(CI%), 1, 2)
   '  HOII% = HOY(HOS$)
   '  Call REPLA(ZZX$, MKI$(HOII%), 3, 2)
   '  Call REPLA(ZZX$, MKI$(NPROA%), 5, 2)
   '  Call REPLA(ZZX$, MKI$(NPRO%), 7, 2)
   '  Call REPLA(ZZX$, USERTER$, 9, 24)
   '  Call REPLA(ZZX$, "Costos Standard", 33, 32)
   '  Call REGAPP("RECAMPH", ZZX$)
   '  Call CIERRARCH("RECAMPH")
   'End If
   '
   Call CIERRARCH("*.*")
   '
   Call BAJALDISCO
   Screen.MousePointer = 1
   MODIFIC%(0) = 0
   MODESCRI% = 0
   '
End Sub

Private Sub GrCoVe_Click()
OPMAST04.Option1(3).Value = True
Call OPMAST04.Command16_Click
End Sub

Private Sub guar_Click()
Call Command1_Click
End Sub

Private Sub impri_Click()
Call Command39_Click
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ToolTipText = Label5
End Sub

Private Sub Label92_dblClick()
  '
  CI1% = CODINT%(Text1.TEXT)
  If CI1% > 0 Then
    Call FILTERFILE("REPLACOP", "!REPLACOP", 3, MKI$(CI1%))
    '
    ATRI$ = ATRI$ + "/TITUPAN= Códigos Originales - " + Text1.TEXT + " - " + Text3.TEXT
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAS_MAQ04.Left + 4000))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAS_MAQ04.Top + 2250))
    VTB1% = VENTABU%("REPLACOP" + ATRI$)
    If VTB1% >= 0 Then
      '
      List4.Clear
      Call FILTERUPDATE("REPLACOP", "", 3, MKI$(CI1%))
      Call CARTACODPRO(CI1%)
      '
    End If
  End If
  '
End Sub

Private Sub LiCapInv_Click()
  Call OPLIMAS04.COMMAND15_CLICK
End Sub

Private Sub LiCoAtri_Click()
  Call OPLIMAS04.Command9_Click
End Sub

Private Sub LiItIn_Click()
Call OPLIMAS04.Command2_Click
End Sub

Private Sub List4_dblClick()
  '
  Call Label92_dblClick
  '
End Sub

Private Sub LiStMin_Click()
Call OPLIMAS04.Command7_Click
End Sub

Private Sub ListAtGe_Click()
    '
    Call Command38_Click
    '
End Sub

Private Sub MayEq_Click()
  SETGEMA04.Option1(5).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub MnuPlant_Click()
  SETGEMA04.Option1(1).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub nueite_Click()
Call Command21_Click
End Sub

Private Sub Sali_Click()
Call Command3_Click(4)
End Sub

Private Sub SeccyCen_Click()
  SETGEMA04.Option1(2).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub SuCoAr_Click()
Call Command23_Click
End Sub

Private Sub TaTiMa_Click()
  SETGEMA04.Option1(0).Value = True
  Call SETGEMA04.Command5_Click
End Sub

Private Sub TaUnMe_Click()
  SETGEMA04.Option1(6).Value = True
  Call SETGEMA04.Command5_Click
End Sub

'
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
   Text24.TEXT = CSTRING$(MKS$(Val(Text24.TEXT)), 3, 10, 5, 2)
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


Private Sub Text28_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

'
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

Sub ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER
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
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
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
   Call ABREMASTER
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER
   Master.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Codint = " & CODII%
     Master.FindFirst ABUS$
     If Master.NoMatch = False Then
       CODII% = CODII% + 1
       GoTo 20
     End If
   End If
   '
   With Master
      .AddNew
      !Codigo = CODINUE$
      !DESCRIPCION = DENUE$
      !STATUS = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
      .Bookmark = .LastModified
      .Close
      UPDATMASTER% = 1
   End With
   '
   REMAS& = NUMAS% + 2
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
      TTXX$ = Space$(63)
      Call REPLA(TTXX$, "DEPOSITO CENTRAL", 1, 20)
      Call REPLA(TTXX$, "DEPOSITO", 21, 8)
      Call GRAREG("TADEPOSI", Chr$(1) + TTXX$, 1): Call CIERRARCH("TADEPOSI")
      Call GRAREG("ECODEP", Chr$(1) + Left$(TTXX$, 20), 1): Call CIERRARCH("ECODEP")
      Call GRAREG("SUCURSAL", TTXX$, 1): Call CIERRARCH("SUCURSAL")
      '
      '  Screen.MousePointer = 1
      '  Call MENSERR(24, "Falta Definir Depósitos")
      '  OPMAST00.Show 1
      '  GoTo 10
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
     codex$ = Left$(Z$, 16)
     If Mid$(Z$, 16, 1) <> "`" Then
       codex$ = TRIM$(codex$)
       If codex$ >= cod1$ Then
         If codex$ <= cod2$ Then
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
   IMPLOSI.Show 1
   '
End Sub
Sub COULTCOM(CIP1%) 'CARGA EL COSTO DE LA ULTIMA COMPRA
  '
  Call ABRECOMPRAS
  SQLX$ = " SELECT * FROM OCOMDETA "
  SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIP1% & " "
  SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
  SQLX$ = SQLX$ + " ORDER BY Femi_Ocom DESC "
  Set COTMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
  With COTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Label5 = ""
        Exit Sub
      End If
      On Error GoTo 0
      COULCO1 = !PrUn_Neto * TICAMONE(!Mone_Cos)
      Label5 = FORMATNUM$(COULCO1, "F9.2")
  End With
  '
End Sub
Sub CARTACODPRO(CIP1%) 'CARGA LOS CODIGOS ORIGINALES DE PROVEEDORES HABITURALES.
  '
  List4.Clear
  For J& = 1 To ULTREG&("REPLACOP")
    X$ = REGLEIDO$("REPLACOP", J&)
    CI1% = CVI(Mid$(X$, 3, 2))
    If CI1% = CIP1% Then
      NP% = CVI(Mid$(X$, 1, 2))
      NPRO% = NP% * (-1)
      RASO1$ = RASOCLI$(NPRO%)
      CODOR1$ = Mid$(X$, 5, 28)
      L1$ = Space(64)
      Call REPLA(L1$, RASO1$, 1, 24)
      Call REPLA(L1$, CODOR1$, 26, 28)
      Call REPLA(L1$, MKI$(NP%), 100, 2)
      List4.AddItem L1$
    End If
  Next J&
  '
End Sub
