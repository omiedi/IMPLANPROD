VERSION 5.00
Begin VB.Form ALTAMODE 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base de Datos de Modelos"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1500
      Left            =   105
      TabIndex        =   52
      Top             =   6720
      Width           =   10620
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   55
         Top             =   315
         Width           =   10580
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   57
            Top             =   75
            Width           =   10200
         End
         Begin VB.Label Label10 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   56
            Top             =   420
            Width           =   1380
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "O.K"
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
         Left            =   8295
         Picture         =   "ALTAMODE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   760
         Width           =   540
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   735
         Width           =   7155
      End
      Begin VB.Label Label16 
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
         Height          =   750
         Left            =   8295
         TabIndex        =   60
         Top             =   735
         Width           =   2265
      End
      Begin VB.Label Label15 
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
         Height          =   750
         Left            =   0
         TabIndex        =   59
         Top             =   735
         Width           =   1170
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Label Label13 
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
         Height          =   225
         Left            =   8400
         TabIndex        =   58
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1480
         Y2              =   1480
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   10920
      TabIndex        =   47
      ToolTipText     =   "Acceso a Maestro de Items de Stock"
      Top             =   3675
      Width           =   875
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
         Top             =   900
         Width           =   650
      End
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   300
         Width           =   650
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0FFD0&
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
      Height          =   3525
      Left            =   10920
      TabIndex        =   41
      Top             =   105
      Width           =   875
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Baja del Modelo Activo"
         Top             =   2800
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Cierra y Sale de la Aplicación"
         Top             =   300
         Width           =   650
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1500
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   600
         Left            =   105
         Picture         =   "ALTAMODE.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Alta de Nuevo Modelo"
         Top             =   2200
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   600
         Left            =   105
         Picture         =   "ALTAMODE.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   900
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
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
      ForeColor       =   &H00808000&
      Height          =   1335
      Left            =   5775
      TabIndex        =   26
      Top             =   2100
      Width           =   4950
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   875
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   315
         Width           =   4530
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Modelos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1905
      Left            =   105
      TabIndex        =   14
      Top             =   105
      Width           =   10620
      Begin VB.CommandButton Command12 
         Height          =   520
         Left            =   9870
         Picture         =   "ALTAMODE.frx":1608
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Acceso a B.Datos de Fotos y Planos"
         Top             =   250
         Width           =   540
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
         Left            =   9450
         TabIndex        =   62
         Text            =   " "
         Top             =   840
         Width           =   960
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00E0FFD0&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         Height          =   645
         Left            =   105
         TabIndex        =   28
         Top             =   1200
         Width           =   10410
         Begin VB.OptionButton Option8 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Anual"
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
            Left            =   7035
            TabIndex        =   29
            Top             =   45
            Width           =   960
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Trim."
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
            Left            =   6090
            TabIndex        =   31
            Top             =   45
            Width           =   1380
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
            Height          =   285
            Left            =   1155
            TabIndex        =   34
            Text            =   " "
            Top             =   210
            Width           =   2220
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Seman."
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
            Left            =   4935
            TabIndex        =   33
            Top             =   45
            Width           =   1380
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Mens."
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
            Left            =   5460
            TabIndex        =   32
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Semes"
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
            Left            =   6510
            TabIndex        =   30
            Top             =   270
            Width           =   1380
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Ubicación:"
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
            TabIndex        =   40
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Frecuencia de Verificación:"
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
            Left            =   3570
            TabIndex        =   39
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "U.Control:"
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
            Left            =   8295
            TabIndex        =   38
            Top             =   45
            Width           =   960
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0FFD0&
            Caption         =   "Prox.Control:"
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
            Left            =   7980
            TabIndex        =   37
            Top             =   345
            Width           =   1275
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
            Height          =   255
            Left            =   9345
            TabIndex        =   36
            Top             =   0
            Width           =   960
         End
         Begin VB.Label Label9 
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
            Height          =   255
            Left            =   9345
            TabIndex        =   35
            Top             =   315
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0FFD0&
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   600
         Left            =   3780
         TabIndex        =   18
         Top             =   180
         Width           =   6000
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Otros:"
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
            Left            =   2730
            TabIndex        =   23
            Top             =   270
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Madera"
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
            TabIndex        =   21
            Top             =   270
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0FFD0&
            Caption         =   "Aluminio"
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
            Left            =   1470
            TabIndex        =   20
            Top             =   270
            Width           =   1275
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
            Left            =   3570
            TabIndex        =   19
            Text            =   " "
            Top             =   210
            Width           =   2115
         End
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
         Height          =   280
         Left            =   3465
         TabIndex        =   17
         Top             =   420
         Width           =   180
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
         Height          =   285
         Left            =   1260
         TabIndex        =   16
         Text            =   " "
         Top             =   840
         Width           =   6840
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
         Left            =   1260
         TabIndex        =   15
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFD0&
         Caption         =   "Cant.Piezas:"
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
         Left            =   7455
         TabIndex        =   63
         Top             =   900
         Width           =   1905
      End
      Begin VB.Label Ltota 
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
         Index           =   2
         Left            =   -735
         TabIndex        =   61
         Top             =   900
         Width           =   1905
      End
      Begin VB.Label Ltota 
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
         Height          =   330
         Index           =   0
         Left            =   -735
         TabIndex        =   22
         Top             =   480
         Width           =   1905
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Opcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   10920
      TabIndex        =   8
      Top             =   5355
      Width           =   875
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
         Left            =   105
         Picture         =   "ALTAMODE.frx":1912
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Transacciones de Inspección"
         Top             =   900
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
         Height          =   600
         Left            =   105
         Picture         =   "ALTAMODE.frx":1C1C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Movimientos - Entrega y Recepción"
         Top             =   300
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
         Height          =   600
         Left            =   105
         Picture         =   "ALTAMODE.frx":1F26
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consultas y Listados Base de Datos de Modelos"
         Top             =   1500
         Width           =   650
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Pieza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1335
      Left            =   105
      TabIndex        =   5
      Top             =   2100
      Width           =   5580
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3465
         TabIndex        =   11
         Top             =   315
         Width           =   180
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
         Height          =   285
         Left            =   1260
         TabIndex        =   10
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label Ltota 
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
         Height          =   330
         Index           =   1
         Left            =   -735
         TabIndex        =   25
         Top             =   375
         Width           =   1905
      End
      Begin VB.Label Label5 
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
         Height          =   435
         Left            =   1260
         TabIndex        =   24
         Top             =   735
         Width           =   4110
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0FFD0&
         Caption         =   "U.Med:"
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
         Left            =   3780
         TabIndex        =   13
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label3 
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
         Left            =   4830
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3090
      Left            =   105
      TabIndex        =   0
      Top             =   3570
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   4005
         Left            =   10290
         TabIndex        =   7
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   4
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   3
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   0
         TabIndex        =   2
         Top             =   735
         Width           =   10575
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   1
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   6
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "ALTAMODE.frx":2230
      Top             =   7710
      Width           =   2010
   End
End
Attribute VB_Name = "ALTAMODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIDES%
'
Private Sub Command1_Click()
    Command1.Enabled = False
    '
10  VDEF$ = String$(66, 0)
    '
15  OKX% = 0
    OBX$ = OBJPLA$("ALTAMODELO", VDEF$)
    If OBX$ = "" Then GoTo 99
    VDEF$ = OBX$
    MODENUE$ = TRIM$(Mid$(OBX$, 1, 16))
    DESCNUE$ = TRIM$(Mid$(OBX$, 17, 48))
    '
    If MODENUE$ = "" Then
      Call COMUNI("Falta Codigo de Modelo")
      GoTo 15
    End If
    '
    If DESCNUE$ = "" Then
      Call COMUNI("Falta Descripción del Modelo")
      GoTo 15
    End If
    '
    FULINSP% = CVI(Mid$(OBX$, 65, 2))
    If FULINSP% < 1 Or FULINSP% > HOY(HOS$) Then
      Call MENSERR(24, "Mandatorio Indicar Fecha de Ultima Inspección")
      GoTo 15
    End If
    '
    If ECOARCH$("BADAMODE", MODENUE$) = "" Then
      X$ = REGBLAN$("BADAMODE")
      Call REPLA(X$, MODENUE$, 1, 16)
      Call REPLA(X$, DESCNUE$, 17, 48)
      ' Valores por omision para Frecuencia, Ult.Control y Prox.Control
      Call REPLA(X$, "1 - Semanal", 97, 16)
        FEX = FULINSP%
      Call REPLA(X$, MKI$(FULINSP%), 113, 2)
      Call REPLA(X$, MKI$(0), 115, 2)
      Call REGAPP("BADAMODE", X$)
      Call FRESHECHO("BADAMODE")
    Else
      Call COMUNI("Código de Modelo Existente")
      GoTo 10
    End If
    '
    Text3 = Mid$(X$, 1, 16)
    '
99  Command1.Enabled = True
End Sub

Private Sub COMMAND21_CLICK()
   Command21.Enabled = False
   '
   VERIMODE.Show 1
'   Call COPYSTRU("INDIPEDI", "PEPENDES")
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
'   For U& = 1 To FIN& Step 4
'     Call TRACE(20, U&, FIN&)
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     NPX& = CVS(Left$(XX$, 4))
'     If NPX& <> NPXA& Then
'       NPXA& = NPX&
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
'       FEPE = CVI(Mid$(XX$, 5, 2))
'       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
'       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
'       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
'       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
'       JJ& = JJ& + 1
'       Call GRAREG("PEPENDES", YY$, JJ&)
'     End If
'   Next U&
'   '
'   Call SETULTREG("PEPENDES", JJ&)
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
'   NPX& = Val(VALACT1$("PEPENDES"))
'   If NPX& < 1 Then GoTo 99
'   '
'   Label7 = TRIM$(Str$(NPX&))
'   '
99 Command21.Enabled = True
End Sub

Private Sub Command10_Click()
   '
   Command10.Enabled = False
   If ECOARCH("BADAMODE", Text3) <> "" Then
     CODMODE$ = AJUSTI$(Text3, 16)
     '
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERNAME$, 69, 24)
         End If
         Call REPLA(CODMODE$, ALTAMODE.Text3, 1, 16)
         Call REGAPP("ANOTAMOD", CODMODE$ + TTXX$)
       Next KKI%
       Call REGAPP("ANOTAMOD", CODMODE$ + String$(88, "="))
       Call CIERRARCH("ANOTAMOD")
       Text1 = ""
       '
       ALTAMODE.List1.Clear
       RFF$ = RECFILT$("ANOTAMOD", 1, CODMODE$)
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAMOD", RENOTA&)
         ALTAMODE.List1.AddItem Mid$(XX$, 17, 92)
       Next U&
       '
     End If
   End If
   '
   Text1 = ""
   '
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   OPX% = COMALTER%("Base de Datos Maestra a Actualizar:\\Maestro de Items de Stock (Piezas)\Proveedores y Contratistas")
   If OPX% = 1 Then
       Call CONECRUN("ARCH_GES/AMASTO", "Archivos Maestros")
     ElseIf OPX% = 2 Then
       Call CONECRUN("ARCH_GES/AMAPRO", "Archivos Maestros")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If ECOARCH("BADAMODE", Text3) <> "" Then
     Call FORPLANO.CARPLANOS
     FORPLANO.Show 1
   End If
   Command12.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   OPX% = COMALTER%("Opciones de Listados:\\B.Datos Completa\Pendientes de Inspección\Ficha de Inspección")
   CORRESPO$ = "Corresponde a: "
   Select Case OPX%
     Case 1    ' Base de Datos Completa
       Call COPYSTRU("BADAMODE", "BADAMODX")
       '
       CORRESPO$ = CORRESPO$ + "B.Datos Completa - "
       FTX% = COMALTER%("Filtrado por Tipo de Material:\\Todos\Madera\Aluminio\Otros") - 1
         If FTX% < 0 Or FTX% > 3 Then GoTo 99
       If FTX% < 1 Then
           CORRESPO$ = CORRESPO$ + "Todos -"
           Call CIERRARCH("BADAMODE")
           Call CIERRARCH("BADAMODX")
           Call BAJALDISCO
           Call COPYFILE(LUDAT$ + "BADAMODE.DAT", LUDAT$ + "BADAMODX.DAT")
         Else
           If FTX% = 1 Then CORRESPO$ = CORRESPO$ + "Madera - "
           If FTX% = 2 Then CORRESPO$ = CORRESPO$ + "Aluminio - "
           If FTX% = 3 Then CORRESPO$ = CORRESPO$ + "Otros - "
           Call FILTERFILE("BADAMODE", "BADAMODX", 7, TRIM$(Str$(FTX%)))
       End If
       '
       If ULTREG&("BADAMODX") < 1 Then
         Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
         GoTo 99
       End If
       '
       OPY% = COMALTER%("Opciones Alternativas de Ordenamiento:\\Por Modelo\Por Pieza\Por Fecha Próximo Control")
         If OPY% < 1 Or OPY% > 3 Then GoTo 99
       ASORT% = 1: CORRESPO$ = CORRESPO$ + "Por Modelo"
       If OPY% = 2 Then ASORT% = 9: CORRESPO$ = CORRESPO$ + "Por Pieza"
       If OPY% = 3 Then ASORT% = 6: CORRESPO$ = CORRESPO$ + "Por Fecha Proximo Control"
       Call FILESORT("BADAMODX", "", ASORT%, ASORT%, "ASC")
       '
       Call HEADERS("BADAMODX", "")
       Call HEADERS("BADAMODX", CORRESPO$)
       Call FINAL("*LIBAMOD")
       '
     Case 2    ' Pendientes de Inspeccion
       Call COPYSTRU("BADAMODE", "BADAMODX")
       '
       Call SELECHO("SELFECHA/Prox.Insp.Hasta")
       VDEVU$ = VALACT1$("SELFECHA")
       If VDEVU$ <> "" Then
         '
         CORRESPO$ = CORRESPO$ + "Proximo Control Hasta " + VDEVU$ + " - "
         ' FILTRADO POR FECHA
         FLIM% = FECHANUM(VDEVU$)
         FIN& = ULTREG&("BADAMODE")
         Call COPYSTRU("BADAMODE", "BADAMODY")
         JJ& = 0
         For U& = 1 To FIN&
           XX$ = REGLEIDO$("BADAMODE", U&)
           If CVI(Mid$(XX$, 115, 2)) <= FLIM% Then
             JJ& = JJ& + 1
             Call GRAREG("BADAMODY", XX$, JJ&)
           End If
         Next U&
         Call SETULTREG("BADAMODY", JJ&)
         Call CIERRARCH("BADAMODY")
         '
         If ULTREG&("BADAMODY") < 1 Then
           Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
           GoTo 99
         End If
         '
         FTX% = COMALTER%("Filtrado por Tipo de Material:\\Todos\Madera\Aluminio\Otros") - 1
           If FTX% < 0 Or FTX% > 3 Then GoTo 99
         If FTX% < 1 Then
             CORRESPO$ = CORRESPO$ + "Todos -"
             Call CIERRARCH("BADAMODE")
             Call CIERRARCH("BADAMODX")
             Call BAJALDISCO
             Call COPYFILE(LUDAT$ + "BADAMODY.DAT", LUDAT$ + "BADAMODX.DAT")
           Else
             If FTX% = 1 Then CORRESPO$ = CORRESPO$ + "Madera - "
             If FTX% = 2 Then CORRESPO$ = CORRESPO$ + "Aluminio - "
             If FTX% = 3 Then CORRESPO$ = CORRESPO$ + "Otros - "
             Call FILTERFILE("BADAMODY", "BADAMODX", 7, TRIM$(Str$(FTX%)))
         End If
         '
         If ULTREG&("BADAMODX") < 1 Then
           Call MENSERR(24, "Sin Registros que Satisfacen\Criterio de Filtrado")
           GoTo 99
         End If
         '
         OPY% = COMALTER%("Opciones Alternativas de Ordenamiento:\\Por Modelo\Por Pieza\Por Fecha Próximo Control")
           If OPY% < 1 Or OPY% > 3 Then GoTo 99
         ASORT% = 1: CORRESPO$ = CORRESPO$ + "Por Modelo"
         If OPY% = 2 Then ASORT% = 9: CORRESPO$ = CORRESPO$ + "Por Pieza"
         If OPY% = 3 Then ASORT% = 6: CORRESPO$ = CORRESPO$ + "Por Fecha Proximo Control"
         Call FILESORT("BADAMODX", "", ASORT%, ASORT%, "ASC")
         '
         Call HEADERS("BADAMODX", "")
         Call HEADERS("BADAMODX", CORRESPO$)
         Call FINAL("*LIBAMOD")
         '
       End If
       '
     Case 3
        '
        FORIPRE$ = TRIM$(CONTROL$("", "FICHINMO"))
        If FORIPRE$ = "" Then Exit Sub
        '
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
          Call MENSERR(24, "Imposible Imprimir Solicitud Interna de Materiales.\No Existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
          Exit Sub
        End If
        '
        Screen.MousePointer = 11
        '
        If TRIM$(Text3) = "" Then
           Call MENSERR(24, "Imposible Imprimir Ficha de Inspección. Falta Código de Modelo")
           Screen.MousePointer = 0
           Exit Sub
        End If
        '
        CODTABU1$(1) = "FECH-LAN"
        CODTABU1$(2) = "REF-MAT3"
        CODTABU1$(3) = "NRO-LOTE"
        CODTABU1$(4) = "ZONA-VEN"
        CACOLTAB1% = 4
        '
        CAITAB1% = 0
        For I = 0 To (List1.ListCount - 1)
          CAITAB1% = CAITAB1% + 1
          If Left$(List1.List(I), 1) <> "=" Then
            FECHA$ = Mid$(List1.List(I), 1, 8)
            DETALLE$ = Mid$(List1.List(I), 11, 56)
            USUARIO$ = Mid$(List1.List(I), 69, 17)
            TETABU1$(1, CAITAB1%) = FECHA$
            TETABU1$(2, CAITAB1%) = DETALLE$
            TETABU1$(3, CAITAB1%) = USUARIO$
            TETABU1$(4, CAITAB1%) = ""
          Else
            FECHA$ = String$(8, "=")
            DETALLE$ = String$(60, "=")
            USUARIO$ = String$(13, "=")
            FIRMA$ = String$(13, "=")
            TETABU1$(1, CAITAB1%) = FECHA$
            TETABU1$(2, CAITAB1%) = DETALLE$
            TETABU1$(3, CAITAB1%) = USUARIO$
            TETABU1$(4, CAITAB1%) = FIRMA$
          End If
        Next I
        '
        'CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = "001"
        CODPARAM$(2) = "REF-MAT1": DOCPARAM$(2) = TRIM$(Text5)
        CODPARAM$(3) = "DES-MAT"
        If Option1.Value Then
          DOCPARAM$(3) = "Madera"
        ElseIf Option2.Value Then
          DOCPARAM$(3) = "Alumnio"
        Else
          DOCPARAM$(3) = TRIM$(Text4)
        End If
        CODPARAM$(4) = "COD-MAT": DOCPARAM$(4) = TRIM$(Text3)
        CODPARAM$(5) = "UBI-MPRI": DOCPARAM$(5) = TRIM$(Text7)
        CODPARAM$(6) = "CANTIDAD": DOCPARAM$(6) = TRIM$(Text8)
        CODPARAM$(7) = "REF-MAT2"
        If Option4.Value Then
          DOCPARAM$(7) = "Semanal"
        ElseIf Option5.Value Then
          DOCPARAM$(7) = "Mensual"
        ElseIf Option6.Value Then
          DOCPARAM$(7) = "Trimestral"
        ElseIf Option7.Value Then
          DOCPARAM$(7) = "Semestral"
        Else
          DOCPARAM$(7) = "Anual"
        End If
        CAPARDOC% = 7
        '
        Call BLANARCH("!OBSERVOF")
        Call FRATEXTO(Text6.TEXT, 35)
        If CARETEX% >= 1 Then Call GRAREG("!OBSERVOF", RETEX$(1), 1)
        If CARETEX% >= 2 Then Call GRAREG("!OBSERVOF", RETEX$(2), 2)
        If CARETEX% >= 3 Then Call GRAREG("!OBSERVOF", RETEX$(3), 3)
        If CARETEX% >= 4 Then Call GRAREG("!OBSERVOF", RETEX$(4), 4)
        URETE& = CARETEX%
        Call SETULTREG("!OBSERVOF", URETE&)
        Call CIERRARCH("!OBSERVOF")
        '
        FORIPRE$ = TRIM$(CONTROL$("", "FICHINMO"))
        If FORIPRE$ <> "" Then
          Call PRINTDOC("FICHINMO")
        End If
        '
     Case Else
        GoTo 99
        '
   End Select
   '
99 Command23.Enabled = True
'   '
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If ECOARCH("BADAMODE", Text3) <> "" Then
     OPX% = ALTERNA%("Despachar Modelo A Proveedor\Recibir Modelo de Proveedor")
     If OPX% = 1 Then
         Call DESPAMODE
       ElseIf OPX% = 2 Then
         Call RECIMODE
     End If
   End If
   Command24.Enabled = True
End Sub

Private Sub Command26_Click()
'   Command26.Enabled = False
'   CFGDESPA.Show 1
'   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   '
   If ECOARCH("BADAMODE", Text3) <> "" Then
     X$ = FILTERGETRECORD("BADAMODE", 1, Text3)
     Call REPLA(X$, Text2, 17, 48) ' Descripcion del modelo
     ' Material
     If Option1.Value Then Call REPLA(X$, "1", 117, 2)
     If Option2.Value Then Call REPLA(X$, "2", 117, 2)
     If Option3.Value Then ' Selecciono Otros
       Call REPLA(X$, "3", 117, 2)
       Call REPLA(X$, Text4, 119, 16)
     End If
     Call REPLA(X$, Text7, 65, 32) ' Ubicacion
     ' Frecuencia
     If Option4.Value Then Call REPLA(X$, "1 - Semanal", 97, 16)
     If Option5.Value Then Call REPLA(X$, "2 - Mensual", 97, 16)
     If Option6.Value Then Call REPLA(X$, "3 - Trimestral", 97, 16)
     If Option7.Value Then Call REPLA(X$, "4 - Semestral", 97, 16)
     If Option8.Value Then Call REPLA(X$, "5 - Anual", 97, 16)
     ' Valida que el codigo sea valido
     CIXI% = CODINT%(Text5)
     If Text5 <> "" Then
       If CIXI% < 1 Then
         Call COMUNI("Código de Pieza No Válido")
         GoTo 99
       End If
     End If
     '
     FFCONT% = FECHANUM(Label6)             'SE AGREGA GRABACION DE FECHA DE ÚLTIMO CONTROL
     Call REPLA(X$, MKI$(FFCONT%), 113, 2)
     '
     PROCON% = FECHANUM(Label9)
     Call REPLA(X$, MKI$(PROCON%), 115, 2)
     Call REPLA(X$, MKI$(CIXI%), 135, 2)
     '
     CANPIEZAS = Val(Text8)
     Call REPLA(X$, MKS$(CANPIEZAS), 137, 4)
     ' Graba el registro
     Call FILTERPUTRECORD("BADAMODE ", 1, Text3, X$)
     '
     TTYZ$ = Text6.TEXT
     LOCADE& = Len(TTYZ$)
     For U& = LOCADE& To 1 Step -1
       If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
     Next U&
     U& = 0
     '
10   TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
     Call FRATEXTO(TEXTO$, 34)
     For LL% = 1 To CARETEX%
       TTYY$ = AJUSTI$(Text3, 16) + RETEX$(LL%)
       Call GRAREG("!OBMODELO", TTYY$, CLng(LL%))
     Next LL%
     Call SETULTREG("!OBMODELO", CLng(LL%) - 1)
     '
     Call FILTERUPDATE("OBMODELO", "", 1, Text3)
     '
   Else
     Call COMUNI("Error en Código de Modelo")
   End If
   '
   If MODIDES% > 0 Then
     Call FRESHECHO("BADAMODE")
     MODIDES% = 0
   End If
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command1.Enabled = False
   '
   X$ = REGSEL$("BADAMODE")
   If TRIM$(X$) = "" Then GoTo 99
   '
   Text3 = TRIM$(Mid$(X$, 1, 16))
   '
99 Command1.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   OPX% = ALTERNA%("Configuración General\Formularios")
   If OPX% = 1 Then
       Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
       Call FINAL("")
     ElseIf OPX% = 2 Then
       OPALMOD00.Show 1
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   If ECOARCH("BADAMODE", Text3) <> "" Then
     If DERACCE%("BAJAMODE", "Dar de Baja Modelo de Almacén") >= 2 Then
       If COMALTER%("Confirme la Baja Definitiva del Modelo\'" + TRIM$(Text3) + "'.\\Cancelar Baja\Confirmar") = 2 Then
         Call FILTERDELETE("BADAMODE", 1, AJUSTI$(Text3, 16))
         Call COMUNI("Baja Definitiva Procesada")
         Call BLANFORMU
         Text2 = ""
         Text3 = ""
         Call FRESHECHO("BADAMODE")
       End If
     End If
   End If
99 Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If ECOARCH("BADAMODE", Text3) <> "" Then
     If DERACCE%("REGINSPE", "Registro de Inspección de Modelos") >= 2 Then
       With NOVEMODE
         .Text1 = ""
         .Label3 = Text3
         .Label6 = Text2
         .Label7 = Text5
         .Label8 = Label5
         .Text10 = FECHATEX$(HOY(HOS$))
         .Label9 = USERNAME
         .Option1.Value = True
       End With
       '
       NOVEMODE.Show 1
       '
       CODMODE$ = AJUSTI$(Text3, 16)
       List1.Clear
       RFF$ = RECFILT$("ANOTAMOD", 1, CODMODE$)
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAMOD", RENOTA&)
         List1.AddItem Mid$(XX$, 17, 92)
       Next U&
     End If
   End If
   '
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   X$ = REGSEL$("MASTER")
   If TRIM$(X$) = "" Then GoTo 99
   '
   Text5 = TRIM$(Mid$(X$, 1, 15))
   '
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    DoEvents
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    'Label9 = " " + HOS$
    'Label12 = Left$(USERNAME, 17)
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label2_Click()
    '
    If ECOARCH$("BADAMODE", ALTAMODE.Text3) <> "" Then
      NOVEMODE.Show 1
    Else
      Call COMUNI("Error en Código de Modelo")
    End If
    '
End Sub

Private Sub Label7_Change()
'   List1.Clear
'   If Val(Label7) < 1 Then Exit Sub
'   '
'   NPX& = Val(Label7)
'   RFF$ = RECFILT$("PEDENCA", 1, MKS$(NPX&))
'   If Len(RFF$) >= 4 Then
'     REPED& = CVS(Left$(RFF$, 4))
'     XX$ = REGLEIDO$("PEDENCA", REPED&)
'     FEX = CVI(Mid$(XX$, 53, 2))
'     Label8 = FECHATEX$(FEX)
'     NCLIE% = CVI(Mid$(XX$, 55, 2))
'     Label11 = RASOCLI$(NCLIE%)
'     Label16 = Mid$(XX$, 57, 16)
'     Label18 = TRIM$(Str$(Asc(Mid$(XX$, 267, 1))))
'     '
'     LENTREGA(0) = Mid$(XX$, 73, 32)
'     LENTREGA(1) = Mid$(XX$, 105, 32)
'     LENTREGA(2) = Mid$(XX$, 137, 32)
'   End If
'   '
'   Label14.ForeColor = RGB(0, 0, 255)
'   Label14 = "Venta"
'   RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NPX&))
'   For UKKK& = 1 To Len(RFQ$) Step 4
'     UU& = CVS(Mid$(RFQ$, UKKK&, 4))
'     XX$ = REGLEIDO$("PEDDETA", UU&)
'     If CVS(Left$(XX$, 4)) = NPX& Then
'       If UCase$(Mid$(XX$, 51, 1)) = "R" Then
'         Label14.ForeColor = RGB(255, 0, 0)
'         Label14 = "Reparación"
'       End If
'       GoTo 20
'     End If
'   Next UKKK&
'   '
'20 RFF$ = RECFILT$("ANOTAPED", 1, MKS$(NPX&))
'   For U& = 1 To Len(RFF$) Step 4
'     RENOTA& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("ANOTAPED", RENOTA&)
'     List1.AddItem Mid$(XX$, 5, 92)
'   Next U&
'   '
End Sub

Sub DESPAMAT()
'   '
'   Call COPYSTRU("INDIPEDI", "PEPENDES")
'   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
'   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
'   For U& = 1 To FIN& Step 4
'     Call TRACE(20, U&, FIN&)
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     NPX& = CVS(Left$(XX$, 4))
'     If NPX& <> NPXA& Then
'       NPXA& = NPX&
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
'       FEPE = CVI(Mid$(XX$, 5, 2))
'       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
'       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
'       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
'       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
'       JJ& = JJ& + 1
'       Call GRAREG("PEPENDES", YY$, JJ&)
'     End If
'   Next U&
'   '
'   Call SETULTREG("PEPENDES", JJ&)
'   If JJ& < 1 Then
'     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
'     GoTo 99
'   End If
'   '
'   If Val(Label7) > 0 Then
'     NPX& = Val(Label7)
'     RFF$ = RECFILT$("PEPENDES", 1, MKS$(NPX&))
'     If Len(RFF$) >= 4 Then GoTo 5
'   End If
'   '
'3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
'   NPX& = Val(VALACT1$("PEPENDES"))
'   If NPX& < 1 Then GoTo 99
'   '
'5  Label7 = TRIM$(Str$(NPX&))
'   '
'   FORIPRE$ = TRIM$(CONTROL$("", "CDESPACH"))
'   If FORIPRE$ <> "" Then
'     LU$ = LUDAT$
'     If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
'       RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
'       Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
'       Exit Sub
'     End If
'   End If
'   '
'   DEPOCON% = DEPOEXPE%
'   DEPORIG% = DEPOCON%
'   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'   '
'   Call CIERRARCH("*.*")
'   Call BLANARCH("!BOLREDET")
'   '
'   Screen.MousePointer = 11
'   NPED& = Val(Label7)
'   REBLA$ = REGBLAN$("DESPAPED")
'   RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
'   JJ& = 0
'   '
'   For U& = 1 To Len(RFF$) Step 4
'     REDE& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("DESPAPED", REDE&)
'     STATU% = Asc(Mid$(XX$, 27, 1))
'     If STATU% < 2 Then
'       REPARX% = 0: If Mid$(XX$, 51, 1) = "R" Then REPARX% = 1
'       CI% = CVI(Mid$(XX$, 33, 2))
'       NCLIE% = CVI(Mid$(XX$, 7, 2))
'       CAPENE# = CVS(Mid$(XX$, 75, 4))
'       FENTRESO% = CVI(Mid$(XX$, 25, 2))
'       '
'       Y$ = REGBLAN$("!BOLREDET")
'       Call REPLA(Y$, MKI$(CI%), 83, 2)
'       Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'       Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
'       Call REPLA(Y$, MKS$(NPED&), 151, 4)
'       Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
'       Call REPLA(Y$, MKD$(CAPENE#), 163, 8)
'       Call REPLA(Y$, MKS$(REDE&), 189, 4)
'       Call REGAPP("!BOLREDET", Y$)
'     End If
'   Next U&
'   '
'   VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Despacho Materiales Vendidos - " + TRIM$(RASOCLI$(NCLIE%)))
'   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'     Exit Sub
'   End If
'   '
'   For YX& = 1 To ULTREG&("!BOLREDET")
'     YY$ = REGLEIDO$("!BOLREDET", YX&)
'     CI% = CVI(Mid$(YY$, 83, 2))
'     CAN# = CVD(Mid$(YY$, 163, 8))
'     If CI% > 0 Then
'       If CI% <= NUMAS% Then
'         If CAN# > 0.005 Then
'           GoTo 15
'         End If
'       End If
'     End If
'   Next YX&
'   Call MENSERR(24, "Despacho Cancelado !!!\Faltan Cantidades o Códigos.")
'   Exit Sub
'   '
'15 CONTA% = ULTREG&("!BOLREDET")
'   If CONTA% > 0 Then
'     Call CIERRARCH("MOVISTO")
'     Call BLOQARCH("MOVISTO")
'     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'     '
'     TTXXX$ = "Materiales Vendidos"
'     If REPARX% = 1 Then
'        TTXXX$ = "Reparación"
'        NREP& = 0
'        For KU& = ULTREG&("PRESREPA") To 1 Step -1
'          XYX$ = REGLEIDO$("PRESREPA", KU&)
'          If CVS(Mid$(XYX$, 125, 4)) = NPED& Then
'            NREP& = CVS(Left$(XYX$, 4))
'            GoTo 16
'          End If
'        Next KU&
'     End If
'     '
'16   TTXYZ$ = Space$(76)
'     HOII% = HOY(HOS$)
'     Call REPLA(TTXYZ$, HOS$, 1, 8)
'     Call REPLA(TTXYZ$, "Despacho de " + TTXXX$ + ":", 11, 56)
'     Call REPLA(TTXYZ$, USERTER$, 69, 24)
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + TTXYZ$)
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + Space$(88))
'     If REPARX% = 1 Then
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXYZ$)
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + Space$(88))
'     End If
'     '
'     AINDI% = 0
'     For YX& = 1 To ULTREG&("!BOLREDET")
'       YY$ = REGLEIDO$("!BOLREDET", YX&)
'       CI% = CVI(Mid$(YY$, 83, 2))
'       CAN# = CVD(Mid$(YY$, 163, 8))
'       LOTE$ = Mid$(YY$, 69, 12)
'       If CI% > 0 Then
'         If CI% <= NUMAS% Then
'           If CAN# > 0.005 Then
'             NUOC& = CVS(Mid$(YY$, 151, 4))
'             REDE& = CVS(Mid$(YY$, 189, 4))
'             DOSEC$ = ""
'             CMO$ = "PL"
'             If REPARX% = 1 Then CMO$ = "*D"
'             REMI$ = "PL." + TRIM$(Str$(NPED&))
'             FF% = HOY(HOS$)
'             '
'             CANTI = (-1) * CAN#
'             Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NCLIE%), 0, "$", NC1%, DEPORIG%, CANTI)
'             '
'             ZZ$ = REGLEIDO$("DESPAPED", REDE&)
'             CAENT = CAN#
'             Call REPLA(ZZ$, MKS$(CAENT), 95, 4)
'             Call REPLA(ZZ$, Chr$(2), 27, 1) ' PASA A STATUS ENTREGADO
'             Call GRAREG("DESPAPED", ZZ$, REDE&)
'             '
'             TTXX$ = Space$(76)
'             CIXI% = CI%
'             Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
'             Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
'             Call REPLA(TTXX$, CSTRING$(MKS$(CAENT), 3, 10, 1, 2), 57, 10)
'             Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + TTXX$, 96))
'             If REPARX% = 1 Then
'               Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXX$)
'             End If
'             '
'             CAPEN = CVS(Mid$(ZZ$, 75, 4)) - CAENT
'             If CAPEN > 0.05 Then
'               Call REPLA(ZZ$, MKS$(CAPEN), 75, 4)
'               Call REPLA(ZZ$, MKS$(0), 95, 4)
'               Call REPLA(ZZ$, Chr$(1), 27, 1)
'               Call REGAPP("DESPAPED", ZZ$)
'             End If
'             '
'             IPEDIT$ = Mid$(ZZ$, 113, 10)
'             RFG$ = RECFILT$("PEDIPREP", 1, IPEDIT$)
'             For KU& = 1 To Len(RFG$) Step 4
'               REPRE& = CVS(Mid$(RFG$, KU&, 4))
'               YY$ = REGLEIDO$("PEDIPREP", REPRE&)
'               CAPREP = CVS(Mid$(YY$, 91, 4))
'               If CAPREP <= CAENT + 0.005 Then
'                    Call GRAREG("PEDIPREP", String$(128, 0), REPRE&)
'                    CAENT = CAENT - CAPREP
'                  Else
'                    Call REPLA(YY$, MKS$(CAPREP - CAENT), 91, 4)
'                    Call GRAREG("PEDIPREP", YY$, REPRE&)
'                    CAENT = 0
'               End If
'             Next KU&
'             '
'             JJ& = 0
'             FIN& = ULTREG&("PEDIPREP")
'             For U& = 1 To FIN&
'               ZZ$ = REGLEIDO$("PEDIPREP", U&)
'               If CVS(Mid$(ZZ$, 17, 4)) > 0 Then
'                 JJ& = JJ& + 1
'                 Call GRAREG("PEDIPREP", ZZ$, JJ&)
'               End If
'             Next U&
'             Call SETULTREG("PEDIPREP", JJ&)
'             Call CIERRARCH("PEDIPREP")
'             '
'           End If
'         End If
'       End If
'     Next YX&
'     '
'     Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + Space$(10) + String$(56, "-"), 96))
'     Call REGAPP("ANOTAPED", MKS$(NPED&) + String$(88, "="))
'     If REPARX% = 1 Then
'       Call REGAPP("ANOTAREP", Left$(MKS$(NREP&) + Space$(10) + String$(56, "-"), 96))
'       Call REGAPP("ANOTAREP", MKS$(NREP&) + String$(88, "="))
'     End If
'     '
'     Call CIERRARCH("ANOTAPED")
'     Call CIERRARCH("ANOTAREP")
'     Call CIERRARCH("MOVISTO")
'     Call CIERRARCH("MOVIREPA")
'     Call CIERRARCH("DESPAPED")
'     '
'     Label7 = ""
'     Label7 = TRIM$(Str$(NPED&))
'     '
'     Screen.MousePointer = 1
'     '
'     If TRIM$(FORIPRE$) <> "" Then
'       '
'       FEX = FF%
'       FECHDOC$ = FECHATEX$(FEX)
'       NUMEDOC$ = TRIM$(REMI$)
'       TIPODOC$ = "Despacho de Materiales"
'       DESTIDOC% = NCLIE%
'       '
'       CODPARAM$(1) = "UBI-MPRI": DOCPARAM$(1) = TRIM$(Label14)
'       CAPARDOC% = 1
'       '
'       CODTABU1$(1) = "COD-MAT"
'       CODTABU1$(2) = "DES-MAT"
'       CODTABU1$(3) = "CANTIDAD"
'       CODTABU1$(4) = "CAKILOS"
'       CODTABU1$(5) = "UNIMED"
'       CACOLTAB1% = 5
'       '
'       CAITAB1% = 0
'       '
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         If CI% > 0 Then
'           If CI% <= NUMAS% Then
'             If CAN# > 0.005 Then
'               UNID$ = TRIM$(Mid$(YY$, 85, 4))
'               If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
'               NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'               NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'               '
'               CAITAB1% = CAITAB1% + 1
'               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
'               TETABU1$(3, CAITAB1%) = CANS$
'               TETABU1$(4, CAITAB1%) = KILO$
'               TETABU1$(5, CAITAB1%) = UNID$
'               '
'             End If
'           End If
'         End If
'       Next YX&
'       '
'       CANCELPRINT% = 0
'       If CONTROL$("CDESPACH", "SUPRINT") = "SI" Then CANCELPRINT% = 1
'       Call PRINTDOC("CDESPACH")   ' Comprobante de Despacho
'       '
'     End If
'   End If
'   '
'99 Call CIERRARCH("*.*")
'   '
End Sub

Private Sub Option1_GotFocus()
    '
    Text4.Enabled = False
    '
End Sub

Private Sub Option2_GotFocus()
    '
    Text4.Enabled = False
    '
End Sub

Private Sub Option3_GotFocus()
    '
    Text4.Enabled = True
    '
End Sub

Private Sub Option4_Click()
   Call CALPROCON(7)
End Sub

Private Sub Option5_Click()
   Call CALPROCON(30)
End Sub

Private Sub Option6_Click()
   Call CALPROCON(90)
End Sub

Private Sub Option7_Click()
   Call CALPROCON(180)
End Sub

Private Sub Option8_Click()
   Call CALPROCON(365)
End Sub

Private Sub Text2_Change()
   MODIDES% = 1
End Sub

Private Sub Text3_Change()
   '
   X$ = FILTERGETRECORD$("BADAMODE", 1, Text3)
   '
   Call BLANFORMU
   '
   If ECOARCH("BADAMODE", Text3) <> "" Then
     Text2 = TRIM$(Mid$(X$, 17, 48))
     MODIDES% = 0
     Text7 = TRIM$(Mid$(X$, 65, 32))
     NROMAT% = Val(Mid$(X$, 117, 1))
     Select Case NROMAT%
         Case 1: Option1.Value = True
         Case 2: Option2.Value = True
         Case 3: Option3.Value = True
                 Text4 = TRIM$(Mid$(X$, 119, 16))
     End Select
     FRECUENCIA% = Val(Mid$(X$, 97, 16))
     Label6 = FECHATEX$(CVI(Mid$(X$, 113, 2)))
     Select Case FRECUENCIA%
         Case 1: Option4.Value = True: Call CALPROCON(7)
         Case 2: Option5.Value = True: Call CALPROCON(30)
         Case 3: Option6.Value = True: Call CALPROCON(90)
         Case 4: Option7.Value = True: Call CALPROCON(180)
         Case 5: Option8.Value = True: Call CALPROCON(365)
     End Select
     '
     CI% = CVI(Mid$(X$, 135, 2))
     If CI% > 0 Then
       Label5 = DESCRIT0$(CI%)
       Text5 = CODEXT$(CI%)
       Label3 = UNIMED$(CI%)
     End If
     '
     CANPIEZAS = CVS(Mid$(X$, 137, 4))
     Text8 = FORMATNUM$(CANPIEZAS, "F8.1")
     '
     RFF$ = RECFILT$("ANOTAMOD", 1, Text3)
     For U& = 1 To Len(RFF$) Step 4
       RENOTA& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("ANOTAMOD ", RENOTA&)
       List1.AddItem Mid$(XX$, 17, 92)
     Next U&
     '
     Text6 = ""
     ZZ$ = RECFILT$("OBMODELO", 1, Text3)
     For U& = 1 To Len(ZZ$) Step 4
       RENOTA& = CVS(Mid$(ZZ$, U&, 4))
       XX$ = REGLEIDO$("OBMODELO ", RENOTA&)
       TEXTO$ = TRIM$(Mid$(XX$, 17, 34))
       Text6 = Text6 + TEXTO$ + Space(34 - (Len(TEXTO$)))
     Next U&
  End If
  '
End Sub

Sub BLANFORMU()
    '
    Option1.Value = True
    Text4 = ""
    Text1 = ""
    Text2 = ""
    Text7 = ""
    Option4.Value = True
    Label6 = ""
    Label9 = ""
    Text5 = ""
    Label5 = ""
    Label3 = ""
    Text6 = ""
    List1.Clear
    Text8 = ""
    '
End Sub

Private Sub Text3_DBLCLICK()
    '
    Call Command4_Click
    '
End Sub

Private Sub Text5_Change()
    '
    CI% = CODINT%(Text5)
    If CI% > 0 Then
      Label5 = DESCRIT0$(CI%)
      Label3 = UNIMED$(CI%)
    Else
      Label5 = "": Label3 = ""
    End If
    '
End Sub

Private Sub Text5_DblClick()
    '
    Call Command8_Click
    '
End Sub

Sub CALPROCON(CADIAS%)
    '
    FEBASE = FECHANUM(Label6)
    If FEBASE < 1 Then
      Label9 = ""
      Exit Sub
    End If
    '
    FEPROCO = FEBASE
    While DIENTRE%(FEBASE, FEPROCO) < CADIAS%
      FEPROCO = DIPOST(FEPROCO)
    Wend
    '
    Label9.ForeColor = RGB(0, 0, 0)
    If FEPROCO <= HOY(HOS$) Then
       Label9.ForeColor = RGB(255, 0, 0)
    End If
    '
    Label9 = FECHATEX$(FEPROCO)
    '
End Sub

Sub DESPAMODE()
    '
    If DERACCE%("REMIMODE", "Remitir Modelo a Proveedor") >= 2 Then
      REMIMODE.Text6 = "Modelo: " + TRIM$(Text3) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Text2 + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
      REMIMODE.Text6 = REMIMODE.Text6 + "Cantidad de Piezas: " + TRIM$(Text8) + Chr$(13) + Chr$(10) + "Próximo Control: " + TRIM$(Label9) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
      REMIMODE.Text6 = REMIMODE.Text6 + TRIM$(Text6)
      REMIMODE.Show 1
      '
      CODMODE$ = AJUSTI$(Text3, 16)
      List1.Clear
      RFF$ = RECFILT$("ANOTAMOD", 1, CODMODE$)
      '
      For U& = 1 To Len(RFF$) Step 4
        RENOTA& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("ANOTAMOD", RENOTA&)
        List1.AddItem Mid$(XX$, 17, 92)
      Next U&
      '
    End If
End Sub

Sub RECIMODE()
    '
    If DERACCE%("RECIMODE", "Recibir Modelo de Proveedor") >= 2 Then
      RECEPMODE.Label4 = "Modelo: " + TRIM$(Text3) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Text2
      RECEPMODE.Show 1
      '
      CODMODE$ = AJUSTI$(Text3, 16)
      List1.Clear
      RFF$ = RECFILT$("ANOTAMOD", 1, CODMODE$)
      '
      For U& = 1 To Len(RFF$) Step 4
        RENOTA& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("ANOTAMOD", RENOTA&)
        List1.AddItem Mid$(XX$, 17, 92)
      Next U&
      '
    End If
End Sub

