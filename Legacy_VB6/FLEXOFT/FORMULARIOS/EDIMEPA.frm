VERSION 5.00
Begin VB.Form EDIMEPA 
   Caption         =   "Edición Mensajes de Pantalla"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11880
   Begin VB.CommandButton Command3 
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
      Height          =   765
      Left            =   9765
      Picture         =   "EDIMEPA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Listado de Secuencias de Demostracion"
      Top             =   7455
      Width           =   750
   End
   Begin VB.Frame Frame10 
      Caption         =   "Presentación Preliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   6300
      TabIndex        =   40
      Top             =   4515
      Width           =   5475
      Begin VB.Label Label16 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Información"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   1155
         TabIndex        =   42
         Top             =   1365
         Width           =   3165
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   570
         Left            =   1155
         TabIndex        =   41
         Top             =   1365
         Width           =   3150
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame9 
      Height          =   855
      Left            =   6300
      TabIndex        =   13
      Top             =   7350
      Width           =   3375
      Begin VB.Label Label6 
         Caption         =   "T. Presentación"
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
         TabIndex        =   15
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Despliegue"
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
         TabIndex        =   14
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Demo Unit."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10605
      TabIndex        =   72
      ToolTipText     =   "Generación Código Fuente para Demo Unitaria"
      Top             =   7875
      Width           =   1170
   End
   Begin VB.ListBox LISECUEN 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "EDIMEPA.frx":030A
      Left            =   210
      List            =   "EDIMEPA.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   67
      Top             =   270
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.Frame Frame16 
      Caption         =   "Indice de Secuencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1720
      Left            =   105
      TabIndex        =   55
      Top             =   50
      Width           =   5400
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         TabIndex        =   68
         Top             =   315
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1785
         TabIndex        =   56
         Top             =   660
         Width           =   3480
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Left            =   315
            TabIndex        =   66
            Top             =   260
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Left            =   945
            TabIndex        =   65
            Top             =   260
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Left            =   1575
            TabIndex        =   64
            Top             =   260
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Left            =   2205
            TabIndex        =   63
            Top             =   260
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Left            =   2835
            TabIndex        =   62
            Top             =   260
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Index           =   5
            Left            =   315
            TabIndex        =   61
            Top             =   530
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Index           =   6
            Left            =   945
            TabIndex        =   60
            Top             =   530
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Index           =   7
            Left            =   1575
            TabIndex        =   59
            Top             =   530
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Index           =   8
            Left            =   2205
            TabIndex        =   58
            Top             =   530
            Width           =   540
         End
         Begin VB.OptionButton Option1 
            Caption         =   "0"
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
            Index           =   9
            Left            =   2835
            TabIndex        =   57
            Top             =   530
            Width           =   540
         End
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
         Left            =   1785
         TabIndex        =   69
         Top             =   315
         Width           =   3480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Secuencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10605
      TabIndex        =   54
      Top             =   7455
      Width           =   1170
   End
   Begin VB.PictureBox Picture3 
      Height          =   390
      Left            =   9190
      ScaleHeight     =   330
      ScaleWidth      =   2535
      TabIndex        =   51
      Top             =   105
      Width           =   2600
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2520
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   0
         Left            =   2205
         Picture         =   "EDIMEPA.frx":030E
         Stretch         =   -1  'True
         ToolTipText     =   "Cierra la Aplicación"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   1
         Left            =   1890
         Picture         =   "EDIMEPA.frx":0458
         Stretch         =   -1  'True
         ToolTipText     =   "Guarda la Leyenda"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   2
         Left            =   1575
         Picture         =   "EDIMEPA.frx":0762
         Stretch         =   -1  'True
         ToolTipText     =   "Visualizar Imagen Completa"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   4
         Left            =   1260
         Picture         =   "EDIMEPA.frx":0A6C
         Stretch         =   -1  'True
         ToolTipText     =   "Visualizar Secuencia"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   3
         Left            =   945
         Stretch         =   -1  'True
         ToolTipText     =   "Pausa de Ejecución"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   5
         Left            =   630
         Stretch         =   -1  'True
         ToolTipText     =   "Imprime Pantalla"
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   6
         Left            =   315
         Stretch         =   -1  'True
         ToolTipText     =   "Más Información (Si Disponible)"
         Top             =   0
         Width           =   345
      End
      Begin VB.Line Line2 
         X1              =   2520
         X2              =   2520
         Y1              =   0
         Y2              =   420
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Index           =   7
         Left            =   0
         Stretch         =   -1  'True
         ToolTipText     =   "Solicitud de Instalación"
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Flecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   6300
      TabIndex        =   44
      Top             =   3045
      Width           =   2640
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Caption         =   "Frame12"
         Height          =   1125
         Left            =   105
         TabIndex        =   52
         Top             =   210
         Width           =   855
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   210
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   53
            Top             =   210
            Width           =   330
            Begin VB.Line Line3 
               BorderWidth     =   2
               Index           =   4
               X1              =   10
               X2              =   10
               Y1              =   0
               Y2              =   315
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               Index           =   5
               X1              =   0
               X2              =   315
               Y1              =   315
               Y2              =   315
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               Index           =   6
               X1              =   315
               X2              =   315
               Y1              =   0
               Y2              =   315
            End
            Begin VB.Line Line3 
               BorderWidth     =   2
               Index           =   7
               X1              =   0
               X2              =   315
               Y1              =   10
               Y2              =   10
            End
         End
         Begin VB.Label Label10 
            Caption         =   "3"
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
            Left            =   0
            TabIndex        =   71
            Top             =   735
            Width           =   225
         End
         Begin VB.Line Line5 
            BorderWidth     =   3
            Index           =   2
            X1              =   630
            X2              =   750
            Y1              =   945
            Y2              =   735
         End
         Begin VB.Line Line5 
            BorderWidth     =   4
            Index           =   1
            X1              =   525
            X2              =   735
            Y1              =   740
            Y2              =   740
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            Index           =   0
            X1              =   315
            X2              =   735
            Y1              =   1000
            Y2              =   735
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   1
            Visible         =   0   'False
            X1              =   525
            X2              =   675
            Y1              =   370
            Y2              =   370
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   2
            Visible         =   0   'False
            X1              =   370
            X2              =   370
            Y1              =   675
            Y2              =   525
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   3
            Visible         =   0   'False
            X1              =   60
            X2              =   210
            Y1              =   370
            Y2              =   370
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            Index           =   4
            Visible         =   0   'False
            X1              =   370
            X2              =   370
            Y1              =   210
            Y2              =   60
         End
      End
      Begin VB.TextBox Text8 
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
         Left            =   1890
         TabIndex        =   48
         Text            =   "6000"
         Top             =   630
         Width           =   645
      End
      Begin VB.TextBox Text7 
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
         Left            =   1890
         TabIndex        =   47
         Text            =   "4350"
         Top             =   945
         Width           =   645
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   960
         LargeChange     =   250
         Left            =   1050
         Max             =   8700
         SmallChange     =   10
         TabIndex        =   46
         Top             =   315
         Value           =   4350
         Width           =   180
      End
      Begin VB.HScrollBar HScroll5 
         Height          =   180
         LargeChange     =   250
         Left            =   1260
         Max             =   12000
         SmallChange     =   10
         TabIndex        =   45
         Top             =   315
         Value           =   6000
         Width           =   1275
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   420
         X2              =   420
         Y1              =   630
         Y2              =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Xfin ="
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
         Left            =   1050
         TabIndex        =   50
         Top             =   735
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Yfin ="
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
         Left            =   1155
         TabIndex        =   49
         Top             =   1050
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   645
      Left            =   2520
      ScaleHeight     =   585
      ScaleWidth      =   1005
      TabIndex        =   39
      Top             =   5145
      Width           =   1065
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Label3"
         Height          =   120
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Titulo de Leyenda"
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
      Left            =   5670
      TabIndex        =   35
      Top             =   1785
      Width           =   2430
      Begin VB.TextBox Text6 
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
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "EDIMEPA.frx":0D76
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alinear"
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
      Left            =   10605
      TabIndex        =   31
      Top             =   630
      Width           =   1170
      Begin VB.OptionButton Option11 
         Caption         =   "Izq."
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
         TabIndex        =   34
         Top             =   315
         Width           =   750
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Der"
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
         Left            =   210
         TabIndex        =   33
         Top             =   735
         Width           =   750
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Cent"
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
         Left            =   210
         TabIndex        =   32
         Top             =   525
         Value           =   -1  'True
         Width           =   750
      End
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
      Height          =   1695
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   1890
      Width           =   5370
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFF80&
      Height          =   4185
      Left            =   105
      ScaleHeight     =   4125
      ScaleWidth      =   5940
      TabIndex        =   23
      Top             =   3990
      Width           =   6000
      Begin VB.Image Image1 
         Height          =   4500
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6000
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Cierra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   10920
      TabIndex        =   16
      Top             =   3045
      Width           =   855
      Begin VB.OptionButton Option5 
         Caption         =   "No"
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
         TabIndex        =   18
         Top             =   630
         Width           =   570
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Si"
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
         TabIndex        =   17
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Pos.Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   9030
      TabIndex        =   7
      Top             =   3045
      Width           =   1800
      Begin VB.HScrollBar HScroll4 
         Height          =   180
         LargeChange     =   250
         Left            =   315
         Max             =   12000
         SmallChange     =   10
         TabIndex        =   38
         Top             =   315
         Value           =   6000
         Width           =   1380
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   960
         LargeChange     =   250
         Left            =   105
         Max             =   8700
         SmallChange     =   10
         TabIndex        =   37
         Top             =   315
         Value           =   4350
         Width           =   180
      End
      Begin VB.TextBox Text2 
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
         Left            =   1050
         TabIndex        =   9
         Text            =   "4350"
         Top             =   945
         Width           =   645
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
         Left            =   1050
         TabIndex        =   8
         Text            =   "6000"
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Yfin ="
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
         TabIndex        =   11
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Xfin ="
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
         TabIndex        =   10
         Top             =   735
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Colores"
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
      Left            =   8190
      TabIndex        =   6
      Top             =   1785
      Width           =   3585
      Begin VB.OptionButton Option2 
         Caption         =   "Flecha"
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
         Left            =   315
         TabIndex        =   30
         Top             =   785
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Fondo"
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
         Left            =   315
         TabIndex        =   29
         Top             =   550
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Letra"
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
         Left            =   315
         TabIndex        =   28
         Top             =   315
         Width           =   960
      End
      Begin VB.PictureBox Picture1 
         Height          =   645
         Left            =   1365
         ScaleHeight     =   585
         ScaleWidth      =   270
         TabIndex        =   22
         Top             =   315
         Width           =   330
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   180
         LargeChange     =   16
         Left            =   1890
         Max             =   255
         TabIndex        =   21
         Top             =   800
         Value           =   150
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   180
         LargeChange     =   16
         Left            =   1890
         Max             =   255
         TabIndex        =   20
         Top             =   560
         Value           =   220
         Width           =   1275
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   180
         LargeChange     =   16
         Left            =   1890
         Max             =   255
         TabIndex        =   19
         Top             =   315
         Value           =   240
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   2
         Left            =   3255
         Top             =   795
         Width           =   180
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   1
         Left            =   3255
         Top             =   555
         Width           =   180
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   180
         Index           =   0
         Left            =   3255
         Top             =   315
         Width           =   180
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Efectos"
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
      Left            =   9240
      TabIndex        =   2
      Top             =   630
      Width           =   1275
      Begin VB.CheckBox Check3 
         Caption         =   "Subray"
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
         TabIndex        =   5
         Top             =   735
         Width           =   960
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italica"
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
         TabIndex        =   4
         Top             =   525
         Width           =   960
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Negrita"
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
         TabIndex        =   3
         Top             =   315
         Value           =   1  'Checked
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tamaño"
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
      Left            =   8190
      TabIndex        =   1
      Top             =   630
      Width           =   960
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "8"
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Fuente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   5670
      TabIndex        =   0
      Top             =   50
      Width           =   2430
      Begin VB.ListBox FONTLIST 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   660
         Width           =   2220
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
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
         TabIndex        =   25
         Top             =   315
         Width           =   2220
      End
   End
   Begin VB.Label NOSECACT 
      Height          =   330
      Left            =   2625
      TabIndex        =   74
      Top             =   630
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label9 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   105
      TabIndex        =   70
      Top             =   3730
      Width           =   6000
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   6300
      TabIndex        =   27
      Top             =   2730
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "EDIMEPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COLETRA, COFONDO, COFLECHA
Dim ROJO(2), VERDE(2), AZUL(2)
Dim CONTROCOLO%
Dim ORIFLECHA%
Dim REDUC
Dim ATRIBU$
Dim ORDEGUARDA%
Dim MUEVELE%, MXORI%, MYORI%
'
Private Sub Check1_Click()
   Call AJUSTANCHO
End Sub

Private Sub Check2_Click()
   Call AJUSTANCHO
End Sub

Private Sub Check3_Click()
   Call AJUSTANCHO
End Sub

Private Sub Command1_Click()
   For KKU& = 1 To SECUDEMO.LISECUEN.ListCount
     If Trim$(Left$(SECUDEMO.LISECUEN.List(KKU& - 1), 8)) = Trim$(NOSECACT) Then
       SECUDEMO.LISECUEN.ListIndex = KKU& - 1
       Call SECUDEMO.LISECUEN_Click
       Exit For
     End If
   Next KKU&
   '
   SECUDEMO.Show 1
End Sub

Private Sub Command2_Click()
   Dim ARIMAX$(128)
   '
   Command2.Enabled = False
   LISECUDE.Show 1
   Command2.Enabled = True
   TTXX$ = Trim$(LISECUDE.List1.Text)
   If TTXX$ <> "" Then
     '
     SECUDE$ = Trim$(Left$(TTXX$, 10))
     ASECU$ = SECUDE$ + ".SEC"
     LOSECU& = 0
     On Error Resume Next
     LOSECU& = FileLen(DEMOPATH$ + "\SCRIPTS\" + ASECU$)
     On Error GoTo 0
     If LOSECU& < 128 Then
       MsgBox "Secuencia '" + ASECU$ + "' no Instalada."
       Exit Sub
     End If
     '
     Call CARLISECUTO(SECUDE$)
     If SECUDE$ = "" Then GoTo 99
     '
     CARPESAL$ = DEMOPATH$ + "\UNIDEMO\" + SECUDE$
     On Error Resume Next
     MkDir CARPESAL$
     On Error GoTo 0
     '
4    OPUTFILE$ = CARPESAL$ + "\SECUDEMO.BAS"
5    OPUTFILE1$ = InputBox$("Ruta\Archivo", "Localización de Archivo Módulo Fuente", OPUTFILE$)
     If OPUTFILE1$ = "" Then GoTo 99
     If OPUTFILE1$ <> OPUTFILE$ Then
       MsgBox "No se Puede cambiar Archivo de Salida"
       GoTo 4
     End If
     '
     Screen.MousePointer = 11
     LOUTPUT& = 0
     On Error Resume Next
     LOUTPUT& = FileLen(OPUTFILE$)
     On Error GoTo 0
     If LOUTPUT& > 0 Then
       If MsgBox("Archivo Existente !!! - Desea Sobre-Escribirlo ?", 259) <> 6 Then GoTo 5
     End If
     '
     NX% = FreeFile
     Open OPUTFILE$ For Output As #NX%
     Print #NX%, "Public SECUDE$(256)"
     Print #NX%, "PUBLIC CASECU%"
     Print #NX%, "PUBLIC IMANAME$(64)"
     Print #NX%, "PUBLIC CAIMAG%"
     Print #NX%, "PUBLIC LEYENDX$(512)"
     Print #NX%, "PUBLIC CALEYEN%"
     Print #NX%, "'"
     Print #NX%, "SUB CARSECU()"
     Print #NX%, "   '"
     '
     CASECU% = 0: CAIMA% = 0: CALEYEN% = 0
     For UI& = 1 To SECUENDE.List1.ListCount
       TTXX$ = RTrim$(SECUENDE.List1.List(UI& - 1))
       If TTXX$ <> "" Then
         CASECU% = CASECU% + 1
         TTYY$ = "   SECUDE$(" + Trim$(Str$(CASECU%)) + ") = " + Chr$(34) + RTrim$(TTXX$) + Chr$(34)
         Print #NX%, TTYY$
         '
         ARIMA$ = Trim$(Mid$(TTXX$, 91, 16))
         If InStr(ARIMA$, ".TXT") < 1 Then
           For KKI% = 1 To CAIMA%
             If ARIMAX$(KKI% - 1) = ARIMA$ Then GoTo 19
           Next KKI%
           CAIMA% = CAIMA% + 1
           ARIMAX$(CAIMA% - 1) = ARIMA$
         End If
         '
       End If
19   Next UI&
     '
     Print #NX%, "   '"
     '
     For KKI% = 1 To CAIMA%
       IMAFI$ = Trim$(Str$(KKI% - 1))
       While Len(IMAFI$) < 3: IMAFI$ = "0" + IMAFI$: Wend
       IMAFI$ = "IMAGE" + IMAFI$
       FileCopy DEMOPATH$ + "\SCREENS\" + ARIMAX$(KKI% - 1), CARPESAL$ + "\" + IMAFI$ + ".JPG"
       '
       TTYY$ = "   IMANAME$(" + Trim$(Str$(KKI% - 1)) + ") = " + Chr$(34) + RTrim$(ARIMAX$(KKI% - 1)) + Chr$(34)
       Print #NX%, TTYY$
     Next KKI%
     '
     Print #NX%, "   '"
     '
     For UI& = 1 To SECUENDE.List1.ListCount
       TTXX$ = RTrim$(SECUENDE.List1.List(UI& - 1))
       If TTXX$ <> "" Then
         AGUARDA$ = Trim$(UCase$(Left$(TTXX$, 10)))
         PPXX% = InStr(AGUARDA$, ".")
         If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
         AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
         '
         For ORLEYE% = 0 To 9
10         ALEYEN$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
           LOARGUA& = 0
           On Error Resume Next
           LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + ALEYEN$)
           On Error GoTo 0
           If LOARGUA& > 0 Then
             ORDEGUARDA% = ORLEYE%
             '
             N1% = FreeFile
             Open DEMOPATH$ + "\LABELS\" + ALEYEN$ For Binary Access Read Shared As #N1%
             TEGRABA$ = Input(LOARGUA&, #N1%)
             Close #N1%
             '
15           PPXY% = InStr(TEGRABA$, Chr$(13) + Chr$(10))
             If PPXY% > 0 Then
               TEGRABA$ = Left$(TEGRABA$, PPXY% - 1) + Chr$(125) + Mid$(TEGRABA$, PPXY% + 2)
               GoTo 15
             End If
             '
16           PPXY% = InStr(TEGRABA$, Chr$(34))
             If PPXY% > 0 Then
               TEGRABA$ = Left$(TEGRABA$, PPXY% - 1) + Chr$(126) + Mid$(TEGRABA$, PPXY% + 1)
               GoTo 15
             End If
             '
             TTXX$ = Left$(ALEYEN$ + Space$(16), 16) + TEGRABA$
             CALEYEN% = CALEYEN% + 1
             TTYY$ = "   LEYENDX$(" + Trim$(Str$(CALEYEN%)) + ") = " + Chr$(34) + RTrim$(TTXX$) + Chr$(34)
             Print #NX%, TTYY$
             '
           End If
         Next ORLEYE%
       End If
     Next UI&
     '
     Print #NX%, "   '"
     Print #NX%, "   CASECU% = " + Trim$(Str$(CASECU%))
     Print #NX%, "   '"
     Print #NX%, "   CAIMAG% = " + Trim$(Str$(CAIMA%))
     Print #NX%, "   '"
     Print #NX%, "   CALEYEN% = " + Trim$(Str$(CALEYEN%))
     Print #NX%, "   '"
     Print #NX%, "END SUB"
     Close #NX%
     Close #N1%
     '
     Screen.MousePointer = 1
     MsgBox "Archivos Creados en " + CARPESAL$
     '
   End If
99 End Sub

Private Sub Command3_Click()
   '
   Dim ARIMAX$(128)
   Dim RELE$(32)
   '
   Command2.Enabled = False
5  LISECUDE.Show 1
   Command2.Enabled = True
   TTXX$ = Trim$(LISECUDE.List1.Text)
   OPENFORMS = DoEvents
   If TTXX$ <> "" Then
     '
     Screen.MousePointer = 11
     SECUDE$ = Trim$(Left$(TTXX$, 10))
     ASECU$ = SECUDE$ + ".SEC"
     LSECU$ = SECUDE$ + ".TXT"
     TITUSE$ = Trim$(Mid$(TTXX$, 11))
     LOSECU& = 0
     On Error Resume Next
     LOSECU& = FileLen(DEMOPATH$ + "\SCRIPTS\" + ASECU$)
     On Error GoTo 0
     If LOSECU& < 128 Then
       MsgBox "Secuencia '" + ASECU$ + "' no Instalada."
       GoTo 5
     End If
     '
     Call CARLISECUTO(SECUDE$)
     '
     On Error Resume Next
     MkDir "C:\TESECUDE"
     On Error GoTo 0
     '
     NX% = FreeFile
     Open "C:\TESECUDE\" + LSECU$ For Output As #NX%
     '
     Print #NX%, "Secuencia: " + Trim$(SECUDE$) + " - " + Trim$(TITUSE$)
     Print #NX%, "**********"
     Print #NX%, " "
     Print #NX%, String$(128, "=")
     '
     CASECU% = 0
     For UI& = 1 To SECUENDE.List1.ListCount
       TTXX$ = RTrim$(SECUENDE.List1.List(UI& - 1))
       If TTXX$ <> "" Then
         CASECU% = CASECU% + 1
         TEGRABA$ = "": CAPII% = 0
         CODISE$ = Left$(Trim$(Left$(TTXX$, 8)) + Space$(8), 8)
         ARIMA$ = Left$(Trim$(Mid$(TTXX$, 91, 16)) + Space$(16), 16)
         TITUSE$ = Trim$(Mid$(TTXX$, 11, 80))
         If InStr(ARIMA$, ".TXT") > 0 Then
           TEGRABA$ = TITUSE$
8          For KKI% = 1 To Len(TEGRABA$)
             If Mid$(TEGRABA$, KKI%, 1) = "\" Then
               TEGRABA$ = Left$(TEGRABA$, KKI% - 1) + Chr$(13) + Chr$(10) + Mid$(TEGRABA$, KKI% + 1)
               GoTo 8
             End If
           Next KKI%
           TITUSE$ = "Nuevo Capitulo"
           CAPII% = 1
         End If
         '
         CASECX$ = Trim$(Str$(CASECU%))
         While Len(CASECX$) < 3: CASECX$ = " " + CASECX$: Wend
         Print #NX%, CASECX$ + " - Codigo: " + CODISE$ + "    Imagen: " + ARIMA$ + "  Titulo: " + TITUSE$
         Print #NX%, "      *******             *******                   *******"
         '
         If CAPII% = 1 Then GoTo 12
         '
         AGUARDA$ = Trim$(UCase$(CODISE$))
         PPXX% = InStr(AGUARDA$, ".")
         If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
         AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
         '
         For ORLEYE% = 0 To 9
10         ALEYEN$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
           LOARGUA& = 0
           On Error Resume Next
           LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + ALEYEN$)
           On Error GoTo 0
           If LOARGUA& > 0 Then
             ORDEGUARDA% = ORLEYE%
             '
             N1% = FreeFile
             Open DEMOPATH$ + "\LABELS\" + ALEYEN$ For Binary Access Read Shared As #N1%
             TEGRABA$ = Input(LOARGUA&, #N1%)
             Close #N1%
             '
12           PPXX% = InStr(TEGRABA$, "|")
             If PPXX% < 1 Then PPXX% = 1 + Len(TEGRABA$)
             TELEYEN$ = Left$(TEGRABA$, PPXX% - 1) + Chr$(13) + Chr$(10)
             '
             CARELE% = 0
15           PPXY% = InStr(TELEYEN$, Chr$(13) + Chr$(10))
             If PPXY% > 0 Then
               CARELE% = CARELE% + 1
               RELE$(CARELE%) = Left$(TELEYEN$, PPXY% - 1)
               TELEYEN$ = Mid$(TELEYEN$, PPXY% + 2)
               While Len(RELE$(CARELE%)) < 48
                 RELE$(CARELE%) = " " + RELE$(CARELE%) + " "
               Wend
               If Len(TELEYEN$) > 0 Then GoTo 15
             End If
             '
             Print #NX%, String$(128, "-")
             For KKI% = 1 To CARELE%
               Print #NX%, "      " + RELE$(KKI%)
             Next KKI%
             '
             If CAPII% = 1 Then GoTo 18
             '
           End If
         Next ORLEYE%
18       Print #NX%, String$(128, "=")
       End If
19   Next UI&
     '
     Close #NX%
     Screen.MousePointer = 1
     GoTo 5
     '
   End If

End Sub

Private Sub HScroll5_Change()
   If MUEVELE% = 0 Then
     Text8.Text = HScroll5.Value
   End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Text8.Text = Int(X / REDUC + 0.5)
   Text7.Text = Int(Y / REDUC + 0.5)
End Sub

Private Sub Label10_Click()
   Label10.Caption = Trim$(Str$(Val(Label10.Caption) Mod 3 + 1))
End Sub

Private Sub LISECUEN_DblClick()
   '
   LISECUEN.Clear
   N1% = FreeFile
   Dim BUFLE As String * 128
   Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Write Shared As #N1% Len = 128
   For KKI% = 1 To LOF(N1%) / 128
     Get #N1%, KKI%, BUFLE$
     If Trim$(BUFLE$) <> "" Then
       LISECUEN.AddItem Trim$(BUFLE$)
     End If
   Next KKI%
   Close #N1%
   '
   Option1(0).Value = True
   LISECUEN.Visible = True

End Sub

Private Sub List1_Click()
   '
   TTXX$ = List1.Text
   ADEMO$ = Trim$(UCase$(Mid$(TTXX$, 91, 16)))
   TITUPAN$ = Trim$(Mid$(TTXX$, 11, 80))
   If TITUPAN$ = "" Then
     TITUPAN$ = "Pantalla " + ADEMO$
   End If
   '
   If InStr(ADEMO$, ".SEC") > 0 Then
     TTYY1$ = "Imposible Mostrar Archivo de Secuencia"
     TTYY2$ = "'" + Trim$(TITUPAN$) + "'"
     While Len(TTYY1$) < Len(TTYY2$): TTYY1$ = " " + TTYY1$ + " ": Wend
     While Len(TTYY2$) < Len(TTYY1$): TTYY2$ = " " + TTYY2$ + " ": Wend
     MsgBox TTYY1$ + Chr$(10) + Chr$(13) + TTYY2$
     Exit Sub
   End If
   '
   Image1.Stretch = False
   Image1.Picture = LoadPicture("")
   On Error GoTo 90
   Image1.Picture = LoadPicture(DEMOPATH$ + "\SCREENS\" + ADEMO$)
   On Error GoTo 0
   REDUV = Picture4.ScaleHeight / Image1.Height
   REDUH = Picture4.ScaleWidth / Image1.Width
   REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
   Image1.Stretch = True
   Image1.Width = Image1.Width * REDUC
   Image1.Height = Image1.Height * REDUC
   Image1.Top = (Picture4.ScaleHeight - Image1.Height) / 2
   Image1.Left = (Picture4.ScaleWidth - Image1.Width) / 2
   '
   Label8.Caption = TITUPAN$
   Label9.Caption = TITUPAN$
   '
   Option1(0).Value = True
   Call CARGALEYEN
   '
90 On Error GoTo 0
   '
End Sub

Private Sub FONTLIST_Click()
   Label14.Caption = FONTLIST.Text
End Sub

Private Sub Form_Load()
   '
   CULOG$ = Left$(App.Path, 2)
   If InStr(UCase$(App.Path), "ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\") > 0 Then
     CULOG$ = "F:"
   End If
   DEMOPATH$ = CULOG$ + "\DEMOFLEX"
   '
   Call CENTRAFORM(Me)
   '
   For KKI% = 0 To 9
     Option1(KKI%).Caption = Trim$(Str$(KKI%))
   Next KKI%
   '
   FONTLIST.Clear
   For KKI% = 0 To Screen.FontCount - 1  ' Determina el número de fuentes.
      FONTLIST.AddItem Screen.Fonts(KKI%)   ' Las mueve al cuadro de lista.
   Next KKI%
   '
   ROJO(0) = 0: VERDE(0) = 0: AZUL(0) = 0
   ROJO(1) = 240: VERDE(1) = 220: AZUL(1) = 150
   ROJO(2) = 255: VERDE(2) = 255: AZUL(2) = 0
   '
   N1% = FreeFile
   Dim BUFLE As String * 128
   Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Write Shared As #N1% Len = 128
   For KKI% = 1 To LOF(N1%) / 128
     Get #N1%, KKI%, BUFLE$
     If Trim$(BUFLE$) <> "" Then
       LISECUEN.AddItem Trim$(BUFLE$)
     End If
   Next KKI%
   Close #N1%
   '
   Option1(0).Value = True
   LISECUEN.Visible = True
   ' List1.ListIndex = 0
   'Text5.Text = ""
   Call AJUSTANCHO
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Close: End
End Sub

Private Sub Frame16_DblClick()
   LISECUEN.Visible = True
End Sub

Private Sub HScroll1_Change()
   If CONTROCOLO% = 0 Then Call AJUCOLOR
End Sub

Private Sub HScroll2_Change()
   If CONTROCOLO = 0 Then Call AJUCOLOR
End Sub

Private Sub HScroll3_Change()
   If CONTROCOLO% = 0 Then Call AJUCOLOR
End Sub

Private Sub HScroll4_Change()
   If MUEVELE% = 0 Then
     Text1.Text = HScroll4.Value
   End If
End Sub

Private Sub Image2_Click(Index As Integer)
   '
   Image2(Index).Visible = False
   TINIX = Timer
   If Index% = 0 Then
      Close: End
    ElseIf Index% = 1 Then
      Call GUARDALEYEN
    ElseIf Index = 2 Then
      Call GUARDALEYEN
      Call MUESTRAPANDEMO
    ElseIf Index = 3 Then
      '
    ElseIf Index = 4 Then
      '
    ElseIf Index = 5 Then
      '
   End If
   '
99 While Timer < TINIX + 0.2
     DoEvents
   Wend
   Image2(Index).Visible = True
   '
End Sub

Private Sub Label14_Change()
   Call AJUSTANCHO
End Sub

Private Sub LISECUEN_Click()
   '
   ILISTA& = LISECUEN.ListIndex
   If ILISTA& >= 0 Then
     TTXX$ = Trim$(Mid$(LISECUEN.List(ILISTA&), 10) + " (" + Trim$(Left$(LISECUEN.List(ILISTA&), 8))) + ")"
     SECU$ = Trim$(Left$(LISECUEN.List(ILISTA&), 8))
     NOSECACT = SECU$
     Frame16.Caption = TTXX$
     LISECUEN.Visible = False
     '
     Call CARLISECUEN(SECU$)
     On Error Resume Next
     List1.ListIndex = 0
     On Error GoTo 0
     '
   End If
   '
End Sub

Private Sub Option1_Click(Index As Integer)
   Call CARGALEYEN
End Sub

Private Sub Option11_Click(Index As Integer)
   Call AJUSTANCHO
End Sub

Private Sub Option2_Click(Index As Integer)
   CONTROCOLO% = 1
   HScroll1.Value = ROJO(Index)
   HScroll2.Value = VERDE(Index)
   HScroll3.Value = AZUL(Index)
   CONTROCOLO% = 0
   Call AJUCOLOR
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If MUEVELE% = 0 Then
     If Button = 1 Then
       MUEVELE% = 1
       MXORI% = X
       MYORI% = Y
     End If
   End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If MUEVELE% = 1 Then
      Picture2.Left = Picture2.Left + X - MXORI%
      Picture2.Top = Picture2.Top + Y - MYORI%
   End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Text1.Text = Int((Picture2.Left - Picture4.Left - Image1.Left) / REDUC + 0.5)
   Text2.Text = Int((Picture2.Top - Picture4.Top - Image1.Top) / REDUC + 0.5)
   MUEVELE% = 0
End Sub

Private Sub Picture5_Click()
   ORIFLECHA% = (ORIFLECHA% + 1) Mod 5
   Call AJUSTANCHO
End Sub

Private Sub Text1_Change()
   On Error Resume Next
   HScroll4.Value = Text1.Text
   On Error GoTo 0
   If MUEVELE% = 0 Then Call AJUSTANCHO
End Sub

Private Sub Text2_Change()
   On Error Resume Next
   VScroll1.Value = Text2.Text
   On Error GoTo 0
   If MUEVELE% = 0 Then Call AJUSTANCHO
End Sub

Private Sub Text3_Change()
   Call AJUSTANCHO
End Sub

Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text5_Change()
   Call AJUSTANCHO
End Sub

Sub AJUSTANCHO()
   '
   If MUEVELE% = 1 Then
     Picture2.Top = Picture4.Top + Image1.Top + REDUC * Val(Text2.Text)
     Picture2.Left = Picture4.Left + Image1.Left + REDUC * Val(Text1.Text)
     Exit Sub
   End If
   '
   TEPRUEBA.Font = Label14.Caption
   Label15.Font = Label14.Caption
   '
   TAFUENTE = Val(Text3)
   If TAFUENTE < 2 Then
     TAFUENTE = 8
   End If
   '
   TEPRUEBA.FontSize = TAFUENTE
   Label15.FontSize = TAFUENTE
   '
   If Check1.Value = 1 Then
      TEPRUEBA.FontBold = True
      Label15.FontBold = True
     Else
      TEPRUEBA.FontBold = False
      Label15.FontBold = False
   End If
   '
   If Check2.Value = 1 Then
      TEPRUEBA.FontItalic = True
      Label15.FontItalic = True
     Else
      TEPRUEBA.FontItalic = False
      Label15.FontItalic = False
   End If
   '
   If Check3.Value = 1 Then
      TEPRUEBA.FontUnderline = True
      Label15.FontUnderline = True
     Else
      TEPRUEBA.FontUnderline = False
      Label15.FontUnderline = False
   End If
   '
   For KKI% = 0 To 2
     If Option11(KKI%).Value = True Then
       Label15.Alignment = KKI%
     End If
   Next KKI%
   '
   AMAXI = 0
   TTXX$ = Text5.Text
10 PPXX% = InStr(TTXX$, Chr$(10) + Chr$(13))
   If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
   TEPRU$ = Left$(TTXX$, PPXX% - 1)
   TTXX$ = Mid$(TTXX$, PPXX% + 2)
   TEPRUEBA.Caption = TEPRU$
   ATEPRU = TEPRUEBA.Width
   If ATEPRU > AMAXI Then AMAXI = ATEPRU
   If Len(Trim$(TTXX$)) > 0 Then GoTo 10
   '
   Label15.Width = AMAXI
   Label15.Left = (Frame10.Width - Label15.Width) / 2
   Label16.Width = Label15.Width + 10
   Label16.Left = Label15.Left
   Label15.Caption = Text5.Text
   '
   Label16.Top = 80 + (Frame10.Height - Label15.Height - Label16.Height) / 2
   Label15.Top = Label16.Top + Label16.Height - 40
   '
   COLETRA = RGB(ROJO(0), VERDE(0), AZUL(0))
   COFONDO = RGB(ROJO(1), VERDE(1), AZUL(1))
   COFLECHA = RGB(ROJO(2), VERDE(2), AZUL(2))
   Label15.ForeColor = COLETRA
   Label15.BackColor = COFONDO
   For KKI% = 0 To 2: Line5(KKI%).BorderColor = COFLECHA: Next KKI%
   '
   Picture2.BackColor = COFONDO
   Picture2.Width = (Label15.Width + 210) * REDUC
   Picture2.Height = (Label15.Height + Label16.Height) * REDUC
   Picture2.Top = Picture4.Top + Image1.Top + REDUC * Val(Text2.Text)
   Picture2.Left = Picture4.Left + Image1.Left + REDUC * Val(Text1.Text)
   '
   Label3.Height = Label16.Height * REDUC
   Label3.Width = Picture4.Width + 100
   '
   If Trim$(Text5.Text) <> "" Then
        Picture2.Visible = True
      Else
        Picture2.Visible = False
   End If
   '
   For KKI% = 1 To 4
     Line4(KKI%).Visible = False
     If KKI% = ORIFLECHA% Then Line4(KKI%).Visible = True
   Next KKI%
   '
End Sub

Sub AJUCOLOR()
   '
   IOBJETO% = 1
   For KKI% = 0 To 2
     If Option2(KKI%) = True Then IOBJETO% = KKI%
   Next KKI%
   '
   ROJO(IOBJETO%) = Val(HScroll1)
   VERDE(IOBJETO%) = Val(HScroll2)
   AZUL(IOBJETO%) = Val(HScroll3)
   '
   CCOLOR = RGB(ROJO(IOBJETO%), VERDE(IOBJETO%), AZUL(IOBJETO%))
   Picture1.BackColor = CCOLOR
   '
   Call AJUSTANCHO
   '
End Sub

Private Sub Text6_Change()
   Label16.Caption = Text6
End Sub

Private Sub Text7_Change()
   On Error Resume Next
   VScroll2.Value = Text7.Text
   On Error GoTo 0
End Sub

Private Sub Text8_Change()
   On Error Resume Next
   HScroll5.Value = Text8.Text
   On Error GoTo 0
End Sub

Private Sub VScroll1_Change()
   If MUEVELE% = 0 Then
     Text2.Text = VScroll1.Value
   End If
End Sub

Sub GUARDALEYEN()
   '
   Screen.MousePointer = 11
   '
   AGUARDA$ = UCase$(Trim$(Left$(List1.Text, 8)))
   PPXX% = InStr(AGUARDA$, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
   AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
   ORLEYE% = 0
   For KKI% = 0 To 9
     If Option1(KKI%).Value = True Then
       ORLEYE% = KKI%
     End If
   Next KKI%
   AGUARDA$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
   '
   TELEYEN$ = Trim$(Text5.Text)
   '
   ATRIBU$ = ""
   TITULEYE$ = "": For KKI% = 1 To Len(Text6.Text)
                     CARA% = Asc(Mid$(Text6.Text, KKI%, 1))
                     If CARA% >= 32 Then TITULEYE$ = TITULEYE$ + Chr$(CARA%)
                   Next KKI%
   ATRIBU$ = ATRIBU$ + "TITULE=" + Trim$(TITULEYE$) + "\"
   ATRIBU$ = ATRIBU$ + "XFINAL=" + Trim$(Text1.Text) + "\"
   ATRIBU$ = ATRIBU$ + "YFINAL=" + Trim$(Text2.Text) + "\"
   ATRIBU$ = ATRIBU$ + "FONAME=" + Trim$(Label14.Caption) + "\"
   ATRIBU$ = ATRIBU$ + "FOSIZE=" + Trim$(Text3.Text) + "\"
   ATRIBU$ = ATRIBU$ + "FOBOLD=" + Trim$(Check1.Value) + "\"
   ATRIBU$ = ATRIBU$ + "FOITAL=" + Trim$(Check2.Value) + "\"
   ATRIBU$ = ATRIBU$ + "FUNDER=" + Trim$(Check3.Value) + "\"
   '
   ATRIBU$ = ATRIBU$ + "LEROJO=" + Trim$(Str$(ROJO(0))) + "\"
   ATRIBU$ = ATRIBU$ + "LEVERD=" + Trim$(Str$(VERDE(0))) + "\"
   ATRIBU$ = ATRIBU$ + "LEAZUL=" + Trim$(Str$(AZUL(0))) + "\"
   '
   ATRIBU$ = ATRIBU$ + "FOROJO=" + Trim$(Str$(ROJO(1))) + "\"
   ATRIBU$ = ATRIBU$ + "FOVERD=" + Trim$(Str$(VERDE(1))) + "\"
   ATRIBU$ = ATRIBU$ + "FOAZUL=" + Trim$(Str$(AZUL(1))) + "\"
   '
   ATRIBU$ = ATRIBU$ + "PUROJO=" + Trim$(Str$(ROJO(2))) + "\"
   ATRIBU$ = ATRIBU$ + "PUVERD=" + Trim$(Str$(VERDE(2))) + "\"
   ATRIBU$ = ATRIBU$ + "PUAZUL=" + Trim$(Str$(AZUL(2))) + "\"
   '
   ALINE% = 0
   For KKI% = 0 To 2
     If Option11(KKI%).Value = True Then ALINE% = KKI%
   Next KKI%
   ATRIBU$ = ATRIBU$ + "ALINEA=" + Trim$(Str$(ALINE%)) + "\"
   '
   ATRIBU$ = ATRIBU$ + "ORIFLE=" + Trim$(Str$(ORIFLECHA%)) + "\"
   ATRIBU$ = ATRIBU$ + "TIPFLE=" + Trim$(Label10.Caption) + "\"
   ATRIBU$ = ATRIBU$ + "XFLECH=" + Trim$(Text8.Text) + "\"
   ATRIBU$ = ATRIBU$ + "YFLECH=" + Trim$(Text7.Text) + "\"
   '
   TEGRABA$ = TELEYEN$ + "|" + ATRIBU$
   '
   N1% = FreeFile
   Open DEMOPATH$ + "\LABELS\" + AGUARDA$ For Binary Access Write As #N1%
   Put #N1%, , TEGRABA$
   Close #N1%
   '
   '
   Screen.MousePointer = 1
   '
End Sub

Sub CARGALEYEN()
   '
   Screen.MousePointer = 11
   '
   AGUARDA$ = UCase$(Trim$(Left$(List1.Text, 8)))
   PPXX% = InStr(AGUARDA$, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
   AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
   ORLEYE% = 0
   For KKI% = 0 To 9
     If Option1(KKI%).Value = True Then
       ORLEYE% = KKI%
     End If
   Next KKI%
   AGUARDA$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
   '
   LOARGUA& = 0
   On Error Resume Next
   LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + AGUARDA$)
   If LOARGUA& < 1 Then
     Screen.MousePointer = 1
     Text5.Text = ""
     Text6.Text = "Información"
     MsgBox "Leyenda no Encontrada"
     Exit Sub
   End If
   '
   N1% = FreeFile
   Open DEMOPATH$ + "\LABELS\" + AGUARDA$ For Binary Access Read As #N1%
   TEGRABA$ = Input(LOARGUA&, #N1%)
   Close #N1%
   '
   PPXX% = InStr(TEGRABA$, "|")
   If PPXX% < 1 Then PPXX% = 1 + Len(TEGRABA$)
   '
   TELEYEN$ = Left$(TEGRABA$, PPXX% - 1)
   Text5.Text = TELEYEN$
   '
   ATRIBU$ = Mid$(TEGRABA$, PPXX% + 1)
   '
   Text6.Text = ATRILE$("TITULE")
   Text1.Text = ATRILE$("XFINAL")
   Text2.Text = ATRILE$("YFINAL")
   Label14.Caption = ATRILE$("FONAME")
   Text3.Text = ATRILE$("FOSIZE")
   Check1.Value = ATRILE$("FOBOLD")
   Check2.Value = ATRILE$("FOITAL")
   Check3.Value = ATRILE$("FUNDER")
   '
   CONTROCOLO% = 1
   ROJO(0) = Val(ATRILE$("LEROJO"))
   VERDE(0) = Val(ATRILE$("LEVERD"))
   AZUL(0) = Val(ATRILE$("LEAZUL"))
   If Option2(0).Value = True Then
     HScroll1.Value = ROJO(0)
     HScroll2.Value = VERDE(0)
     HScroll3.Value = AZUL(0)
   End If
   '
   ROJO(1) = Val(ATRILE$("FOROJO"))
   VERDE(1) = Val(ATRILE$("FOVERD"))
   AZUL(1) = Val(ATRILE$("FOAZUL"))
   If Option2(1).Value = True Then
     HScroll1.Value = ROJO(1)
     HScroll2.Value = VERDE(1)
     HScroll3.Value = AZUL(1)
   End If
   '
   ROJO(2) = Val(ATRILE$("PUROJO"))
   VERDE(2) = Val(ATRILE$("PUVERD"))
   AZUL(2) = Val(ATRILE$("PUAZUL"))
   If Option2(2).Value = True Then
     HScroll1.Value = ROJO(2)
     HScroll2.Value = VERDE(2)
     HScroll3.Value = AZUL(2)
   End If
   '
   ALINE% = Val(ATRILE$("ALINEA"))
   Option11(ALINE%).Value = True
   '
   ORIFLECHA% = Val(ATRILE$("ORIFLE"))
   Label10.Caption = Val(ATRILE$("TIPFLE"))
   Text8.Text = ATRILE$("XFLECH")
   Text7.Text = ATRILE$("YFLECH")
   '
   Call AJUCOLOR
   CONTROCOLO% = 0
   '
   Call AJUSTANCHO
   '
   Screen.MousePointer = 1
   '
End Sub

Function ATRILE$(TEBUS$)
   '
   ATRX$ = ""
   PPXX% = InStr(ATRIBU$, TEBUS$ + "=")
   If PPXX% > 0 Then
     PPYY% = InStr(PPXX% + 1, ATRIBU$, "\")
     If PPYY% <= PPXX% Then PPYY% = 1 + Len(ATRIBU$)
     ATRIPAR$ = Mid$(ATRIBU$, PPXX%, PPYY% - PPXX%)
     ATRX$ = Mid$(ATRIPAR$, 2 + Len(TEBUS$))
   End If
   '
   ATRILE$ = ATRX$
   '
End Function

Function EXISTIMAGE%(AIMAGEN$)
   '
   EXI% = 0
   On Error Resume Next
   FILEN = FileLen(AIMAGEN$)
   On Error GoTo 0
   If FILEN > 0 Then
     EXI% = 1
   End If
   '
   EXISTIMAGE% = EXI%
   '
End Function
'
Sub MUESTRAPANDEMO()
   '
   ADEMO$ = UCase$(Trim$(Mid$(List1.Text, 91, 16)))
   LDEMO$ = Label8.Caption
   '
   If EXISTIMAGE%(DEMOPATH$ + "\SCREENS\" + ADEMO$) < 1 Then Exit Sub
   FORDEMO.Image1.Picture = LoadPicture("")
   '
   For KKI% = 0 To 9
     FORDEMO.Picture2(KKI%).Visible = False
   Next KKI%
   '
   On Error GoTo 30
   FORDEMO.Image1.Picture = LoadPicture(DEMOPATH$ + "\SCREENS\" + ADEMO$)
   FORDEMO.Width = FORDEMO.Image1.Width
   FORDEMO.Height = FORDEMO.Image1.Height
   FORDEMO.Left = (12000 - FORDEMO.Width) / 2
   FORDEMO.Top = (8700 - FORDEMO.Height) / 2
   FORDEMO.Picture1.Left = FORDEMO.Width - FORDEMO.Picture1.Width - 30
   FORDEMO.Label1 = LDEMO$
   '
   For KKI% = 0 To 9
     FORDEMO.TRAMO1(KKI%).Visible = False
     FORDEMO.TRAMO2(KKI%).Visible = False
     FORDEMO.TRAMO3(KKI%).Visible = False
     FORDEMO.SOMBRA1(KKI%).Visible = False
     FORDEMO.SOMBRA2(KKI%).Visible = False
     FORDEMO.SOMBRA3(KKI%).Visible = False
     FORDEMO.PUNTA(KKI%).Visible = False
     FORDEMO.PUNTB(KKI%).Visible = False
   Next KKI%
   '
   Call CENTRAFORM(FORDEMO)
   FORDEMO.Show
   On Error GoTo 0
   '
   Call MUESTRALEYENDAS
   '
30 On Error GoTo 0
   '
End Sub

Sub MUESTRALEYENDAS()
   '
   AGUARDA$ = UCase$(Trim$(Left$(List1.Text, 8)))
   PPXX% = InStr(AGUARDA$, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
   AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
   '
   For ORLEYE% = 0 To 9
10   ALEYEN$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
     LOARGUA& = 0
     On Error Resume Next
     LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + ALEYEN$)
     On Error GoTo 0
     If LOARGUA& > 0 Then
        ORDEGUARDA% = ORLEYE%
        Call LEELEYENDA(ALEYEN$)
     End If
19 Next ORLEYE%
   '
End Sub

Sub LEELEYENDA(ALEYEN$)
   '
   LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + ALEYEN$)
   N1% = FreeFile
   Open DEMOPATH$ + "\LABELS\" + ALEYEN$ For Binary Access Read As #N1%
   TEGRABA$ = Input(LOARGUA&, #N1%)
   Close #N1%
   '
   PPXX% = InStr(TEGRABA$, "|")
   If PPXX% < 1 Then PPXX% = 1 + Len(TEGRABA$)
   '
   TELEYEN$ = Left$(TEGRABA$, PPXX% - 1)
   INFODEM.Label3.Caption = TELEYEN$
   If Trim$(TELEYEN$) = "" Then Exit Sub
   '
   ATRIBU$ = Mid$(TEGRABA$, PPXX% + 1)
   '
   TITUA$ = Trim$(ATRILE$("TITULE"))
   TITUE$ = "": For KKI% = 1 To Len(TITUA$)
                  CARA% = Asc(Mid$(TITUA$, KKI%, 1))
                  If CARA% >= 32 Then TITUE$ = TITUE$ + Chr$(CARA%)
                Next KKI%
   INFODEM.Caption = TITUE$
   INFODEM.Left = Val(ATRILE$("XFINAL"))
   INFODEM.Top = Val(ATRILE$("YFINAL"))
   On Error Resume Next
   INFODEM.TEPRUEBA.FontName = ATRILE$("FONAME")
   INFODEM.TEPRUEBA.FontSize = ATRILE$("FOSIZE")
   INFODEM.TEPRUEBA.FontBold = ATRILE$("FOBOLD")
   INFODEM.TEPRUEBA.FontItalic = ATRILE$("FOITAL")
   INFODEM.TEPRUEBA.FontUnderline = ATRILE$("FUNDER")
   '
   INFODEM.Label3.FontName = ATRILE$("FONAME")
   INFODEM.Label3.FontSize = ATRILE$("FOSIZE")
   INFODEM.Label3.FontBold = ATRILE$("FOBOLD")
   INFODEM.Label3.FontItalic = ATRILE$("FOITAL")
   INFODEM.Label3.FontUnderline = ATRILE$("FUNDER")
   '
   ROJOI = Val(ATRILE$("LEROJO"))
   VERDEI = Val(ATRILE$("LEVERD"))
   AZULI = Val(ATRILE$("LEAZUL"))
   INFODEM.Label3.ForeColor = RGB(ROJOI, VERDEI, AZULI)
   '
   ROJOI = Val(ATRILE$("FOROJO"))
   VERDEI = Val(ATRILE$("FOVERD"))
   AZULI = Val(ATRILE$("FOAZUL"))
   INFODEM.Picture2.BackColor = RGB(ROJOI, VERDEI, AZULI)
   '
   ALINE% = Val(ATRILE$("ALINEA"))
   INFODEM.Label3.Alignment = ALINE%
   '
   AMAXI = 0
   TTXX$ = TELEYEN
10 PPXX% = InStr(TTXX$, Chr$(10) + Chr$(13))
   If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
   TEPRU$ = Left$(TTXX$, PPXX% - 1)
   TTXX$ = Mid$(TTXX$, PPXX% + 2)
   INFODEM.TEPRUEBA.Caption = TEPRU$
   ATEPRU = INFODEM.TEPRUEBA.Width
   If ATEPRU > AMAXI Then AMAXI = ATEPRU
   If Len(Trim$(TTXX$)) > 0 Then GoTo 10
   '
   INFODEM.Picture2.Left = 0
   INFODEM.Picture2.Top = 0
   INFODEM.Label3.Left = 50
   INFODEM.Label3.Top = 50
   INFODEM.Label3.Width = AMAXI
   INFODEM.Picture2.Width = INFODEM.Label3.Width + 150
   INFODEM.Picture2.Height = INFODEM.Label3.Height + 150
   INFODEM.Width = INFODEM.Picture2.Width + 60
   INFODEM.Height = INFODEM.Picture2.Height + 300
   '
   On Error GoTo 0
   Call MUESTRATEXTO
   '
   ROJOI = Val(ATRILE$("PUROJO"))
   VERDEI = Val(ATRILE$("PUVERD"))
   AZULI = Val(ATRILE$("PUAZUL"))
   COLOFLE = RGB(ROJOI, VERDEI, AZULI)
   ORIFLE% = Val(ATRILE$("ORIFLE"))
   TIPFLE% = Val(ATRILE$("TIPFLE"))
   XFLECH% = Val(ATRILE$("XFLECH"))
   YFLECH% = Val(ATRILE$("YFLECH"))
   Call DIBUFLECHA(ORIFLE%, TIPFLE%, COLOFLE, XFLECH%, YFLECH%)
   '
   If COMACON$ <> "" Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub MUESTRATEXTO()
   '
   If Trim$(INFODEM.Label3.Caption) = "" Then
     Exit Sub
   End If
   '
   INFODEM.Picture2.Visible = True
   INFODEM.Label3.Visible = True
   INFODEM.Show
   Call RETIENETEXTO
   INFODEM.Hide
   '
End Sub

Sub RETIENETEXTO()
   '
   FORDEMO.Label2(ORDEGUARDA%).Caption = INFODEM.Caption
   FORDEMO.Picture2(ORDEGUARDA%).Height = INFODEM.Height '+ 60
   FORDEMO.Picture2(ORDEGUARDA%).Width = INFODEM.Width '+ 30
   FORDEMO.Picture2(ORDEGUARDA%).Top = INFODEM.Top - 70
   FORDEMO.Picture2(ORDEGUARDA%).Left = INFODEM.Left + 30
   FORDEMO.Label3(ORDEGUARDA%).Width = INFODEM.Label3.Width
   FORDEMO.Label3(ORDEGUARDA%).Left = 50
   FORDEMO.Label3(ORDEGUARDA%).Font = INFODEM.Label3.Font
   FORDEMO.Label3(ORDEGUARDA%).FontSize = INFODEM.Label3.FontSize
   FORDEMO.Label3(ORDEGUARDA%).FontBold = INFODEM.Label3.FontBold
   FORDEMO.Label3(ORDEGUARDA%).FontItalic = INFODEM.Label3.FontItalic
   FORDEMO.Label3(ORDEGUARDA%).FontUnderline = INFODEM.Label3.FontUnderline
   FORDEMO.Label3(ORDEGUARDA%).Alignment = INFODEM.Label3.Alignment
   FORDEMO.Label3(ORDEGUARDA%).Caption = INFODEM.Label3.Caption
   FORDEMO.Label3(ORDEGUARDA%).ForeColor = INFODEM.Label3.ForeColor
   FORDEMO.Line3(ORDEGUARDA%).Y1 = FORDEMO.Picture2(ORDEGUARDA%).ScaleHeight - 20
   FORDEMO.Line3(ORDEGUARDA%).Y2 = FORDEMO.Picture2(ORDEGUARDA%).ScaleHeight - 20
   FORDEMO.Line4(ORDEGUARDA%).X1 = FORDEMO.Picture2(ORDEGUARDA%).ScaleWidth - 20
   FORDEMO.Line4(ORDEGUARDA%).X2 = FORDEMO.Picture2(ORDEGUARDA%).ScaleWidth - 20
   FORDEMO.Picture2(ORDEGUARDA%).BackColor = INFODEM.Picture2.BackColor
   FORDEMO.Picture2(ORDEGUARDA%).Visible = True
   '
End Sub

Sub DIBUFLECHA(ORIFLE%, TIPFLE%, COLOFLE, XFLECH%, YFLECH%)
   '
   FORDEMO.TRAMO1(ORDEGUARDA%).Visible = False
   FORDEMO.TRAMO2(ORDEGUARDA%).Visible = False
   FORDEMO.TRAMO3(ORDEGUARDA%).Visible = False
   FORDEMO.PUNTA(ORDEGUARDA%).Visible = False
   FORDEMO.PUNTB(ORDEGUARDA%).Visible = False
   '
   FORDEMO.SOMBRA1(ORDEGUARDA%).Visible = False
   FORDEMO.SOMBRA2(ORDEGUARDA%).Visible = False
   FORDEMO.SOMBRA3(ORDEGUARDA%).Visible = False
   
   BIZQ = FORDEMO.Picture2(ORDEGUARDA%).Left
   BSUP = FORDEMO.Picture2(ORDEGUARDA%).Top
   ALTU = FORDEMO.Picture2(ORDEGUARDA%).Height
   ANCH = FORDEMO.Picture2(ORDEGUARDA%).Width
   APAR = 500
   ASOMV = 0: ASOMH = 0
   LARFLE = 180
   '
   FORDEMO.TRAMO1(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.TRAMO2(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.TRAMO3(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.PUNTA(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'colofle
   FORDEMO.PUNTB(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'COLOFLE
   FORDEMO.TRAMO1(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.TRAMO2(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.TRAMO3(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.PUNTA(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.PUNTB(ORDEGUARDA%).BorderWidth = 2
   '
   FORDEMO.SOMBRA1(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'COLOFLE
   FORDEMO.SOMBRA2(ORDEGUARDA%).BorderColor = RGB(0, 0, 0)
   FORDEMO.SOMBRA3(ORDEGUARDA%).BorderColor = RGB(0, 0, 0)
   FORDEMO.SOMBRA1(ORDEGUARDA%).BorderWidth = 4
   FORDEMO.SOMBRA2(ORDEGUARDA%).BorderWidth = 4
   FORDEMO.SOMBRA3(ORDEGUARDA%).BorderWidth = 4
   '
   If TIPFLE% < 1 Or TIPFLE% > 3 Then TIPFLE% = 3
   '
   If TIPFLE% = 1 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
          Case 2
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
          Case 3
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
          Case 4
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + 200
        End Select
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   If TIPFLE% = 2 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = XFLECH%
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1
          Case 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X1
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = YFLECH%
          Case 3
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = XFLECH%
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1
          Case 4
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + 200
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X1
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = YFLECH%
        End Select
        FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO2(ORDEGUARDA%).X2
        FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO2(ORDEGUARDA%).Y2
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   If TIPFLE% = 3 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = BIZQ + ANCH + APAR
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = YFLECH%
          Case 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X1
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = BSUP + ALTU + APAR
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = XFLECH%
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
          Case 3
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = BIZQ - APAR
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = YFLECH%
         Case 4
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + 200
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X1
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = BSUP - APAR
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = XFLECH%
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
        End Select
        '
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO2(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
        FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
        FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO3(ORDEGUARDA%).X1
        FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO3(ORDEGUARDA%).Y1
        FORDEMO.TRAMO1(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   XORI% = FORDEMO.TRAMO3(ORDEGUARDA%).X1
   YORI% = FORDEMO.TRAMO3(ORDEGUARDA%).Y1
   ORIENTA = 90
   If Abs(XORI% - XFLECH%) > 10 Then
     ORIENTA = 360 * Atn((YFLECH% - YORI%) / (XORI% - XFLECH%)) / 6.283
   End If
   If XORI% > XFLECH% Then ORIENTA = ORIENTA + 180
   If ORIENTA = 90 Then
     If YORI% < YFLECH% Then ORIENTA = ORIENTA + 180
   End If
   FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH% - 70 * Cos(6.28 * ORIENTA / 360)
   FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH% + 70 * Sin(6.28 * ORIENTA / 360)
   '
   If FORDEMO.TRAMO1(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA1(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA1(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA1(ORDEGUARDA%).Visible = True
   End If
   '
   If FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA2(ORDEGUARDA%).X1 = FORDEMO.TRAMO2(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA2(ORDEGUARDA%).Y1 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA2(ORDEGUARDA).Visible = True
   End If
   '
   If FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA3(ORDEGUARDA%).X1 = FORDEMO.TRAMO3(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO3(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA3(ORDEGUARDA%).X2 = FORDEMO.TRAMO3(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA3(ORDEGUARDA%).Y2 = FORDEMO.TRAMO3(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA3(ORDEGUARDA).Visible = True
     '
     FORDEMO.PUNTA(ORDEGUARDA%).X1 = XFLECH%
     FORDEMO.PUNTA(ORDEGUARDA%).Y1 = YFLECH%
     FORDEMO.PUNTA(ORDEGUARDA%).X2 = XFLECH% - LARFLE * Cos(6.283 * (ORIENTA - 25) / 360)
     FORDEMO.PUNTA(ORDEGUARDA%).Y2 = YFLECH% + LARFLE * Sin(6.283 * (ORIENTA - 25) / 360)
     FORDEMO.PUNTA(ORDEGUARDA%).Visible = True
     '
     FORDEMO.PUNTB(ORDEGUARDA%).X1 = XFLECH%
     FORDEMO.PUNTB(ORDEGUARDA%).Y1 = YFLECH%
     FORDEMO.PUNTB(ORDEGUARDA%).X2 = XFLECH% - LARFLE * Cos(6.283 * (ORIENTA + 25) / 360)
     FORDEMO.PUNTB(ORDEGUARDA%).Y2 = YFLECH% + LARFLE * Sin(6.283 * (ORIENTA + 25) / 360)
     FORDEMO.PUNTB(ORDEGUARDA%).Visible = True
     '
   End If
   '
End Sub
'
Sub CARLISECUEN(SECU$)
   '
   SECUEN$ = Trim$(UCase$(SECU$))
   If SECUEN$ = "" Then Exit Sub
   ASECU$ = SECUEN$ + ".SEC"
   '
   List1.Clear
   '
   N1% = FreeFile
   Dim BUFLE1 As String * 256
   Open DEMOPATH$ + "\SCRIPTS\" + ASECU$ For Random Access Read Write As #N1% Len = 256
   '
   For II% = 1 To LOF(N1%) / 256
      Get #N1%, II%, BUFLE1$
      INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
      ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
      LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
      PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
      PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
      MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
      '
      If Trim$(INSTRU$) <> "" Then
        TTXX$ = Space$(256)
        Mid$(TTXX$, 1, 8) = INSTRU$
        Mid$(TTXX$, 11, 80) = LEYEN$
        Mid$(TTXX$, 91, 16) = ARIMA$
        Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
        Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
        Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
        List1.AddItem TTXX$
      End If
   Next II%
   Close #N1%
   '
   List1.Width = 1400
   If List1.ListCount <= 5 Then
     List1.Width = 1200
   End If
   '
End Sub
'
Sub CARLISECUTO(SECU$)
   '
   'SECUENDE.LIST1.Visible = True
   'SECUENDE.LIST1.Width = 11500
   'SECUENDE.LIST1.Height = 6000
   SECUENDE.List1.Clear
   '
   SECUEN$ = Trim$(UCase$(SECU$))
   If SECUEN$ = "" Then Exit Sub
   ASECU$ = SECUEN$ + ".SEC"
   LOSECU& = 0
   On Error Resume Next
   LOSECU& = FileLen(DEMOPATH$ + "\SCRIPTS\" + ASECU$)
   On Error GoTo 0
   If LOSECU& < 128 Then
     MsgBox "Secuencia '" + ASECU$ + "' no Instalada."
     Exit Sub
   End If
   '
   N1% = FreeFile
   Dim BUFLE1 As String * 256
   Open DEMOPATH$ + "\SCRIPTS\" + ASECU$ For Random Access Read Shared As #N1% Len = 256
   '
   UFECHA# = 0
   For II% = 1 To LOF(N1%) / 256
      Get #N1%, II%, BUFLE1$
      INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
      ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
      LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
      PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
      PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
      MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
      '
      If Trim$(INSTRU$) <> "" Then
        TTXX$ = Space$(256)
        Mid$(TTXX$, 1, 8) = INSTRU$
        Mid$(TTXX$, 11, 80) = LEYEN$
        Mid$(TTXX$, 91, 16) = ARIMA$
        Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
        Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
        Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
        SECUENDE.List1.AddItem TTXX$
      End If
      On Error Resume Next
      AFECHA# = FileDateTime(DEMOPATH$ + "\SCREENS\" + ARIMA$)
      On Error GoTo 0
      If AFECHA# < Now Then
        If AFECHA# > UFECHA# Then
          UFECHA# = AFECHA#
        End If
      End If
   Next II%
   '
   Close #N1%
   FEDICION$ = Format(UFECHA#, "dd/mm/yyyy") '
   '
10 AGREGA% = 0
   For JJ% = 1 To SECUENDE.List1.ListCount
     TTYY$ = SECUENDE.List1.List(JJ% - 1)
     ARCHIX$ = UCase$(Trim$(Mid$(TTYY$, 91, 16)))
     If InStr(ARCHIX$, ".SEC") > 1 Then
       SECUENDE.List1.RemoveItem JJ% - 1
       POINSE% = JJ% - 1
       Open DEMOPATH$ + "\SCRIPTS\" + ARCHIX$ For Random Access Read Shared As #N1% Len = 256
       '
       For II% = 1 To LOF(N1%) / 256
         Get #N1%, II%, BUFLE1$
         INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
         ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
         LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
         PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
         PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
         MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
         '
         If Trim$(INSTRU$) <> "" Then
           TTXX$ = Space$(256)
           Mid$(TTXX$, 1, 8) = INSTRU$
           Mid$(TTXX$, 11, 80) = LEYEN$
           Mid$(TTXX$, 91, 16) = ARIMA$
           Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
           Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
           Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
           SECUENDE.List1.AddItem TTXX$, POINSE%
           POINSE% = POINSE% + 1
           AGREGA% = 1
         End If
       Next II%
       Close #N1%
     End If
     If AGREGA% > 0 Then GoTo 10
     '
   Next JJ%
   '
End Sub
'
Private Sub VScroll2_Change()
   If MUEVELE% = 0 Then
     Text7.Text = VScroll2.Value
   End If
End Sub
