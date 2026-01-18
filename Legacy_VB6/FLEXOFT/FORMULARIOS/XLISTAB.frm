VERSION 5.00
Begin VB.Form XLISTAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   FillStyle       =   0  'Solid
   Icon            =   "XLISTAB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Impresora:"
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
      Left            =   525
      TabIndex        =   40
      Top             =   2600
      Width           =   1250
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Hoja de Cálculo (Planilla Excel)"
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
      Height          =   540
      Index           =   2
      Left            =   6825
      TabIndex        =   39
      Top             =   3700
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Archivo de Texto"
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
      Left            =   6825
      TabIndex        =   38
      Top             =   1890
      Width           =   1905
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   6510
      TabIndex        =   30
      Top             =   1420
      Width           =   5160
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   46
         Text            =   " "
         Top             =   3150
         Width           =   2430
      End
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   315
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   2520
         TabIndex        =   36
         Top             =   735
         Width           =   2430
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Picture         =   "XLISTAB.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   300
         Width           =   540
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Opciones"
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
         Left            =   160
         TabIndex        =   31
         Top             =   840
         Width           =   2010
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Mostrar Archivo"
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
            TabIndex        =   34
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
            Caption         =   "Agregar al Final"
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
            TabIndex        =   33
            Top             =   315
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
            Caption         =   "Sobre-Escribir"
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
            TabIndex        =   32
            Top             =   525
            Width           =   1695
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000009&
         X1              =   2325
         X2              =   2325
         Y1              =   105
         Y2              =   2940
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   2310
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   5145
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
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
         Left            =   -210
         TabIndex        =   49
         Top             =   3320
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   420
         TabIndex        =   48
         Top             =   3570
         Width           =   4530
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         TabIndex        =   47
         Top             =   3190
         Width           =   1065
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5145
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   2310
         X2              =   2310
         Y1              =   105
         Y2              =   2940
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2310
         Y1              =   2100
         Y2              =   2100
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6510
      ScaleHeight     =   75
      ScaleWidth      =   1215
      TabIndex        =   28
      Top             =   1050
      Width           =   1275
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   12000
      End
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
      Height          =   855
      Left            =   8925
      Picture         =   "XLISTAB.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Guarda Configuración del Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Control de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1720
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   6105
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   1760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   3105
         X2              =   3105
         Y1              =   120
         Y2              =   1760
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   210
         TabIndex        =   45
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   1365
         TabIndex        =   44
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Controlador:"
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
         TabIndex        =   43
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3255
         TabIndex        =   42
         Top             =   840
         Width           =   2640
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   4305
         TabIndex        =   41
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         Left            =   3150
         TabIndex        =   15
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   7980
      Picture         =   "XLISTAB.frx":0A8E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura Impresora"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
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
      Left            =   9870
      Picture         =   "XLISTAB.frx":0ED0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Emite Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton BOTEXIT 
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
      Height          =   855
      Left            =   10815
      Picture         =   "XLISTAB.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona el Listador"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Imprimir"
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
      TabIndex        =   0
      Top             =   2050
      Width           =   6105
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   23
         Text            =   " "
         Top             =   840
         Width           =   2955
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Medidas en Milimetros"
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
         Left            =   1575
         TabIndex        =   16
         Top             =   2100
         Width           =   4320
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   3
            Left            =   3255
            TabIndex        =   25
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   3255
            TabIndex        =   24
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   21
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   19
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Sup:"
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
            Left            =   1995
            TabIndex        =   27
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierdo:"
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
            Left            =   2310
            TabIndex        =   26
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
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
            Height          =   225
            Index           =   1
            Left            =   -105
            TabIndex        =   20
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
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
            Height          =   225
            Index           =   0
            Left            =   -105
            TabIndex        =   18
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0EEFF&
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
         Height          =   780
         Left            =   4830
         TabIndex        =   12
         Top             =   340
         Width           =   1065
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   210
            TabIndex        =   13
            Text            =   " "
            Top             =   350
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Copias"
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
         Left            =   4830
         TabIndex        =   10
         Top             =   1260
         Width           =   1065
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   315
            TabIndex        =   11
            Text            =   " "
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Orientación"
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
         Left            =   1575
         TabIndex        =   7
         Top             =   1260
         Width           =   2955
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Paisaje"
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
            Left            =   1785
            TabIndex        =   9
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Retrato"
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
            TabIndex        =   8
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   6
         Top             =   4305
         Width           =   3690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Papel:"
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
         Left            =   630
         TabIndex        =   17
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fuente:"
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
         Left            =   630
         TabIndex        =   5
         Top             =   900
         Width           =   855
      End
   End
   Begin VB.Label Label13 
      Height          =   330
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
End
Attribute VB_Name = "XLISTAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
Dim VDS$(32, 16), VDX#(32, 16), VD%(32, 16), VFS$(32), VFSA$(32), VFX#(32), VF%(32), FI$(32)
Dim VDP%(32), VDPX#(32), VDPS$(32), VDA$(32), VDAA%(32), VDAB#(32)
Dim SUMA#(32, 2), SUPAR#(32)
Dim LEYE$(16)
'
Dim ARCHISAL$, ASAL%
Dim SUBTI$(16)

Dim PRINTERSET%, LARGOJA, ANCHOJA, ALTUREN, TALEBASI, TAFUESEL
Dim TOTALIX%, POSITOT%, RESUB$, RETRA$, ULTIHOJ%
Dim HAYTOTA%
Dim CXC$
Dim NUHOJ%
Dim DATLISTOK%, SOBRESCRI%
Dim SRY As String * 255
Dim GU As String * 255
Dim LII$(10)
Dim CCC$(10)
'
Dim PPZZ As String * 128
Dim FUFIJA$(256)
'
Dim ANCHOPAP, LARGOPAP, MARGENI, MARGENS
Dim ORIENTA%, CACOPIAS%
'
Dim SALEXCEL%
'

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    Hide
End Sub

Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For I% = 1 To Printers.Count
        If Printers(I% - 1).DeviceName = Combo1.Text Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(I% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.Text = Printer.DeviceName
    '
End Sub

Private Sub Command1_Click()
    '
    If Option1(0) = False Then
       ASALIDA$ = TRIM$(Text2.Text)
       If ASALIDA$ = "" Then
          Call MENSERR(24, "Falta Archivo de Exportación")
          Exit Sub
       End If
       PPXY% = InStr(ASALIDA$, ".")
       If PPXY% > 0 Then
          Text2 = Left$(ASALIDA$, PPXY%)
          Call MENSERR(24, "Suprima Extensión de Archivo de Exportación")
          Exit Sub
       End If
       If Len(Label11) > 40 Then
          Call MENSERR(24, "Ruta de Salida Demasiado Larga.")
          Exit Sub
       End If
       ARCHISAL$ = TRIM$(Label11)
       If Len(ARCHISAL$) > 0 Then
         If Right$(ARCHISAL$, 1) <> "\" Then ARCHISAL$ = ARCHISAL$ + "\"
       End If
       ARCHISAL$ = ARCHISAL$ + ASALIDA$
       If Option1(1).Value = True Then
            If InStr(TRIM$(ARCHISAL$), " ") > 0 Then
               Call MENSERR(24, "Archivo o Carpeta de Salida no Válidos.\  \No se Permiten Espacios.\Cambie Archivo o Carpeta de Salida.")
               Exit Sub
            End If
            ARCHISAL$ = ARCHISAL$ + ".TXT"
            Call GRACONTROL("EXPORTEX", IDENTER$, Label11)
          Else
            ARCHISAL$ = ARCHISAL$ + ".XLS"
            Call GRACONTROL("EXPOREXC", IDENTER$, Label11)
       End If
    End If
    '
    If TRIM$(IMPREDETER$) <> TRIM$(Combo1.Text) Then
        'MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
        'MENSA$ = MENSA$ + "'" + TRIM$(Combo1.Text) + "'\   \"
        'MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
        'MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
        'MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
        'MENSA$ = MENSA$ + "Predeterminada para Windows."
        'Combo1.Text = IMPREDETER$
        'Call COMUNI(MENSA$)
        'Combo1.SetFocus
        'Exit Sub
        Printer.TrackDefault = False
        NPPXX$ = TRIM$(Combo1.Text)
        For Each PRX9 In Printers
          If PRX9.DeviceName = NPPXX$ Then
            On Error Resume Next
            Set Printer = PRX9
            On Error GoTo 0
            GoTo 10
          End If
        Next PRX9
        '
10      If Printer.DeviceName <> Combo1.Text Then
          MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(Combo1.Text) + "'\   \"
          MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
          MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
          MENSA$ = MENSA$ + "Predeterminada para Windows."
          Combo1.Text = IMPREDETER$
          Call COMUNI(MENSA$)
          Combo1.SetFocus
          Exit Sub
        End If
    End If
    '
    Call GRADATIMPRE
    Call CARDATIMPRE
    '
    SOBRESCRI% = 0
    If Option3(1) = True Then
        SOBRESCRI% = 1
    End If
    '
    PORIAN = Printer.Orientation
    If Option2(1).Value = True Then
          Printer.Orientation = 2
        Else
          Printer.Orientation = 1
    End If
    '
    Call VERDATLIS
    INTERLIX = 0
    If REHOJX% > 96 Then INTERLIX = Printer.ScaleHeight / REHOJX%
    '
    If DATLISTOK% <= 0 Then
         On Error Resume Next
         If TRIM$(ARCHISAL$) = "" Then Printer.Orientation = PORIAN
         On Error GoTo 0
         Exit Sub
      ElseIf DATLISTOK% = 1 Then
         On Error GoTo 199
         If SALEXCEL% = 1 Then
           Call ARMAENCA
           Call LISEXCEL
           Call CIERRARCH("*.*")
           Hide
           Exit Sub
         End If
         Call SETPRINTER
         Call BLANHEADERS
         Call CARHEADERS
         Call ARMAENCA
         If CACOPIAS% < 1 Then CACOPIAS% = 1
         CACOP% = CACOPIAS%
         If ARCHISAL$ <> "" Then CACOP% = 1
         For COP% = 1 To CACOP%
            Call INITOTA
            If TNC% < 1 Then
                Call ENCAHOJA
            End If
            Call LISTABU
         Next COP%
         On Error GoTo 0
    End If
    '
    Call CIERRARCH("*.*")
    PRINTERSET% = 0
    Printer.Orientation = PORIAN
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    '
    PROLLAMA$ = ""
    If ARCHISAL$ <> "" Then
        If Check1.Value = 1 Then
            LCB$ = LUBAS$
            If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
                 Shell "C:\FLEXOFT\WORDPAD.EXE " + ARCHISAL$, vbMaximizedFocus
                 PROLLAMA$ = TRIM$(Text2) + " - Wordpad"
               Else
                 MENSA$ = "Imposible Mostrar Archivo Generado."
                 MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
                 MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
                 MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
                 Call COMUNI(MENSA$)
            End If
        End If
    End If
    Hide
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Private Sub Command2_Click()
    Call SETUPRINTER
End Sub

Private Sub Command3_Click()
    Call GRADATIMPRE
    Call CARDATIMPRE
End Sub


Private Sub Dir1_Change()
   RUTASAL$ = Dir1.Path
   Label11.Caption = RUTASAL$
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub Form_Activate()
   '
   ' ya revisado parpadeo
   '
   DoEvents
   '
End Sub

Private Sub Form_GotFocus()
   ' ya revisado parpadeo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ' ya revisado parpadeo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   ' ya revisado parpadeo
End Sub

Private Sub Form_Load()
    '
Static TIMA
If Timer < TIMA + 0.4 Then
MsgBox "load"
End If
TIMA = Timer
    Refresh
    PORIAN = Printer.Orientation
    '
    Printer.TrackDefault = False
    '
    Call CARDATIMPRE
    '
    ARCHISAL$ = ""
    Text2.Text = ARCHISAL$
    Option1(0).Value = True        ' IMPRESORA
    '
    SOBRESCRI% = 0
    FORMALI$ = FORMA$
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
        Combo1.AddItem PRX.DeviceName
        ' La define como predeterminada del sistema.
        'Set Printer = prx
    Next
    Combo1.Text = Printer.DeviceName
    IMPREDETER$ = Printer.DeviceName
    '
    If Left$(UCase$(ATRIBU$), 4) = "LPT1" Then
        Refresh
        Call SETPRINTER
        Call BLANHEADERS
        Call CARHEADERS
        Call ARMAENCA
        Call INITOTA
        If TNC% < 1 Then
            Call ENCAHOJA
        End If
        Call LISTABU
        Call CIERRARCH("*.*")
        PRINTERSET% = 0
        If COMMA$ <> "" Then
           Call BLOQEXE("*")
        End If
        Hide
    End If
    '
    Printer.Orientation = PORIAN
    If EXISTE%("FLEXCEL.EXE") > 0 Then
      Option1(2).Enabled = True
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Private Sub Label13_Change()
   Visible = False
   If Label13 <> "" Then
     DoEvents
     Call CARDATLIS
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        ARCHISAL$ = ""
        Text2.Text = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Drive1.Enabled = False
        Dir1.Enabled = False
        Command4.Enabled = False
        SALEXCEL% = 0
      ElseIf Index = 1 Then
        RUTASAL$ = TRIM$(CONTROL$("EXPORTEX", IDENTER$))
        If RUTASAL$ = "" Then RUTASAL$ = Left$(LUDAT$, Len(LUDAT$) - 1)
        On Error Resume Next
        Drive1.Drive = Left$(RUTASAL$, 2)
        Dir1.Path = RUTASAL$
        On Error GoTo 0
        Label11.Caption = RUTASAL$
        ASALIDA$ = TRIM$(ARCHITRA$)
        Text2.Text = ASALIDA$
        Option3(0).Value = True
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Check1.Enabled = True
        Drive1.Enabled = True
        Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 0
        Text2.Enabled = True
        Text2.SetFocus
      ElseIf Index = 2 Then
        RUTASAL$ = TRIM$(CONTROL$("EXPOREXC", IDENTER$))
        If RUTASAL$ = "" Then RUTASAL$ = Left$(LUDAT$, Len(LUDAT$) - 1)
        On Error Resume Next
        Drive1.Drive = Left$(RUTASAL$, 2)
        Dir1.Path = RUTASAL$
        On Error GoTo 0
        Label11.Caption = RUTASAL$
        ASALIDA$ = TRIM$(ARCHITRA$)
        Text2.Text = ASALIDA$
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Drive1.Enabled = True
        Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 1
        Text2.Enabled = True
        Text2.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub
'

Private Sub Text1_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub

Sub CARDATLIS()
     '
5800 GRABACT% = 0
     TOTALIX% = 0
     POSITOT% = 0
     '
     For KKI% = 0 To 32
       VFS$(KKI%) = ""
       VFSA$(KKI%) = ""
       VFX#(KKI%) = 0
       VF%(KKI%) = 0
       VDP%(KKI%) = 0
       VDPX#(KKI%) = 0
       VDPS$(KKI%) = ""
       VDA$(KKI%) = ""
       VDAA%(KKI%) = 0
       VDAB#(KKI%) = 0
       SUMA#(KKI%, 0) = 0: SUMA#(KKI%, 1) = 0: SUMA#(KKI%, 2) = 0
       SUPAR#(KKI%) = 0
     Next KKI%
     '
     FORMA$ = Label13
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(Archi$)
     Label5 = ARCHITRA$
     Label6 = TITUARCH$(ARCHITRA$)
     If Text2.Enabled Then Text2 = ARCHITRA$
5801 Get #N15%, 1, YYY$
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     Label8 = PROPRIN$
5803 S1$ = TRIM$(Left$(YYY$, 48))
     Label9 = TRIM$(S1$)
     '
     EPAC$ = TRIM$(EMPREAC$)
     If EPAC$ <> "" Then
       If InStr(Caption, EPAC$) < 1 Then
         Caption = Caption + "  -  " + EPAC$
       End If
     End If
    '
     Sali$ = Mid$(YYY$, 49, 8)
     ARCHISAL$ = Mid$(YYY$, 65, 12)
     If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
        CXC$ = "": If BLP% = 1 Then CXC$ = "¦"
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = REHOJ%: REHOJ1% = REHOJ%
     If Asc(Mid$(COD$, 7)) < REHOJ% Then IDL% = Asc(Mid$(COD$, 7))
     '
     If VIV% = 1 Then IDL% = 8 * IDL% / 6
     SUPRIFE% = 0
     HDR& = Asc(Mid$(COD$, 8))
     If HDR& = 99 Then
        HDR& = 0
        SUPRIFE% = 1
     End If
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: TNC% = 0
     RG% = 1: HST% = 0
     HOJINI% = CVI(Mid$(COD$, 10, 2))
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
         If RG% = 2 Then TVL%(TCL%) = 1
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     While PR%(TCL%) < PR%(TCL% - 1)
       PR%(TCL%) = PR%(TCL%) + 256
     Wend
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then PR%(TCL%) = 0: LF%(TCL%) = 2
         '
     If Asc(Mid$(YYY$, 72)) = 99 Then
        TNC% = TNC% + 1
        CLNC%(TNC%) = TCL%
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5817   ' SUPRIME COLUMNAS LONGITUD 0
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     TT%(NCL%) = Asc(Mid$(YYY$, 73))
        If TT%(NCL%) = 1 Then TT%(NCL%) = TCL%
        If TT%(NCL%) = 255 Then HST% = 1
5815 PE%(NCL%) = 1
        If RG% = 2 Then PE%(NCL%) = 2
5816 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
     If TT%(NCL%) > 0 Then
        NTT% = NTT% + 1
        TOTALIX% = 1
     End If
     '
     If POSITOT% = 0 Then
        If TT%(NCL%) < 1 Then
           If TVL%(TCL%) = 6 Then
              If ACL%(NCL%) >= 10 Then
                 POSITOT% = NCL%
              End If
           End If
        End If
     End If
     
5817 GoTo 5805
     '
5820 If POSITOT% = 0 Then POSITOT% = 1
     ATOT% = BLP% * (NCL% + 1)
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
          LC%(KCL%) = ACL%(KCL%)
          ATOT% = ATOT% + ACL%(KCL%) + MG%(KCL%)
          CL%(KCL%) = CLII%
          CLII% = CLII% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
          If CE%(KCL%) <= 0 Then GoTo 5824
5822      For KCL1% = 1 To NCL%
              If CL0%(KCL1%) = CE%(KCL%) Then CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
5823      Next KCL1%
5824 Next KCL%
5825 For KCL% = 1 To NCL%
          If TT%(KCL%) <= 0 Then GoTo 5828
5826      For KCL1% = 1 To NCL%
             If CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = KCL1%: GoTo 5828
             If 100 + CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = 100 + KCL1%: GoTo 5828
5827      Next KCL1%
5828 Next KCL%
     '
5830 NFT% = 0    ' NO HAY FILTROS
5850 If N15% > 0 Then
         Close #N15%
         N15% = 0
     End If
     '
End Sub
'

Sub CARDATIMPRE()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    MARGENI = 0
    MARGENS = 0
    ORIENTA% = 1
    CACOPIAS% = 1
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN("C:\FLEXOFT\LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             ANCHOPAP1 = CVS(Mid$(PPXX$, 49, 4))
             LARGOPAP1 = CVS(Mid$(PPXX$, 53, 4))
             '
             If ANCHOPAP1 > 0 Then
                If ANCHOPAP1 < ANCHOPAP Then
                   ANCHOPAP = ANCHOPAP1
                End If
             End If
             '
             If LARGOPAP1 > 0 Then
                If LARGOPAP1 < LARGOPAP Then
                   LARGOPAP = LARGOPAP1
                End If
             End If
             '
             GoTo 10
             '
           End If
         End If
    Next IPRIN&
    '
10  Close #NX%
    NX% = FILEOPEN%("C:\FLEXOFT\LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             MARGENI1 = CVS(Mid$(PPXX$, 9, 4))
             MARGENS1 = CVS(Mid$(PPXX$, 13, 4))
             ORIENTA1% = Asc(Mid$(PPXX$, 17, 1))
             CACOPIAS1% = Asc(Mid$(PPXX$, 18, 1))
             '
             If MARGENI1 > MARGENI Then
                 MARGENI = MARGENI1
             End If
             '
             If MARGENS1 > MARGENS Then
                   MARGENS = MARGENS1
             End If
             '
             If ORIENTA1% > 0 Then
                If ORIENTA1% <= 2 Then
                    ORIENTA% = ORIENTA1%
                End If
             End If
             '
             If CACOPIAS1% > 0 Then
                If CACOPIAS1% <= 32 Then
                    CACOPIAS% = CACOPIAS1%
                End If
             End If
             '
             GoTo 20
             '
         End If
    Next IPRIN&
    '
20  If ORIENTA% = 1 Then
         Option2(0).Value = True        ' RETRATO
       Else
         Option2(1).Value = True        ' PAISAJE
    End If
    Text4(0).Text = CSTRING$(MKS$(ANCHOPAP), 3, 6, 1, 1)
    Text4(1).Text = CSTRING$(MKS$(LARGOPAP), 3, 6, 1, 1)
    Text4(2).Text = CSTRING$(MKS$(MARGENS), 3, 6, 1, 1)
    Text4(3).Text = CSTRING$(MKS$(MARGENI), 3, 6, 1, 1)
    Text6.Text = CACOPIAS%
    '
End Sub


Sub GRADATIMPRE()
    '
5   On Error GoTo 199
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    On Error GoTo 0
    '
    If Val(Text4(0).Text) > 0 Then
        If Val(Text4(0).Text) < ANCHOPAP Then
            ANCHOPAP = Val(Text4(0).Text)
        End If
    End If
    If Val(Text4(1).Text) > 0 Then
        If Val(Text4(1).Text) < LARGOPAP Then
            LARGOPAP = Val(Text4(1).Text)
        End If
    End If
    MARGENI = Val(Text4(3).Text)
    MARGENS = Val(Text4(2).Text)
    ORIENTA% = 1
    If Option2(1).Value = True Then
        ORIENTA% = 2
    End If
    CACOPIAS% = 1
    If Val(Text6.Text) > 0 Then
        If Val(Text6.Text) <= 32 Then
            CACOPIAS% = Val(Text6.Text)
        End If
    End If
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN("C:\FLEXOFT\LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             Call REPLA(PPXX$, MKS$(ANCHOPAP), 49, 4)
             Call REPLA(PPXX$, MKS$(LARGOPAP), 53, 4)
             PPZZ$ = PPXX$
             Put #NX%, IPRIN&, PPZZ$
             GoTo 9
           End If
         End If
    Next IPRIN&
    PPXX$ = Space$(48) + String$(80, 0)
    Call REPLA(PPXX$, NOMIMPRE$, 1, 48)
    Call REPLA(PPXX$, MKS$(ANCHOPAP), 49, 4)
    Call REPLA(PPXX$, MKS$(LARGOPAP), 53, 4)
    Call REPLA(PPXX$, PROPRIN$, 57, 8)
    PPZZ$ = PPXX$
    Put #NX%, FIN& + 1, PPZZ$
9   Close #NX%
    '
10  NX% = FILEOPEN%("C:\FLEXOFT\LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             GoTo 20
         End If
    Next IPRIN&
    '
    IPRIN& = FIN& + 1
    PPXX$ = String$(128, 0)
    '
20  Call REPLA(PPXX$, PROPRIN$, 1, 8)
    Call REPLA(PPXX$, MKS$(MARGENI), 9, 4)
    Call REPLA(PPXX$, MKS$(MARGENS), 13, 4)
    Call REPLA(PPXX$, Chr$(ORIENTA%), 17, 1)
    Call REPLA(PPXX$, Chr$(CACOPIAS%), 18, 1)
    PPZZ$ = PPXX$
    Put #NX%, IPRIN&, PPZZ$
    Close #NX%
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 5
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Sub VERDATLIS()
    '
    DATLISTOK% = 0
    If ULTREG&(ARCHITRA$) < 1 Then
        Call MENSERR(24, "Archivo de Trabajo Vacío")
        DATLISTOK% = (-1)
        Exit Sub
    End If
    '
    If SALEXCEL% < 1 Then
      ASALI$ = TRIM$(UCase$(Text2.Text))
      If ASALI$ = "" Then
         DATLISTOK% = 1
         Exit Sub
      End If
    End If
    '
    If SALEXCEL% = 1 Then
      ASAL% = FILEOPEN%("C:\FLEXOFT\TREXCEL.TRF", 2, 0)
      DATLISTOK% = 1
      Exit Sub
    End If
    '
    'PPUN% = InStr(ASALI$, ".")
    'If PPUN% < 1 Then PPUN% = 1 + Len(ASALI$)
    'ASALI$ = Left$(ASALI$, PPUN% - 1) + ".TXT"
    '
    'ARCHISAL$ = ASALI$
    If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
    'Text2.Text = ARCHISAL$
    '
10  On Error GoTo 100
    ASAL% = FreeFile
    If SOBRESCRI% = 1 Then
        ASAL% = FILEOPEN%(ARCHISAL$, 2, 0)
      Else
        ASAL% = FILEOPEN%(ARCHISAL$, 3, 0)
    End If
    DATLISTOK% = 1
    Exit Sub
    '
100 Resume 101
101 Close #ASAL%
    ASAL% = 0
    MsgBox "Error de Acceso al Archivo" + Chr$(10) + Chr$(13) + ARCHISAL$
    DATLISTOK% = 0
    ARCHISAL$ = LUDAT$ + TRIM$(ARCHITRA$) + ".TXT"
    Text2.Text = ARCHISAL$

    '
End Sub

Sub LISTABU()
     
     '
     Screen.MousePointer = 11
     '
     largomax = LARGOJA
     If ORIENTA% = 1 Then
         If largomax > LARGOPAP * 56.7 Then
             largomax = LARGOPAP * 56.7
         End If
       ElseIf ORIENTA% = 2 Then
         If largomax > ANCHOPAP * 56.7 Then
             largomax = ANCHOPAP * 56.7
         End If
     End If
     '
     PRIHOJ% = 1
     If ARCHISAL$ = "" Or Left$(ARCHISAL$, 3) = "LPT" Or ATRAN% = 1 Then
         On Error GoTo 199
         CAHOJ& = ULTREG&(ARCHITRA$) / (IDL% - RGF& - 4)
'         If CAHOJ& >= 3 Then
'             PRIHOJ% = Int(Val(InputBox$("Imprime desde Hoja Numero: 1", 4)))
'         End If
       Else
         IDL% = 32767
     End If
     '
     PRIENCA% = 1
     '
     FINX& = ULTREG&(ARCHITRA$)
     FINDEX% = FILEINDEX%(ARCHITRA$)
     FINUME% = FILENBR%(ARCHITRA$)
     LOREGI% = LOREGIS%(FINDEX%)
     '
     'ProgressBar1.Min = 0
     'If FINX& > 0 Then
     '   ProgressBar1.Max = FINX&
     'End If
     '
     For K1& = HDR& + 1 To FINX&
        '
        'ProgressBar1.Value = K1&
        Call TRACE(20, K1&, FINX&)
        '
        'If BOSTOP%("Pulse Boton de Parada\o Tecla de  <ESCAPE>\para Interrumpir Listador") Then
         ' TINTE$ = Time$
'          146011  X% = COMALTER%("Impresion Interrumpida a su Comando\  \Fecha: " + HO$ + "    Hora: " + TINTE$ + "\\Reanudar\Encuadrar\Abortar Listado")
         ' On X% GoTo 14602, 14603, 14604
'          14602  ' Call TRAZA(21, FINX% + 1, FINX% + 1)
          'GoTo 14605
'          14603   'Call MENSERR(24, "Opcion no Disponible por el Momento")
          'GoTo 146011
'          14604   'Printer.Print , Space$(OFS%) + "INTERRUMPIDO" + Mid$(LII$(3), 13 + OFS%)
          'GoSub 25500                                  ' PRINTER.PRINT, CCI$(2);
          'INTE% = 1
          'GoTo 3800
        'End If
        '
        'XXX$ = REGLEIDO$(ARCHITRA$, K1&)
        XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, K1&)
        If K1& = HDR& + 1 Then
            If TNC% > 0 Then
                Call ENCAHOJA1(XXX$)
                Call ENCAHOJA
            End If
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "=") Then
            ULTIHOJ% = 1
            If K1& < ULTREG&(ARCHITRA$) Then
              Call LITRANSPO
              'XXX2$ = REGLEIDO$(ARCHITRA$, K1& + 1)
              XXX2$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, K1& + 1)
              If ARCHISAL$ <> "" Then
                  Print #ASAL%, "   " ' LII$(0)
                  Print #ASAL%, LII$(4)
                  Print #ASAL%, "   " ' LII$(0)
                  Call INITOTA
                  Call ENCAHOJA1(XXX2$)
                  If TNC% > 0 Then
                      Call ENCAHOJA
                  End If
                Else
                  Call EJECT
                  Printer.NewPage
                  NUHOJ% = 0
                  Call INITOTA
                  Call ENCAHOJA1(XXX2$)
                  Call ENCAHOJA
              End If
            End If
            '
            For K2% = 1 To NCL%
                VFSA$(K2%) = ""
            Next K2%
            GoTo 99
        End If
        '
        If Printer.CurrentY > largomax - (7 + 6 * TOTALIX%) * ALTUREN Then
           '
           If ARCHISAL$ = "" Then
               ULTIHOJ% = 0
               Call LITRANSPO
               '
               Call EJECT
               For K2% = 1 To NCL%
                  VFSA$(K2%) = ""
               Next K2%
               Printer.NewPage
               Call ENCAHOJA1(XXX$)
               Call ENCAHOJA
               '
               If TOTALIX% > 0 Then
                  If ARCHISAL$ = "" Then
                      Printer.Print MIZQ$ + LII$(0)
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Printer.Print MIZQ$ + RETRA$
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Printer.Print MIZQ$ + LII$(0)
                      If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                    Else
                      Print #ASAL%, LII$(0)
                      Print #ASAL%, RETRA$
                      Print #ASAL%, LII$(0)
                  End If
               End If
           End If
           '
        End If
        '
        If XXX$ = REGBLAN$(ARCHITRA$) Then
            RENGLO$ = LII$(0)
            For K3% = 1 To NCL%
               VFSA$(K3%) = ""
            Next K3%
            GoTo 5
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "-") Then
            RENGLO$ = LII$(4)
            GoTo 5
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "*") Then
            RENGLO$ = LII$(5)
            GoTo 5
        End If
        '
        KI% = 0: KF% = 0
        For K2% = 1 To NCL%
            If PE%(K2%) < 3 Then
               If CFT%(K2%) = 1 Then
                     XXZ$ = MKS$(K1& + HDR&)
                     If NFT% > 0 Then XXZ$ = RECNO$
                   Else
                     XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
               End If
               '
               TIVA% = TVL%(CFT%(K2%))
               If TIVA% = 6 Then
                 If Left$(XXZ$, 1) = "@" Then
                   For K3% = 1 To NCL%
                     VFS$(K3%) = Space$(ACL%(K3%))
                   Next K3%
                   VFS$(K2%) = Mid$(XXZ$, 2)
                   GoTo 4
                 End If
               End If
               '
               LOTOT% = ACL%(K2%)
               LODEC% = LDL%(CFT%(K2%))
               JUSTI% = JT%(K2%)
               '
               VFS$(K2%) = CSTRING$(XXZ$, TIVA%, LOTOT%, LODEC%, JUSTI%)
               If CE%(K2%) > 0 Then
                  VFS$(CE%(K2%)) = AJUSTI$(ECOARCH$(AE$(K2%), XXZ$), ACL%(CE%(K2%)))
               End If
               '
               VALO# = 0: If TIVA% <= 4 Then VALO# = Val(VFS$(K2%))
               If TIVA% = 8 Then
                  VALO# = Val(CSTRING$(XXZ$, 4, 12, 2, 1))
                  VFX#(K2%) = VALO#
               End If
               If TT%(K2%) > K2% Then
                  TT1% = TT%(K2%) Mod 100: SIG% = 1
                  If TT%(K2%) > 100 Then SIG% = (-1)
                  SUPAR#(TT1%) = SUPAR#(TT1%) + VALO# * SIG%
               End If
               '
            End If
        Next K2%
        '
        For K2% = 1 To NCL%
          If TVL%(CFT%(K2%)) < 5 Or TVL%(CFT%(K2%)) = 8 Then
            If TVL%(CFT%(K2%)) <> 8 Then VFX#(K2%) = 0
            LOTOT% = ACL%(K2%)
            LODEC% = LDL%(CFT%(K2%))
            TIVA% = TVL%(CFT%(K2%))
            '
            If TT%(K2%) = 255 Then
                 VFX#(K2%) = SUPAR#(K2%) ' : GoTo 14660
                 VFS$(K2%) = NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
               Else
                 If TT%(K2%) > 0 And TT%(K2%) < K2% Then
                    VFX#(K2%) = 0
                    'YYY$ = REGLEIDO$(ARCHITRA$, K1& + 1)
                    YYY$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, K1& + 1)
                    If Mid$(YYY$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) <> Mid$(XXX$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) Then
                       If KI% < TT%(K2%) Then KI% = TT%(K2%)
                       VFX#(K2%) = SUPAR#(K2%)
                       SUPAR#(K2%) = 0
                       If K2% > KF% Then KF% = K2%
                    End If
                    'XXX$ = REGLEIDO$(ARCHITRA$, K1&)
                    XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, K1&)
                    VFS$(K2%) = NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
                    If TIVA% = 8 Then
                       TIETO = VFX#(K2%)
                       JUSTI% = JT%(K2%)
                       VFS$(K2%) = CSTRING$(MKS$(TIETO), TIVA%, LOTOT%, LODEC%, JUSTI%)
                       If TIETO < 0.01 Then VFS$(K2%) = Space$(LOTOT%)
                    End If
                End If
            End If
            '
            If TT%(K2%) > 0 Then
               For K4% = 1 To 2
                  TIVA% = TVL%(CFT%(K2%))
                  If TIVA% <> 8 Then
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + CDbl(Val(VFS$(K2%)))
                     Else
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + VFX#(K2%)
                  End If
                  If TT%(K2%) = 255 Then
                      SUMA#(K2%, 1) = 0
                      SUMA#(K2%, 2) = CDbl(Val(VFS$(K2%)))
                  End If
               Next K4%
            End If
            '
          End If
        Next K2%
        '
4       RENGLO$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           '
           TPAR$ = VFS$(K2%)
           If LK%(K2%) = 0 Then
               If VFS$(K2%) = VFSA$(K2%) Then
                   TPAR$ = Space$(Len(VFS$(K2%)))
                 Else
                   For KK22% = K2% + 1 To NCL%
                     VFSA$(KK22%) = ""
                   Next KK22%
               End If
           End If
           VFSA$(K2%) = VFS$(K2%)
           If JT%(K2%) = 1 Then
               TPAR$ = Space$(MG%(K2%)) + TPAR$
             ElseIf JT%(K2%) = 2 Then
               TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RENGLO$ = RENGLO$ + TPAR$ + CXC$
        Next K2%
        '
        '
5       RGXX1$ = RENGLO$
        RENGLO$ = REPLACAR$(RGXX1$, "@", "~")
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + RENGLO$
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
           Else
             Print #ASAL%, RENGLO$
        End If
        '
        If KI% > 0 Then
            RESEP$ = CXC$: If KI% = 1 Then RESEP$ = CCC$(5)
            
            For K2% = 1 To NCL%
               If K2% < KI% Or K2% > KF% Then
                    RESEP$ = RESEP$ + Space$(MG%(K2%) + Len(VFS$(K2%)))
                    If (K2% < KI% - 1 Or K2% > KF%) Then RESEP$ = RESEP$ + CXC$ Else RESEP$ = RESEP$ + CCC$(5)
                  Else
                    RESEP$ = RESEP$ + String$(MG%(K2%) + Len(VFS$(K2%)), "-") + CCC$(5)
                End If
                VFSA$(K2%) = ""
            Next K2%
            '
            If K1& >= FINX& Then
                RESEP$ = LII$(3)
            End If
            '
            If ARCHISAL$ = "" Then
                 Printer.Print MIZQ$ + RESEP$
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                 UIU% = 1
               Else
                 Print #ASAL%, RESEP$
                 UIU% = 2
            End If
        End If
        '
99  Next K1&
    '
    ULTIHOJ% = 1
    Call LITRANSPO
    If ARCHISAL$ = "" Then
        Call EJECT
        For K2% = 1 To NCL%
            VFSA$(K2%) = ""
        Next K2%
190     On Error GoTo 199
        Printer.EndDoc
        On Error GoTo 0
    End If
    '
    If ASAL% > 0 Then
        Close #ASAL%
        ASAL% = 0
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 190
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Sub LITRANSPO()
    '
    If TOTALIX% > 0 Then
        '
        RETRA$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           LOTOT% = ACL%(K2%)
           LODEC% = LDL%(CFT%(K2%))
           '
           TPAR$ = Space$(LOTOT%)
           If K2% = POSITOT% Then
                TPAR$ = AJUSTI$("Transporte", LOTOT%)
                If ULTIHOJ% = 1 Then
                   TPAR$ = AJUSTI$("T O T A L   G E N E R A L", LOTOT%)
                End If
              ElseIf TT%(K2%) > 0 Then
                TPAR$ = NUMSTRI$(SUMA#(K2%, 2), LOTOT%, LODEC%, 0)
                If TVL%(CFT%(K2%)) = 8 Then
                   TIETO = SUMA#(K2%, 2)
                   TPAR$ = CSTRING$(MKS$(TIETO), 8, LOTOT%, LODEC%, JT%(K2%))
                End If
           End If
           '
           If JT%(K2%) = 1 Then
                TPAR$ = Space$(MG%(K2%)) + TPAR$
              ElseIf JT%(K2%) = 2 Then
                TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RETRA$ = RETRA$ + TPAR$ + CXC$
           '
        Next K2%
        '
        If ULTIHOJ% = 0 Then
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + LII$(0)  ' renglon en blanco
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + RETRA$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, LII$(0)    ' renglon en blanco
                Print #ASAL%, RETRA$
                Print #ASAL%, LII$(0)
             End If
        End If
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + LII$(2)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        End If
        '
        RESUB$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           LOTOT% = ACL%(K2%)
           LODEC% = LDL%(CFT%(K2%))
           '
           TPAR$ = Space$(LOTOT%)
           If K2% = POSITOT% Then
                TPAR$ = AJUSTI$("Parciales", LOTOT%)
              ElseIf TT%(K2%) > 0 Then
                TPAR$ = NUMSTRI$(SUMA#(K2%, 1), LOTOT%, LODEC%, 0)
                If TVL%(CFT%(K2%)) = 8 Then
                   TIETO = SUMA#(K2%, 1)
                   TPAR$ = CSTRING$(MKS$(TIETO), 8, LOTOT%, LODEC%, JT%(K2%))
                End If
           End If
           SUMA#(K2%, 1) = 0
           '
           If JT%(K2%) = 1 Then
                TPAR$ = Space$(MG%(K2%)) + TPAR$
              ElseIf JT%(K2%) = 2 Then
                TPAR$ = TPAR$ + Space$(MG%(K2%))
           End If
           '
           RESUB$ = RESUB$ + TPAR$ + CXC$
           '
        Next K2%
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + RESUB$
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
             Printer.Print MIZQ$ + LII$(3)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
        End If
        '
        If ULTIHOJ% = 1 Then
            '
            Dim HDRX As String * 256
            Dim HDRZ$(32)
            CASBT% = 0
            If EXISTE%(LUCOM$ + ARCHITRA$ + ".SBT") > 0 Then
               NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".SBT", 0, 256)
               For u& = 1 To LOF(NXX%) / 256
                 If u& <= 32 Then
                   Get #NXX%, u&, HDRX$
                   HDRZ$(u&) = Left$(HDRX$, ATOT%)
                   CASBT% = u&
                 End If
               Next u&
               Close #NXX%
            End If
            '
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + " "
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(1)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)  ' renglon en blanco
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                If CASBT% > 0 Then
                  For SBT% = 1 To CASBT%
                    Printer.Print MIZQ$ + HDRZ$(SBT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Next SBT%
                  Printer.Print MIZQ$ + LII$(0)
                End If
                Printer.Print MIZQ$ + RETRA$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(0)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Printer.Print MIZQ$ + LII$(3)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, " "
                Print #ASAL%, LII$(1)
                Print #ASAL%, LII$(0)    ' renglon en blanco
                If CASBT% > 0 Then
                  For SBT% = 1 To CASBT%
                    Print #ASAL%, HDRZ$(SBT%)
                  Next SBT%
                  Print #ASAL%, LII$(0)
                End If
                Print #ASAL%, RETRA$
                Print #ASAL%, LII$(0)
                Print #ASAL%, LII$(3)
            End If
        End If
        '
      Else
        '
        If ARCHISAL$ = "" Then
             Printer.Print MIZQ$ + LII$(3)
             If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
           Else
             Print #ASAL%, LII$(3)
        End If
        '
    End If
    '
End Sub
'
'
Sub ARMAENCA()
    '
    If SALEXCEL% = 1 Then
      If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDP") > 0 Then
        HDRX$ = LOADFILE$(LUCOM$ + ARCHITRA$ + ".HDP")
        ENCA$ = AJUSTI(TRIM$(HDRX$), ATOT%)
        '
        For KCL% = 1 To NCL%
          ECLX1$ = TRIM$(Left$(ENCA$, ACL%(KCL%)))
          If Left$(ECL$(KCL%), 1) = "<" Or JT%(KCL%) = 1 Then
              ECL$(KCL%) = "<" + ECLX1$
            ElseIf Left$(ECL$(KCL%), 1) = ">" Or JT%(KCL%) = 2 Then
              ECL$(KCL%) = ">" + ECLX1$
            Else
              ECL$(KCL%) = ECLX1$
          End If
          ENCA$ = Mid$(ENCA$, 1 + ACL%(KCL%) + MG%(KCL%))
        Next KCL%
      End If
    End If
    '
    ENCA$ = Space$(OFS%) + CXC$
    '
    For KCL% = 1 To NCL%
        If Left$(ECL$(KCL%), 1) = "<" Then
              ECL$(KCL%) = Left$(Space$(MG%(KCL%)) + Mid$(ECL$(KCL%), 2) + Space$(ACL%(KCL%)), ACL%(KCL%) + MG%(KCL%))
          ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ECL$(KCL%) = Right$(Space$(ACL%(KCL%)) + Mid$(ECL$(KCL%), 2) + Space$(MG%(KCL%)), ACL%(KCL%) + MG%(KCL%))
          Else
              While Len(ECL$(KCL%)) < ACL%(KCL%)
                  ECL$(KCL%) = " " + ECL$(KCL%) + " "
              Wend
              ECL$(KCL%) = Space$(MG%(KCL%)) + Left$(ECL$(KCL%), ACL%(KCL%))
        End If
        '
        If Right$(ECL$(KCL%), 2) = ". " Then
            ECL$(KCL%) = " " + Left$(ECL$(KCL%), Len(ECL$(KCL%)) - 1)
        End If
        '
        ENCA$ = ENCA$ + ECL$(KCL%) + CXC$
        '
    Next KCL%
    '
End Sub
'
Sub LISEXCEL()
     '
     Screen.MousePointer = 11
     '
     Print #ASAL%, ARCHISAL$
     '
     For K2% = 1 To NCL%
       '
       TPAR$ = ECL$(K2%)
       If JT%(K2%) = 1 Then
           TPAR$ = Space$(MG%(K2%)) + TPAR$
         ElseIf JT%(K2%) = 2 Then
           TPAR$ = TPAR$ + Space$(MG%(K2%))
       End If
       '
       Print #ASAL%, "#T#" + TPAR$
     Next K2%
     Print #ASAL%, "@NEWLINE@"
     '
     For K2% = 1 To NCL%
       '
       TPAR$ = String$(Len(ECL$(K2%)), "-")
       If JT%(K2%) = 1 Then
           TPAR$ = Space$(MG%(K2%)) + TPAR$
         ElseIf JT%(K2%) = 2 Then
           TPAR$ = TPAR$ + Space$(MG%(K2%))
       End If
       '
       Print #ASAL%, "#T#" + TPAR$
     Next K2%
     Print #ASAL%, "@NEWLINE@"
     Print #ASAL%, "@NEWLINE@"
     '
     FINX& = ULTREG&(ARCHITRA$)
     '
     For K1& = HDR& + 1 To FINX&
        Call TRACE(20, K1&, FINX&)
        XXX$ = REGLEIDO$(ARCHITRA$, K1&)
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "=") Then
            GoTo 19
        End If
        '
        If XXX$ = REGBLAN$(ARCHITRA$) Then
            GoTo 19
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "-") Then
            GoTo 19
        End If
        '
        If XXX$ = String$(RECLEN%(ARCHITRA$), "*") Then
            GoTo 19
        End If
        '
        KI% = 0: KF% = 0
        For K2% = 1 To NCL%
            If PE%(K2%) < 3 Then
               If CFT%(K2%) = 1 Then
                     XXZ$ = MKS$(K1& + HDR&)
                     If NFT% > 0 Then XXZ$ = RECNO$
                   Else
                     XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
               End If
               '
               TIVA% = TVL%(CFT%(K2%))
               If TIVA% = 6 Then
                 If Left$(XXZ$, 1) = "@" Then
                   For K3% = 1 To NCL%
                     VFS$(K3%) = Space$(ACL%(K3%))
                   Next K3%
                   VFS$(K2%) = "#T#" + Mid$(XXZ$, 2)
                   GoTo 4
                 End If
               End If
               '
               LOTOT% = ACL%(K2%)
               LODEC% = LDL%(CFT%(K2%))
               JUSTI% = JT%(K2%)
               '
               If TIVA% = 5 Then
                 VFS$(K2%) = "#F#" + FETEXCOR$(CVI(XXZ$))
                 GoTo 3
               End If
               '
               VFS$(K2%) = CSTRING$(XXZ$, TIVA%, LOTOT%, LODEC%, JUSTI%)
               If TIVA% < 5 Then
                    VFS$(K2%) = "#N#" + VFS$(K2%)
                  ElseIf TIVA% = 6 Or TIVA% = 7 Or TIVA% = 8 Then
                    VFS$(K2%) = "#T#" + VFS$(K2%)
               End If
               '
               If CE%(K2%) > 0 Then
                  VFS$(CE%(K2%)) = "#T#" + AJUSTI$(ECOARCH$(AE$(K2%), XXZ$), ACL%(CE%(K2%)))
               End If
               '
               VALO# = 0: If TIVA% <= 4 Then VALO# = Val(Mid$(VFS$(K2%), 4))
               If TIVA% = 8 Then
                  VALO# = Val(CSTRING$(XXZ$, 4, 12, 2, 1))
                  VFX#(K2%) = VALO#
               End If
               If TT%(K2%) > K2% Then
                  TT1% = TT%(K2%) Mod 100: SIG% = 1
                  If TT%(K2%) > 100 Then SIG% = (-1)
                  SUPAR#(TT1%) = SUPAR#(TT1%) + VALO# * SIG%
               End If
               '
            End If
3       Next K2%
        '
        For K2% = 1 To NCL%
          If TVL%(CFT%(K2%)) < 5 Or TVL%(CFT%(K2%)) = 8 Then
            If TVL%(CFT%(K2%)) <> 8 Then VFX#(K2%) = 0
            LOTOT% = ACL%(K2%)
            LODEC% = LDL%(CFT%(K2%))
            TIVA% = TVL%(CFT%(K2%))
            '
            If TT%(K2%) = 255 Then
                 VFX#(K2%) = SUPAR#(K2%) ' : GoTo 14660
                 VFS$(K2%) = "#N#" + NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
               Else
                 If TT%(K2%) > 0 And TT%(K2%) < K2% Then
                    VFX#(K2%) = 0
                    YYY$ = REGLEIDO$(ARCHITRA$, K1& + 1)
                    If Mid$(YYY$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) <> Mid$(XXX$, PR%(CFT%(TT%(K2%))), LF%(CFT%(TT%(K2%)))) Then
                       If KI% < TT%(K2%) Then KI% = TT%(K2%)
                       VFX#(K2%) = SUPAR#(K2%)
                       SUPAR#(K2%) = 0
                       If K2% > KF% Then KF% = K2%
                    End If
                    XXX$ = REGLEIDO$(ARCHITRA$, K1&)
                    VFS$(K2%) = "#N#" + NUMSTRI$(VFX#(K2%), LOTOT%, LODEC%, 0)
                    If TIVA% = 8 Then
                       TIETO = VFX#(K2%)
                       JUSTI% = JT%(K2%)
                       VFS$(K2%) = "#T#" + CSTRING$(MKS$(TIETO), TIVA%, LOTOT%, LODEC%, JUSTI%)
                       If TIETO < 0.01 Then VFS$(K2%) = Space$(LOTOT%)
                    End If
                End If
            End If
            '
            If TT%(K2%) > 0 Then
               For K4% = 1 To 2
                  TIVA% = TVL%(CFT%(K2%))
                  If TIVA% <> 8 Then
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + CDbl(Val(VFS$(K2%)))
                     Else
                       SUMA#(K2%, K4%) = SUMA#(K2%, K4%) + VFX#(K2%)
                  End If
                  If TT%(K2%) = 255 Then
                      SUMA#(K2%, 1) = 0
                      SUMA#(K2%, 2) = CDbl(Val(VFS$(K2%)))
                  End If
               Next K4%
            End If
            '
          End If
        Next K2%
        '
4       RENGLO$ = Space$(OFS%) + CXC$
        For K2% = 1 To NCL%
           TPAR$ = VFS$(K2%)
           Print #ASAL%, TPAR$
        Next K2%
        '
        '
19      Print #ASAL%, "@NEWLINE@"
        '
99  Next K1&
    '
    Close #ASAL%
    ASAL% = 0
    Shell "FLEXCEL.EXE", vbNormalFocus
    PROLLAMA$ = "Generación de Planilla Excel"
    '
    Screen.MousePointer = 1
    '
End Sub

Sub SETPRINTER()
    '
    Dim FUENTESEL
    Dim FUENPRUE
    Dim NADAFUEN
    '
    If IDL% = 32767 Then
        For u% = 1 To 17
           CCI$(u%) = ""
        Next u%
        Exit Sub
    End If
    '
    SRY$ = String$(255, "*")
    GU$ = String$(255, "-")
    CXC$ = "": CC$ = ""
       If BLP% = 1 Then CXC$ = "¦": CC$ = "="
    For I% = 1 To 9
         CCC$(I%) = "": If BLP% = 1 Then CCC$(I%) = "+"
    Next I%
    '
    For I% = 1 To 3
       LII$(I%) = Space$(OFS%) + CCC$(I%)
       For J% = 1 To NCL%
          LII$(I%) = LII$(I%) + String$(LC%(J%) + MG%(J%), "-")
          If J% < NCL% Then
              LII$(I%) = LII$(I%) + CCC$(I% + 1)
            Else
              LII$(I%) = LII$(I%) + CCC$(I% + 2)
          End If
       Next J%
    Next I%
    '
    LII$(0) = Space$(OFS%) + CXC$
    LII$(4) = Space$(OFS%) + CC$
    LII$(5) = Space$(OFS%) + CC$
    '
    For J% = 1 To NCL%
       LII$(0) = LII$(0) + Space$(LC%(J%) + MG%(J%)) + CXC$
       LII$(4) = LII$(4) + String$(LC%(J%) + MG%(J%), "=") + CC$
       LII$(5) = LII$(5) + String$(LC%(J%) + MG%(J%), "*") + CC$
    Next J%
    '
    If TRIM$(ARCHISAL$) <> "" Then
        MIZQ$ = ""
        PRINTERSET = 1
    End If
    '
    If PRINTERSET% = 0 Then
        '
        If Printer.CurrentX > 0 Then
            Printer.NewPage
        End If
        '
10      NOMIMPRE$ = TRIM$(Printer.DeviceName)
        If NOMIMPRE$ = "" Then
            Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
            Call FINAL("")
        End If
        '
        YAESTA% = 0
        CAFUFIJ% = 0
        NX% = FILEOPEN("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
        FIN& = LOF(NX%) / 128
        For IPRIN& = 1 To FIN&
            Get #NX%, IPRIN&, PPZZ$
            PPXX$ = PPZZ$
            If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
                YAESTA% = 1
                If UCase$(Mid$(PPXX$, 125, 2)) = "SI" Then
                    CAFUFIJ% = CAFUFIJ% + 1
                    FUFIJA$(CAFUFIJ%) = PPZZ$
                End If
            End If
        Next IPRIN&
        '
        Close #NX%
        NOMIMP$ = REPLACAR$(NOMIMPRE$, "\", "/")
        If YAESTA% < 1 Then
            OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
        If CAFUFIJ% < 1 Then
            OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
            If OPX% <> 1 Then Call FINAL("")
            Call SETUPRINTER
            GoTo 10
        End If
        '
        Printer.ScaleMode = 1
        ANCHOJA = 56.7 * (ANCHOPAP - MARGENI)
        If ORIENTA% = 2 Then ANCHOJA = 56.7 * (LARGOPAP - MARGENI)
        'ANCHOMAX = Int(56.7 * CVS(Mid$(FUFIJA$(1), 49, 4)))
        'If ANCHOMAX < ANCHOJA Then ANCHOJA = ANCHOMAX
        '
        ATEXMAX = 0
        TAFUESEL = 0
        '
        TALEBASI = 24
        '
100     TEPRUEBA$ = String$(ATOT%, "*")
        Printer.FontBold = False
        Printer.FontItalic = False
        For I% = 1 To CAFUFIJ%
            '
            Printer.FontName = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
            Printer.FontBold = False
            Printer.FontItalic = False
            '
            FUENPRUE = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
            If Printer.FontName <> FUENPRUE Then
'MsgBox Printer.FontName & TRIM$(Mid$(FUFIJA$(I%), 65, 48))
                GoTo 109   ' modificado 20/11/01 para que busque la fuente
                           ' disponible
                Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
                Call CIERRARCH("*.*")
                Close
                Call FILEKILL("C:\FLEXOFT\PRINTERS.DRV")
                Call FINAL("")
            End If
            '
            FUENTEMIN = CVS(Mid$(FUFIJA$(I%), 113, 4))
            FUENTEMAX = CVS(Mid$(FUFIJA$(I%), 117, 4))
            '
            TALEBA = FUENTEMAX + 0.4
            While TALEBA > FUENTEMIN
               TALEBA = TALEBA - 0.4
               Printer.FontSize = TALEBA
               If Printer.FontSize > TALEBA Then GoTo 105
               ATEXTO = Printer.TextWidth(TEPRUEBA$)
               If ATEXTO <= ANCHOJA Then
                  If ATEXTO > ATEXMAX Then
                    FUENTESEL = FUENPRUE    ' Printer.FontName
                    TAFUESEL = Printer.FontSize ' TALEBASI
                    ATEXMAX = ATEXTO
                  End If
                  GoTo 109
               End If
               TALEBA = Printer.FontSize
105         Wend
        '
109     Next I%
        '
        If FUENTESEL = NADAFUEN Or TAFUESEL < 1 Then
            Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
            Call FINAL("")
        End If
        '
        Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        Text5.Text = FUENTESEL
        Text5.Refresh
        Text7.Text = CSTRING$(MKS$(TAFUESEL), 3, 6, 1, 1)
        Text7.Refresh
       '
        If FUENTESEL = "" Or TAFUESEL < 0.4 Then
            Call MENSERR(24, "No puede Establecerse\el Formato de Impresion.\Consulte al Soporte de Sistemas.")
            Call FINAL("")
        End If
        '
110     Printer.FontName = FUENTESEL
        Printer.FontSize = TAFUESEL
        LARGOJA = Val(Printer.ScaleHeight)
        'ANCHOJA = Printer.ScaleWidth
        TEPRUEBB$ = TEPRUEBA$ ' : If LX810 = 1 Then TEPRUEBB$ = TEPRUEBA$ + "*****"
        ANCHOJB = Val(Printer.TextWidth(TEPRUEBB$))
        TALEBASI = Printer.FontSize
        ALTUREN = Printer.TextHeight(TEPRUEBA$)
        Printer.FontBold = False
        Printer.FontItalic = False
        MIZMAX = ANCHOJA - ANCHOJB
        If MARGENI > 0 Then
            MIZMAX = 56.7 * MARGENI
        End If
        MIZQ$ = ""
        While Printer.TextWidth(MIZQ$ + "*") <= MIZMAX
           MIZQ$ = MIZQ$ + " "
        Wend
        MIZQ$ = Mid$(MIZQ$, 2)
        PRINTERSET% = 1
        '
    End If
    '
End Sub
'
'
Sub BLANHEADERS()
   For u& = 0 To 16
     LEYE$(u&) = ""
   Next u&
End Sub
'
Sub CARHEADERS()
     '
     For u& = 1 To HDR&
         SUBTI$(u&) = REGLEIDO$(ARCHITRA$, u&)
     Next u&
     '
     X1% = 0
     ARCH$ = AJUSTI$(UCase$(ARCHITRA$), 8)
     For X% = 1 To 10
        If LTrim$(RTrim$(LEYE$(X%))) <> "" Then
             X1% = X1% + 1
             LEYE$(X1%) = LEYE$(X%)
        End If
     Next X%
     '
     For u& = 1 To ULTREG&("HEADERS")
        XXX$ = REGLEIDO$("HEADERS", u&)
        If UCase$(Left$(XXX$, 8)) = ARCH$ Then
            For X4% = 1 To X1%
                 If LEYE$(X4%) = Mid$(XXX$, 9) Then GoTo 1024
            Next X4%
            X1% = X1% + 1
            If X1% < 10 Then
                 LEYE$(X1%) = Mid$(XXX$, 9)
            End If
        End If
1024 Next u&
     '
     Call CIERRARCH("HEADERS")
     For X3% = X1% + 1 To 10
            LEYE$(X3%) = ""
     Next X3%
     '
End Sub
'
Sub INITOTA()
    '
    For K3% = 1 To NCL%
        For K4% = 1 To 2
           SUMA#(K3%, K4%) = 0
        Next K4%
    Next K3%
    '
    NUHOJ% = 0
    PPXX% = InStr(ATRIBU$, "PRIHOJA=")
    If PPXX% > 0 Then
       NUHOJ% = Val(Mid$(ATRIBU$, PPXX% + 8)) - 1
    End If
    '
End Sub
'
Sub ENCAHOJA()
    '
    Dim HDRX As String * 256
    Dim HDRZ$(32)
    Dim ENCATAB$(32)
    '
    If InStr(ARCHISAL$, ".XLS") > 0 Then
      Exit Sub
    End If
    '
    NUHOJ% = NUHOJ% + 1
    NH# = CDbl(NUHOJ%)
    '
    For K3% = 1 To NCL%
         VDAA%(K3%) = 0
         VDAB#(K3%) = 0
         VDA$(K3%) = ""
    Next K3%
    '
    If NUHOJ% >= PRIHOJ% Then
           ' Call MEMBRETE
         If ARCHISAL$ = "" Then
           If TNC% < 1 Then
              Printer.CurrentY = MARGENS * 56.7
              CURRY = Printer.CurrentY
              For RI% = 1 To RGI&
                 Printer.Print , " "
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Next RI%
           End If
         End If
         '
         If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
         '
         ESPAMED% = ATOT% - Len(S1$) - 35
         S2$ = Left$(SRY$, Len(S1$))
         '
         If TNC% > 0 Then GoTo 10
         '
         HOII% = HOY(HOS$)
         TEFE$ = "Fecha: " + HOS$: If SUPRIFE% = 1 Then TEFE$ = "               "
         SUFE$ = "******": If SUPRIFE% = 1 Then SUFE$ = "      "
         HDRZ$(1) = S1$ + Space$(ESPAMED%) + "   " + TEFE$ + "    Hoja No:#####"
         HDRZ$(2) = S2$ + Space$(ESPAMED%) + "   " + SUFE$ + "             ********"
         CAENCA% = 2
         '
         If Left$(S1$, 1) = "%" Then
           HDRZ$(1) = Space$(10)
           HDRZ$(2) = Space$(10)
         End If
         '
         CAENCOL% = 0
         If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDR") > 0 Then
            NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".HDR", 0, 256)
            For u& = 1 To LOF(NXX%) / 256
               If u& <= 32 Then
                  Get #NXX%, u&, HDRX$
                  ENCATAB$(u&) = Left$(HDRX$, ATOT%)
                  CAENCOL% = u&
               End If
            Next u&
            Close #NXX%
         End If
         '
         If NUHOJ% <= PRIHOJ% Or NUHOJ% = 1 Then
            If EXISTE%(LUCOM$ + ARCHITRA$ + ".HDP") > 0 Then
               'CAENCA% = CAENCA% + 1: HDRZ$(CAENCA%) = Space$(ATOT%)
               '
               NXX% = FILEOPEN%(LUCOM$ + ARCHITRA$ + ".HDP", 0, 256)
               'For U& = 1 To 1 + Int((LOF(NXX%) - 1) / 256)
               Get #NXX%, 1, HDRX$
               ENCA$ = AJUSTI(TRIM$(HDRX$), ATOT%)
               '   CAENCA% = CAENCA% + 1
               '   HDRZ$(CAENCA%) = Left$(HDRX$, ATOT%)
               'Next U&
               Close #NXX%
            End If
         End If
         '
         PPXX% = InStr(HDRZ$(1), "#####")
         If PPXX% > 0 Then
            Call REPLA(HDRZ$(1), NUMSTRI$(NH#, 5, 0, 0), PPXX%, 5)
         End If
         '
         For u& = 1 To CAENCA%
            If ARCHISAL$ = "" Then
                 Printer.Print MIZQ$ + Space$(OFS%) + HDRZ$(u&)
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
               Else
                 Print #ASAL%, Space$(OFS%);
                 Print #ASAL%, HDRZ$(u&)
            End If
         Next u&
         '
10       For RI% = 1 To 10
              LEYEN$ = TRIM$(LEYE$(RI%))
              If LEYEN$ <> "" Then
                 If Left$(LEYEN$, 1) <> ":" Then
                    PRI% = InStr(LEYEN$, ":")
                    If PRI% = 0 Then PRI% = Len(LEYEN$) + 1
                    SUBRA$ = String$(PRI%, "*")
                    While Len(SUBRA$) < Len(S1$)
                       LEYEN$ = " " + LEYEN$
                       SUBRA$ = " " + SUBRA$
                    Wend
                    '
                    If ARCHISAL$ = "" Then
                        Printer.Print " "
                        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                      Else
                        Print #ASAL%, " "
                    End If
                    '
                    LEYF$ = ""
                    If Len(LEYEN$) > ATOT% Then
                       PPXY% = InStr(UCase$(LEYEN$), "HASTA:")
                       If PPXY% > 0 Then
                          LEYF$ = Mid$(LEYEN$, PPXY%)
                          LEYEN$ = Left$(LEYEN$, PPXY% - 1)
                       End If
                    End If
                    '
                    If ARCHISAL$ = "" Then
                         Printer.Print MIZQ$ + Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                         If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                         Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBRA$ + Space$(ATOT%), ATOT%)
                         If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                       Else
                         Print #ASAL%, Space$(OFS%) + Left$(LEYEN$ + Space$(ATOT%), ATOT%)
                         Print #ASAL%, Space$(OFS%) + Left$(SUBRA$ + Space$(ATOT%), ATOT%)
                    End If
                    '
                    If Len(LEYF$) > 0 Then
                       If ARCHISAL$ = "" Then
                           Printer.Print MIZQ$ + Space$(OFS%) + Left$(Space$(1 + Len(SUBRA$)) + LEYF$, ATOT%)
                           If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                         Else
                           Print #ASAL%, Space$(OFS%) + Left$(Space$(1 + Len(SUBRA$)) + LEYF$, ATOT%)
                       End If
                    End If
                    '
                 End If
              End If
              '
         Next RI%
         '
         If HDR& > 0 Then
            Printer.Print " "
            If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            For RI% = 1 To HDR&
               If ARCHISAL$ = "" Then
                    Printer.Print MIZQ$ + Space$(OFS%) + Left$(SUBTI$(RI%), ATOT%)
                    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                  Else
                    Print #ASAL%, Space$(OFS%) + Left$(SUBTI$(RI%), ATOT%)
               End If
            Next RI%
         End If
         '
30       For RI% = 1 To RFT%
              If ARCHISAL$ = "" Then
                  Printer.Print " "
                  If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                Else
                  Print #ASAL%, " "
              End If
         Next RI%
         '
         If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + LII$(1)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                If CAENCOL% < 1 Then
                     Printer.Print MIZQ$ + ENCA$
                     If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                   Else
                     For KKCOL% = 1 To CAENCOL%
                       Printer.Print MIZQ$ + ENCATAB$(KKCOL%)
                       If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
                     Next KKCOL%
                End If
                Printer.Print MIZQ$ + LII$(2)
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, LII$(1)
                If CAENCOL% < 1 Then
                     Print #ASAL%, ENCA$
                   Else
                     For KKCOL% = 1 To CAENCOL%
                       Print #ASAL%, ENCATAB$(KKCOL%)
                     Next KKCOL%
                End If
                Print #ASAL%, LII$(2)
         End If
         '
    End If
    '
    For K3% = 1 To NCL%
       SUMA#(K3%, 1) = 0
    Next K3%
    '
End Sub
'

Sub ENCAHOJA1(NCXX$)
    '
    NRENCA% = 0
    If ARCHISAL$ = "" Then
        If TNC% > 0 Then
            Printer.CurrentY = 56.7 * MARGENS
            CURRY = Printer.CurrentY
            For RI% = 1 To RGI&
                 Printer.Print , " "
                 If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
            Next RI%
        End If
    End If
    '
    For CLIJ% = 1 To TNC%
        TTXI$ = Mid$(NCXX$, PR%(CLNC%(CLIJ%)), LF%(CLNC%(CLIJ%)))
        If TVL%(CLNC%(CLIJ%)) = 6 Then
            If ARCHISAL$ = "" Then
                Printer.Print MIZQ$ + TTXI$
                If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
              Else
                Print #ASAL%, TTXI$
            End If
            NRENCA% = NRENCA% + 1
        End If
    Next CLIJ%
    '
    '
End Sub
'

Sub EJECT()
    '
10  On Error GoTo 99
    While Printer.CurrentY < Printer.ScaleHeight - 5 * ALTUREN
        Printer.Print " "
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Wend
    RNGL$ = MIZQ$ + Space$(OFS%) + String$(ATOT%, "-")
    ANCHOJB = Printer.TextWidth(RNGL$)
    If ULTIHOJ% > 0 Or TOTALIX% = 0 Then
        Printer.Print RNGL$
        If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    End If
    Printer.FontSize = 6
    '
    TXT$ = "Sistema FLEXOFT s/n " + SERINUM$
    TXT$ = TXT$ + " licenciado a " + EMPRELI$
    'If UCase$(FECHALI$) <> "S/LIMITE" Then
    '    TXT$ = TXT$ + " hasta " + FECHALI$
    'End If
    TXT$ = TXT$ + " - Derechos Reservados L.11723"
    TXT$ = MIZQ$ + Space$(OFS%) + TXT$
    '
    While Printer.TextWidth(TXT$) < ANCHOJB
       TXT$ = " " + TXT$ + " "
    Wend
    Printer.Print TXT$
    If INTERLIX > 0 Then CURRY = CURRY + INTERLIX: Printer.CurrentY = CURRY
    Printer.FontSize = TAFUESEL
    On Error GoTo 0
    Exit Sub
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
100 MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume 10
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'
Sub SETUPRINTER()
    '
    ' Hide
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    Call BLANARCH("!PRINTERS")
    NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    NX% = FILEOPEN("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
    For IPRIN& = 1 To LOF(NX%) / 128
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
            Call REGAPP("!PRINTERS", PPXX$)
            YAESTA% = 1
            GoTo 10
        End If
    Next IPRIN&
    '
10  Close #NX%
    If YAESTA% = 1 Then
        PPXX$ = Left$(PPXX$, 52)
        'GoTo 50
    End If
    '
    Printer.ScaleMode = 1
    ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    CAFUFIJ% = 0
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        A1 = 0: A2 = 0
        On Error Resume Next
        A1 = Printer.TextWidth(String$(80, "l"))
        A2 = Printer.TextWidth(String$(80, "M"))
        On Error GoTo 0
        If A1 > 0 And A2 > 0 Then
          If A1 = A2 Then
            If CAFUFIJ% < 128 Then
              CAFUFIJ% = CAFUFIJ% + 1
              FUFIJA$(CAFUFIJ%) = Printer.Fonts(I% - 1)
            End If
          End If
        End If
    Next I%
    '
    If CAFUFIJ% < 1 Then
       Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(I%)
        If Printer.FontName <> FUFIJA$(I%) Then GoTo 19
        '
        Printer.FontSize = 24
        FUENTEMAX = Printer.FontSize
        Printer.FontSize = 2
        FUENTEMIN = Printer.FontSize
        '
14      If FUENTEMIN <= FUENTEMAX Then
            XX$ = REGBLAN$("!PRINTERS")
            Call REPLA(XX$, NOMIMPRE$, 1, 48)
            Call REPLA(XX$, MKS$(ANCHOJMM), 49, 4)
            Call REPLA(XX$, FUFIJA$(I%), 65, 48)
            Call REPLA(XX$, MKS$(FUENTEMIN), 113, 4)
            Call REPLA(XX$, MKS$(FUENTEMAX), 117, 4)
            Call REPLA(XX$, "SI", 125, 2)
            Call REGAPP("!PRINTERS", XX$)
        End If
        '
19  Next I%
    Call CIERRARCH("!PRINTERS")
    '
    PPXX$ = AJUSTI$(NOMIMPRE$, 48) + MKS$(ANCHOJMM) + String$(12, 0)
    '
50  PPYY$ = OBJPLA$("DATIMPRE", PPXX$)
    If PPYY$ = "" Then
        Call CIERRARCH("PRINTERS")
        GoTo 99
    End If
    '
    ANCHOJMM = CVS(Mid$(PPYY$, 49, 8))
    If ANCHOJMM > Int(Printer.ScaleWidth / 56.7) Or ANCHOJMM < 48 Then
        PPXX$ = PPYY$
        GoTo 50
    End If
    '
    VTB% = VENTABU%("PRINSET/" + AJUSTI$(NOMIMPRE$, 48))
    If VTB% >= 0 Then
        NX% = FILEOPEN%("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
        NOMI$ = AJUSTI$(NOMIMPRE$, 48)
        JJ& = 0
        For IPRIN& = 1 To LOF(NX%) / 128
            Get #NX%, IPRIN&, PPZZ$
            If Left$(PPZZ$, 48) <> NOMI$ Then
               JJ& = JJ& + 1
               Put #NX%, JJ&, PPZZ$
            End If
        Next IPRIN&
        For IPRIN& = 1 To ULTREG&("!PRINTERS")
            PPXX$ = REGLEIDO$("!PRINTERS", IPRIN&)
            Call REPLA(PPXX$, NOMI$, 1, 48)
            Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
            If UCase$(Mid$(PPXX$, 125, 2)) <> "SI" Then
                Call REPLA(PPXX$, "  ", 125, 2)
            End If
            PPZZ$ = PPXX$
            JJ& = JJ& + 1
            Put #NX%, JJ&, PPZZ$
        Next IPRIN&
        For IPRIN& = JJ& + 1 To LOF(NX%) / 128
            PPZZ$ = String$(128, 0)
            Put #NX%, IPRIN&, PPZZ$
        Next IPRIN&
        Close #NX%
    End If
    Call CIERRARCH("!PRINTERS")
    '
99  FORMA$ = FORMALI$
    If TRIM$(FORMA$) <> "" Then
        Call CARDATLIS
        Call CARDATIMPRE
      Else
        Call COMUNI("Para Efectivizar la nueva Configuración\Debe Volver a pedir el Listado.")
        Call FINAL("")
    End If
    '
End Sub
'

