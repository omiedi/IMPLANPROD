VERSION 5.00
Begin VB.Form WLITAB04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9720
   FillStyle       =   0  'Solid
   Icon            =   "WLITAB04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9720
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
      Height          =   600
      Left            =   8100
      Picture         =   "WLITAB04.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura Impresora"
      Top             =   5115
      Width           =   700
   End
   Begin VB.CommandButton BOTEXIT 
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
      Left            =   8860
      Picture         =   "WLITAB04.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona el Listador"
      Top             =   5115
      Width           =   700
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
      Left            =   7350
      Picture         =   "WLITAB04.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Emite Listado"
      Top             =   5115
      Width           =   700
   End
   Begin VB.Timer Timer2 
      Interval        =   350
      Left            =   420
      Top             =   0
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   7350
      TabIndex        =   49
      Top             =   4725
      Width           =   2220
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   945
         ScaleHeight     =   75
         ScaleWidth      =   1005
         TabIndex        =   50
         Top             =   210
         Width           =   1065
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Presentación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4830
      TabIndex        =   47
      Top             =   4725
      Width           =   2430
      Begin VB.Image Image1 
         Height          =   375
         Index           =   3
         Left            =   210
         Picture         =   "WLITAB04.frx":0BA0
         Stretch         =   -1  'True
         Top             =   400
         Width           =   500
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Vista Previa en Pantalla"
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
         Left            =   1050
         TabIndex        =   48
         Top             =   400
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   4830
      TabIndex        =   36
      Top             =   1635
      Width           =   4740
      Begin VB.CommandButton Command6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3990
         Picture         =   "WLITAB04.frx":28C92
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Actualizar Lista"
         Top             =   1575
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         Enabled         =   0   'False
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
         Left            =   3675
         TabIndex        =   58
         Top             =   1575
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   840
         TabIndex        =   57
         Text            =   " "
         Top             =   1575
         Width           =   2850
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
         Left            =   4330
         Picture         =   "WLITAB04.frx":28DDC
         TabIndex        =   52
         ToolTipText     =   "Click para Seleccionar Carpeta de Salida"
         Top             =   2520
         Width           =   175
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1250
         Left            =   2310
         TabIndex        =   41
         Top             =   200
         Width           =   2220
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
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
            Height          =   435
            Left            =   210
            TabIndex        =   44
            Top             =   740
            Width           =   1800
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   43
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   42
            Top             =   500
            Width           =   1695
         End
      End
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
         Left            =   2100
         TabIndex        =   37
         Text            =   " "
         Top             =   2170
         Width           =   2430
      End
      Begin VB.Label Label18 
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
         Left            =   240
         TabIndex        =   61
         ToolTipText     =   "Doble Click para Seleccionar Carpeta de Salida"
         Top             =   1920
         Visible         =   0   'False
         Width           =   4110
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Envio x Mail"
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
         Left            =   210
         TabIndex        =   60
         Top             =   1470
         Width           =   540
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Sheet"
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
         Left            =   840
         TabIndex        =   46
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Text File"
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
         Left            =   840
         TabIndex        =   45
         Top             =   430
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   2
         Left            =   210
         Picture         =   "WLITAB04.frx":290E6
         Stretch         =   -1  'True
         Top             =   900
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   210
         Picture         =   "WLITAB04.frx":2AA78
         Stretch         =   -1  'True
         Top             =   300
         Width           =   375
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5145
         Y1              =   2020
         Y2              =   2020
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
         Left            =   945
         TabIndex        =   40
         Top             =   2250
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
         TabIndex        =   39
         ToolTipText     =   "Doble Click para Seleccionar Carpeta de Salida"
         Top             =   2550
         Width           =   4110
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
         TabIndex        =   38
         Top             =   2310
         Width           =   1065
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   5145
         Y1              =   2010
         Y2              =   2010
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
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
      Left            =   3045
      TabIndex        =   22
      Top             =   6720
      Width           =   5160
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Top             =   315
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   2520
         TabIndex        =   24
         Top             =   735
         Width           =   2430
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Picture         =   "WLITAB04.frx":2AD82
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   540
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
      Left            =   5670
      Picture         =   "WLITAB04.frx":2B1FC
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guarda Configuración del Listado"
      Top             =   7560
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1410
      Left            =   240
      TabIndex        =   10
      Top             =   105
      Width           =   9360
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   4715
         X2              =   4715
         Y1              =   120
         Y2              =   1760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   4725
         X2              =   4725
         Y1              =   120
         Y2              =   1400
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
         Height          =   510
         Left            =   210
         TabIndex        =   30
         Top             =   735
         Width           =   4320
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
         Left            =   2940
         TabIndex        =   29
         Top             =   345
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
         Left            =   1785
         TabIndex        =   28
         Top             =   420
         Width           =   1065
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
         Height          =   510
         Left            =   4830
         TabIndex        =   27
         Top             =   735
         Width           =   4320
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
         Left            =   7560
         TabIndex        =   26
         Top             =   345
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
         Left            =   6405
         TabIndex        =   11
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4090
      Left            =   210
      TabIndex        =   0
      Top             =   1635
      Width           =   4530
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3570
         TabIndex        =   55
         Top             =   840
         Width           =   750
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   150
            TabIndex        =   56
            Top             =   525
            Width           =   450
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Renum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   2520
         TabIndex        =   53
         Top             =   1960
         Width           =   855
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   54
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tipo de Letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   31
         Top             =   840
         Width           =   3270
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   2520
            TabIndex        =   34
            Text            =   " "
            Top             =   270
            Width           =   540
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   285
            Left            =   315
            TabIndex        =   32
            Text            =   " "
            Top             =   630
            Width           =   2745
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño:"
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
            Left            =   1575
            TabIndex        =   35
            Top             =   315
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
            Left            =   -105
            TabIndex        =   33
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   735
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   3585
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Area de Impresión (en Milimetros)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   12
         Top             =   2730
         Width           =   4110
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   3
            Left            =   3150
            TabIndex        =   19
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   3150
            TabIndex        =   18
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   16
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   14
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
            Left            =   1890
            TabIndex        =   21
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
            Left            =   2205
            TabIndex        =   20
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
            TabIndex        =   15
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
            TabIndex        =   13
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   3465
         TabIndex        =   8
         Top             =   1960
         Width           =   855
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   210
            TabIndex        =   9
            Text            =   " "
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Orientación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   210
         TabIndex        =   5
         Top             =   1960
         Width           =   2220
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Paisaje"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1155
            TabIndex        =   7
            Top             =   360
            Width           =   960
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Retrato"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Image Image1 
         Height          =   345
         Index           =   0
         Left            =   210
         Picture         =   "WLITAB04.frx":2B506
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "WLITAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
Dim MODOSAL%
'
Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    WLITAB04.Hide
End Sub

Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For I% = 1 To Printers.Count
        If Printers(I% - 1).DeviceName = Combo1.TEXT Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(I% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.TEXT = Printer.DeviceName
    '
End Sub

Sub Command1_Click()
    '
    If MODOSAL% = 3 Then
      PREVIEW.List4.Clear
    End If
    '
    If MODOSAL% <> 0 And MODOSAL% <> 3 Then
       ASALIDA$ = TRIM$(Text2.TEXT)
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
       If Len(Label11) > 80 Then
          Call MENSERR(24, "Ruta de Salida Demasiado Larga.")
          Exit Sub
       End If
       ARCHISAL$ = TRIM$(Label11)
       If Len(ARCHISAL$) > 0 Then
         If Right$(ARCHISAL$, 1) <> "\" Then ARCHISAL$ = ARCHISAL$ + "\"
       End If
       ARCHISAL$ = ARCHISAL$ + ASALIDA$
       'If Option1(1).Value = True Then
       If MODOSAL% = 1 Then
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
       GoTo 20
    End If
    '
    If TRIM$(IMPREDETER$) <> TRIM$(Combo1.TEXT) Then
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
        NPPXX$ = TRIM$(Combo1.TEXT)
        For Each PRX9 In Printers
          If PRX9.DeviceName = NPPXX$ Then
            On Error Resume Next
            Set Printer = PRX9
            On Error GoTo 0
            GoTo 10
          End If
        Next PRX9
        '
10      If Printer.DeviceName <> Combo1.TEXT Then
          MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(Combo1.TEXT) + "'\   \"
          MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
          MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
          MENSA$ = MENSA$ + "Predeterminada para Windows."
          Combo1.TEXT = IMPREDETER$
          Call COMUNI(MENSA$)
          Combo1.SetFocus
          Exit Sub
        End If
    End If
    '
    Call GRADATIMPRE
    Call CARDATIMPRE
    '
20  SOBRESCRI% = 0
    If Option3(1) = True Then
        SOBRESCRI% = 1
    End If
    '
    On Error Resume Next
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
         If ARCHISAL$ = "" Then
           On Error GoTo 199
         End If
         If SALEXCEL% = 1 Then
           Call ARMAENCA
           Call LISEXCEL
           Call CIERRARCH("*.*")
            '18/04/07 (OT-07-0158)
            If Text1 <> "" Then
               If InStr(Text1, "@") > 0 Then
                  WINMAIL04.Text1 = TRIM$(Label18)
                     If WINMAIL04.Text1 = "" Then
                        WINMAIL04.Text1 = Text1
                     End If
                  WINMAIL04.Text3 = Text1
                  WINMAIL04.Text2 = "Envío de Listado"
                  WINMAIL04.Label5.Caption = ARCHISAL$
                     ASALIXX$ = TRIM$(ARCHISAL$)
25                   PPXX% = InStr(ASALIXX$, "\")
                     If PPXX% > 0 Then
                        ASALIXX$ = TRIM$(Mid$(ASALIXX$, PPXX% + 1))
                        GoTo 25
                     End If
                  WINMAIL04.Text5 = "Enviamos Archivo Adjunto: " + ASALIXX$
                  WINMAIL04.Show 1
               Else
                  Call MENSERR(24, "Dirección de Mail No Valida")
                  Text1.SetFocus
                  Exit Sub
               End If
            End If
            '18/04/07 (FIN)
            '
           WLITAB04.Hide
           Exit Sub
         End If
         Call SETPRINTER
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
    If ARCHISAL$ = "" Then Printer.Orientation = PORIAN
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    '
    If ARCHISAL$ = "SCREEN" Then
      PREVIEW.Show 1
      Exit Sub
    End If
    '
    On Error Resume Next
    Unload PREVIEW
    On Error GoTo 0
    '
    '18/04/07 (OT-07-0158)
    If Text1 <> "" Then
       If InStr(Text1, "@") > 0 Then
          WINMAIL04.Text1 = TRIM$(Label18)
             If WINMAIL04.Text1 = "" Then
                WINMAIL04.Text1 = Text1
             End If
          WINMAIL04.Text3 = Text1
          WINMAIL04.Label5.Caption = ARCHISAL$
          WINMAIL04.Text2 = "Envío de Listado"
             ASALIXX$ = TRIM$(ARCHISAL$)
30           PPXX% = InStr(ASALIXX$, "\")
             If PPXX% > 0 Then
                ASALIXX$ = TRIM$(Mid$(ASALIXX$, PPXX% + 1))
                GoTo 30
             End If
          WINMAIL04.Text5 = "Enviamos Archivo Adjunto: " + ASALIXX$
          WINMAIL04.Show 1
       Else
          Call MENSERR(24, "Dirección de Mail No Valida")
          Text1.SetFocus
          Exit Sub
       End If
    End If
    '18/04/07 (FIN)
    '
    WLITAB04.Hide
    '
    If ARCHISAL$ <> "" Then
        If WLITAB04.Check1.Value = 1 Then
            LCB$ = LUBAS$
            ULOWP$ = LUCOM$
            If Left$(JOBID$, 2) = "WF" Then ULOWP$ = App.Path + "\"
            If EXISTE%(ULOWP$ + "WORDPAD.EXE") = 1 Then
                 Close
                 Shell ULOWP$ + "WORDPAD.EXE " + ARCHISAL$, vbMaximizedFocus
               Else
                 MENSA$ = "Imposible Mostrar Archivo Generado."
                 MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
                 MENSA$ = MENSA$ + "\en Directorio '" + REPLACAR$(ULOWP$, "\", "/") + "'."
                 MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
                 Call COMUNI(MENSA$)
            End If
        End If
    End If
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
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Sub Command10_Click()
   Command10.Enabled = False
   Call Label11_DblClick
   Command10.Enabled = True
End Sub

Private Sub Command2_Click()
    '\\\OT-5-622-FG-13/09/05///
    OPX% = COMALTER%("Opciones de Configuración:\\Impresora\Listado\Leyendas Fijas")
    Select Case OPX%
      Case 1
           Call SETUPRINTER
      Case 2
           Call CONECRUN("#PARALIS", "Definicion de Listadores Tabulados")
      Case 3
           Call LEYENLIS
           Call COMUNI("Configuración de Leyendas Completada.\  \Vuelva a Pedir el Listado.")
           Call FINAL("")
    End Select
    '\\\OT-5-622-FG-FIN///
End Sub

Private Sub Command3_Click()
    Call GRADATIMPRE
    Call CARDATIMPRE
End Sub


Private Sub Command5_Click()
   '
   '17/04/07 (OT-07-0157)
   Call SELECHO("MAILIST")
   RAZO1$ = TRIM$(VALACT1$("MAILIST"))
   If RAZO1$ <> "" Then
      Label18 = RAZO1$
      Text1 = TRIM$(VALACT2$("MAILIST"))
   End If
   '17/04/07 (FIN)
   '
End Sub

Sub Command6_Click()
   '
   '17/04/07 (OT-07-0157)
   J& = 0
   Call APERBASDAT("CLIEN")
   Set CLIEN = Badaclie.OpenRecordset("SELECT * FROM Deudor12", dbOpenSnapshot)
   '
   With CLIEN
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 10
     End If
     '
     Do While Not .EOF
        '
        Y$ = REGBLAN$("MAILIST")
        On Error Resume Next
        RAZCLI$ = !Raso_Cli
        MAILCLI$ = !Mail_Cli
        On Error GoTo 0
        '
        If InStr(MAILCLI$, "@") > 0 Then
           Call REPLA(Y$, RAZCLI$, 1, 32)
           Call REPLA(Y$, LCase$(MAILCLI$), 33, 48)
           J& = J& + 1
           Call GRAREG("MAILIST", Y$, J&)
        End If
        '
        MAILCLI$ = ""
        RAZCLI$ = ""
        '
     .MoveNext
     Loop
     '
   End With
   '
10 Call APERBASDAT("PROVE")
   Set PROVE = Badaprov.OpenRecordset("SELECT * FROM Proved12", dbOpenSnapshot)
   '
   With PROVE
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 20
     End If
     '
     Do While Not .EOF
        '
        Y$ = REGBLAN$("MAILIST")
        On Error Resume Next
        RAZPROV$ = !Raso_Cli
        MAILPROV$ = !Mail_Cli
        On Error GoTo 0
        '
        If InStr(MAILPROV$, "@") > 0 Then
           Call REPLA(Y$, RAZPROV$, 1, 32)
           Call REPLA(Y$, LCase$(MAILPROV$), 33, 48)
           J& = J& + 1
           Call GRAREG("MAILIST", Y$, J&)
        End If
        '
        RAZPROV$ = ""
        MAILPROV$ = ""
        '
     .MoveNext
     Loop
     '
     Call SETULTREG("MAILIST", J&)
     '
20 End With
   Call COMUNI("Operación Completa")
   Call FILESORT("MAILIST", "", 1, 1, "ASC")
   '17/04/07 (FIN)
End Sub


Private Sub Dir1_Change()
   On Error Resume Next
   RUTASAL$ = Dir1.Path
   Label11.Caption = RUTASAL$
   On Error GoTo 0
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   On Error GoTo 0
End Sub

Private Sub Form_Load()
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFFFFFF
      Me.Frame1.BackColor = &HFFFFFF
      Me.Frame10.BackColor = &HFFFFFF
      Me.Frame11.BackColor = &HFFFFFF
      Me.Frame12.BackColor = &HFFFFFF
      Me.Frame13.BackColor = &HFFFFFF
      Me.Frame2.BackColor = &HFFFFFF
      Me.Frame3.BackColor = &HFFFFFF
      Me.Frame4.BackColor = &HFFFFFF
      Me.Frame5.BackColor = &HFFFFFF
      Me.Frame6.BackColor = &HFFFFFF
      Me.Frame7.BackColor = &HFFFFFF
      Me.Frame8.BackColor = &HFFFFFF
      Me.Frame9.BackColor = &HFFFFFF
      
      Me.Option2(0).BackColor = &HFFFFFF: Me.Option2(1).BackColor = &HFFFFFF
      Me.Option3(0).BackColor = &HFFFFFF: Me.Option3(1).BackColor = &HFFFFFF
      
      Me.Check1.BackColor = &HFFFFFF
      Me.Check2.BackColor = &HFFFFFF
      
      ' &H8000000F
    End If

   Refresh
   DoEvents
   Timer1.Enabled = True
   Timer2.Enabled = True
   MODOSAL% = 3
   DoEvents
End Sub
'

Sub Image1_Click(Index As Integer)
   MODOSAL% = Index
   If Index = 0 Then
        ARCHISAL$ = ""
        PRINTERSET% = 0
        IDL% = IDL0%
        Text2.TEXT = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        '17/04/07 (OT-07-0157)
        Text1.TEXT = ""
        '17/04/07 (fin)
        'Drive1.Enabled = False
        'Dir1.Enabled = False
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
        Text2.TEXT = ASALIDA$
           On Error Resume Next
        Option3(0).Value = True
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Check1.Enabled = True
        'Drive1.Enabled = True
        'Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 0
        Text2.Enabled = True
        Text2.SetFocus
        '17/04/07 (OT-07-0157)
        Text1.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        '17/04/07 (fin)
        On Error GoTo 0
      ElseIf Index = 2 Then
        RUTASAL$ = TRIM$(CONTROL$("EXPOREXC", IDENTER$))
        If RUTASAL$ = "" Then RUTASAL$ = Left$(LUDAT$, Len(LUDAT$) - 1)
        On Error Resume Next
        Drive1.Drive = Left$(RUTASAL$, 2)
        Dir1.Path = RUTASAL$
        On Error GoTo 0
        Label11.Caption = RUTASAL$
        ASALIDA$ = TRIM$(ARCHITRA$)
        Text2.TEXT = ASALIDA$
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        'Drive1.Enabled = True
        'Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 1
        Text2.Enabled = True
        Text2.SetFocus
        '17/04/07 (OT-07-0157)
        Text1.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        '17/04/07 (fin)
      ElseIf Index = 3 Then
        ARCHISAL$ = "SCREEN"
        '17/04/07 (OT-07-0157)
        Text1.TEXT = ""
        '17/04/07 (FIN)
        Text2.TEXT = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Command4.Enabled = False
        SALEXCEL% = 0
    End If
End Sub

Private Sub Text1_DblClick()
    Call Command5_Click '17/04/07 (OT-07-0157)
End Sub

Sub CARDATLIS()
     '
5800 GRABACT% = 0
     TOTALI% = 0
     POSITOT% = 0
     '
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
     WLITAB04.Label5 = ARCHITRA$
     WLITAB04.Label6 = TITUARCH$(ARCHITRA$)
5801 Get #N15%, 1, YYY$
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     WLITAB04.Label8 = PROPRIN$
5803 S1$ = TRIM$(Left$(YYY$, 48))
     WLITAB04.Label9 = TRIM$(S1$)
     '
     EPAC$ = TRIM$(EMPREAC$)
     If EPAC$ <> "" Then
         WLITAB04.Caption = WLITAB04.Caption + "  -  " + EPAC$
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
     IDL0% = IDL%
     REHOJX% = 0
     If Asc(Mid$(COD$, 7)) > 96 Then
       REHOJX% = Asc(Mid$(COD$, 7))
     End If
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
     If RG% > LOF(N15%) / 128 Or RG% > 49 Then GoTo 5820
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
        TOTALI% = 1
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
     Label15 = TRIM$(Str$(ATOT%))
     '
End Sub
'
        

Private Sub Image1_DblClick(Index As Integer)
   Call Image1_Click(Index)
   Call Command1_Click
End Sub

Private Sub Label11_DblClick()
   If Text2.Enabled = True Then
     SELCARPE.Show 1
   End If
End Sub

Private Sub Timer1_Timer()
   Static YACAR%, VECES
   If YACAR% = 0 Then
       Screen.MousePointer = 11
    On Error Resume Next
    PORIAN = Printer.Orientation
    '
    Printer.TrackDefault = False
    '
    DoEvents
    Call CARDATLIS
    DoEvents
    Call CARDATIMPRE
    DoEvents
    '
    ARCHISAL$ = ""
    Text2.TEXT = ARCHISAL$
    'Option1(0).Value = True        ' IMPRESORA
    MODOSAL% = 0
    Call Image1_Click(0)
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
    Combo1.TEXT = Printer.DeviceName
    IMPREDETER$ = Printer.DeviceName
    '
    If Left$(UCase$(ATRIBU$), 4) = "LPT1" Then
        Refresh
        Call SETPRINTER
        Call CARHEADERS
        Call ARMAENCA
        '
        If Option2(1).Value = True Then
              Printer.Orientation = 2
            Else
              Printer.Orientation = 1
        End If
        '
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
    'If EXISTE%("FLEXCEL.EXE") > 0 Then
      'Option1(2).Enabled = True
      Image1(2).Enabled = True
    'End If
    DoEvents
    Screen.MousePointer = 1
    YACAR% = 1
    AppActivate Caption
    SendKeys Chr$(1)
  End If
  '
  VECES = VECES + 1
  If VECES > 30 And VECES < 32 Then
    AppActivate Caption
    SendKeys Chr$(1)
    DoEvents
  End If
  '
End Sub

Private Sub Timer2_Timer()
   For KKI% = 0 To 3
     If KKI% <> MODOSAL% Then
         Image1(KKI%).Visible = True
       Else
         Image1(KKI%).Visible = Not Image1(KKI%).Visible
     End If
   Next KKI%
End Sub


Private Sub LEYENLIS()
    '\\\OT-5-622-FG-13/09/05///
    '
    'If TRIM$(CODFOR$) = "" Then
    '    Call MENSERR(24, "Falta Seleccionar Formulario")
    '    PRINFORM.Text1.SetFocus
    '    Exit Sub
    'End If
    '
    CODFOR$ = TRIM$(Label8)
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".PRL", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next I%
    '
    VTB% = VENTABU%("CARLEYEW")
    If VTB% >= 0 Then
        J% = 0
        For I% = 1 To LOF(NX%) / 128
            Get #NX%, I%, RF$
            If TRIM$(Left$(RF$, 8)) <> "TEXTWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next I%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "TEXTWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For I% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next I%
        '
    End If
    Close #NX%: NX% = 0
    '
    '\\\OT-5-622-FG-FIN///
End Sub
