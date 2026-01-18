VERSION 5.00
Begin VB.Form OPMAST04 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Configuración"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   8055
      Left            =   5040
      ScaleHeight     =   7995
      ScaleWidth      =   1170
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
      Begin VB.CommandButton Command5 
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
         Left            =   120
         Picture         =   "OPMAST04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Guardar "
         Top             =   7020
         Width           =   640
      End
      Begin VB.Image Image3 
         Height          =   8415
         Left            =   0
         Picture         =   "OPMAST04.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fórmula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   3550
      TabIndex        =   36
      Top             =   6900
      Width           =   1305
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cantidad"
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
         Left            =   90
         TabIndex        =   38
         Top             =   465
         Width           =   1125
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Uso"
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
         Left            =   90
         TabIndex        =   37
         Top             =   210
         Value           =   -1  'True
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1850
      TabIndex        =   27
      Top             =   6900
      Width           =   1575
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nro. Plano"
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
         Left            =   90
         TabIndex        =   29
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Código Alias"
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
         Left            =   120
         TabIndex        =   28
         Top             =   465
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   8100
      Left            =   5085
      ScaleHeight     =   8040
      ScaleWidth      =   900
      TabIndex        =   20
      Top             =   0
      Width           =   960
      Begin VB.CommandButton Command4 
         Caption         =   "TMU"
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
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "A-B-M Tabla de TMU's"
         Top             =   2835
         Width           =   640
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Check"
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
         Left            =   120
         Picture         =   "OPMAST04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   5985
         Width           =   640
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fam"
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
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "A-B-M Familias de Articulos"
         Top             =   2100
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   6930
         Width           =   640
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
      Begin VB.CommandButton command1 
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
         Left            =   120
         Picture         =   "OPMAST04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   300
         Width           =   640
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "OPMAST04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar "
         Top             =   990
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPMAST04.frx":2E00
         Top             =   7230
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reposición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   210
      TabIndex        =   17
      Top             =   6900
      Width           =   1530
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Prov. Habit."
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
         Left            =   90
         TabIndex        =   19
         Top             =   465
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterio"
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
         Left            =   90
         TabIndex        =   18
         Top             =   210
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Conversión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   210
         TabIndex        =   40
         Top             =   5750
         Width           =   4215
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Habilita Fórmula Por Millar"
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
            Left            =   735
            TabIndex        =   41
            Top             =   315
            Width           =   2745
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Atributos Paramétricos"
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
         TabIndex        =   30
         Top             =   4620
         Width           =   4215
         Begin VB.TextBox Text3 
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
            TabIndex        =   33
            Top             =   630
            Width           =   2355
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
            Left            =   1680
            TabIndex        =   31
            Top             =   315
            Width           =   2355
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo 'B':"
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
            Left            =   105
            TabIndex        =   34
            Top             =   630
            Width           =   1500
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo 'A':"
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
            Left            =   105
            TabIndex        =   32
            Top             =   315
            Width           =   1500
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterios y Tablas Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   210
         TabIndex        =   7
         Top             =   1470
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Venta según Tasa I.V.A."
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
            Index           =   4
            Left            =   315
            TabIndex        =   26
            Top             =   810
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tabla de Unidades de Medida"
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
            Index           =   9
            Left            =   315
            TabIndex        =   14
            Top             =   1065
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tabla de Tipos de Material"
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
            Index           =   0
            Left            =   315
            TabIndex        =   9
            Top             =   315
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos Contables de Venta"
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
            Index           =   3
            Left            =   315
            TabIndex        =   8
            Top             =   570
            Width           =   3375
         End
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
         Height          =   360
         Left            =   3255
         TabIndex        =   1
         Top             =   420
         Width           =   175
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Abrir"
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
         Left            =   3570
         Picture         =   "OPMAST04.frx":4D22
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   210
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
         Width           =   3225
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   3
         Text            =   " "
         Top             =   420
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Configuración Industrial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1450
         Left            =   210
         TabIndex        =   2
         Top             =   3045
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Categorias de Operarios Productivos"
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
            Index           =   8
            Left            =   210
            TabIndex        =   13
            Top             =   1065
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Máquinas y Equipos de Producción"
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
            Index           =   7
            Left            =   210
            TabIndex        =   12
            Top             =   810
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Secciones y Centros de Costos"
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
            Index           =   6
            Left            =   210
            TabIndex        =   11
            Top             =   570
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Depósitos y Almacenes"
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
            Left            =   210
            TabIndex        =   10
            Top             =   315
            Width           =   3795
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
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
         Left            =   -735
         TabIndex        =   6
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Criterios Técnicos de Clasificación"
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
      Index           =   1
      Left            =   630
      TabIndex        =   15
      Top             =   2940
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Clasificación Técnica"
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
      Index           =   2
      Left            =   630
      TabIndex        =   16
      Top             =   2940
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuFamilia 
         Caption         =   "A-B-M Familias de Articulos"
      End
      Begin VB.Menu mnutbu 
         Caption         =   "A-B-M Tabla de TMU's"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Aspecto 
      Caption         =   "Aspecto"
      Begin VB.Menu mnunewaspecto 
         Caption         =   "Nuevo Aspecto"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "OPMAST04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command16_Click()
    '
    Command16.Enabled = False
    For Index = 0 To 9
      If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
      End If
   Next Index
99 Command16.Enabled = True
   '
End Sub

Private Sub Command2_Click()
  MOREPOSI$ = "CRITREPO"
  If Option3.Value = True Then MOREPOSI$ = "PROHABIT"
  Call GRACONTROL("", "MOREPOSI", MOREPOSI$)
  Call GRACONTROL("AMASTO", "ATRIPAR1", Text12)
  Call GRACONTROL("AMASTO", "ATRIPAR2", Text3)
  MOFORMU$ = "USOUNI"
  If Option7.Value = True Then MOFORMU$ = "CANTIDAD"
  Call GRACONTROL("AMASTO", "FORMULA", MOFORMU$)
  '
  PORMILES$ = "NO": HABIPORMI% = 0
  If Check1.Value = 1 Then PORMILES$ = "SI": HABIPORMI% = 1
  Call GRACONTROL("AMASTO", "PORMILES", PORMILES$)
  '
  IDSECUN$ = "NUPLANO": INCLUALIAS% = (-1)
  If Option4.Value = True Then
    IDSECUN$ = "ALIAS"
    INCLUALIAS% = 1
  End If
  Call GRACONTROL("", "IDSECUN", IDSECUN$)
End Sub

Private Sub Command25_Click()
   Command25.Enabled = False
   PUNCODINT$ = String$(32767, 0)
   PUNCODDOS$ = String$(32767, 0)
   HUBORECU% = 0: HUBOERROR% = 0
   Call APERBASDAT("MASTER")
   With Master
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 99
     On Error GoTo 0
     '
     Do While Not .EOF
       CIXI% = !CODINT
       If CIXI% > 0 Then
         Mid$(PUNCODDOS$, CIXI%, 1) = Chr$(1)
       End If
     .MoveNext
     Loop
     '
     .MoveFirst
     Do While Not .EOF
       CIXI% = !CODINT
       If CIXI% > 0 Then
         If Asc(Mid$(PUNCODINT$, CIXI%, 1)) > 0 Then
           NUCOI% = InStr(PUNCODDOS$, Chr$(0))
           HUBOERROR% = 1
           TTXX$ = "Error Recuperable de Indices en Código '" + TRIM$(!CODIGO)
           TTXX$ = TTXX$ + "\   \Desea Recuperar Indices de este Item ?\\Si, Recuperar\No, Continuar Control"
           If COMALTER%(TTXX$) = 1 Then
             .Edit
             !CODINT = NUCOI%
             .Update
             Mid$(PUNCODDOS$, NUCOI%, 1) = Chr$(1)
             Call COMUNI("Recuerde Revisar Fórmulas y Movimientos\del Código Recuperado.")
             HUBORECU% = 1
           End If
           GoTo 19
         End If
         Mid$(PUNCODINT$, CIXI%, 1) = Chr$(1)
       End If
19   .MoveNext
     Loop
   End With
   '
   If HUBORECU% > 0 Then
     UPDATMASTER% = 1
     Call GENECOMAS
   End If
   '
90 On Error GoTo 0
   TTXX$ = "Fin del Control."
   If HUBOERROR% < 1 Then
     TTXX$ = TTXX$ + "\No se Detectaron Errores Recuperables."
   End If
   Call COMUNI(TTXX$)
   '
   If HUBOERROR% < 1 Then
     TTXX$ = "Desea Controlar la Integridad\de las Bases de Datos de Movimientos ?"
     TTXX$ = TTXX$ + "\   \Este Proceso Solo Debe Realizarse\por Personal Debidamente Calificado !!!"
     If COMALTER%(TTXX$ + "\\No, Finalizar Control\Si, Controlar") = 2 Then
       If USNBR% < 101 Then
         Call MENSERR(24, "La Opción Requiere Privilegios de Supervisor")
         GoTo 99
       End If
       Call VERIMOST
     End If
   End If
   '
99 On Error GoTo 0
   Screen.MousePointer = 1
   Command25.Enabled = True
   '
End Sub

Sub VERIMOST()
   '
10 Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then Exit Sub
   '
   FEDES = DES%
   HOYY = HOY(HOS$)
   If DIENTRE%(FEDES, HOYY) > 45 Then
     Call MENSERR(24, "Solo Pueden Controlarse\los Ultimos 45 Días.")
     GoTo 10
   End If
   '
   FEHAS = HAS%
   If FEHAS > HOYY Then
     Call MENSERR(24, "Fecha 'Hasta' Posterior a Hoy")
     GoTo 10
   End If
   '
   If HAS% < DES% Then
     Call MENSERR(24, "Rango de Fechas no Procesable")
     GoTo 10
   End If
   '
   HUBOERRO% = 0
   ' controla BOLRECEP
   Call APERBASDAT("STOCKS")
   '
   FF1D = CDbl(CVDAT(DES%))
   FF2H = CDbl(CVDAT(HAS%))
   '
   SQLX$ = " SELECT * FROM BOLRECEP "
   SQLX$ = SQLX$ + " WHERE cdbl(FechRecep) >= " & FF1D & " "
   SQLX$ = SQLX$ + " AND cdbl(FechRecep) <= " & FF2H & " "
   Set BRECEPRATMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With BRECEPRATMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       Call MENSERR(24, "No Existen Boletas de Recepcion\en Rango de Fechas Seleccionado.")
       On Error GoTo 0
       GoTo 30
     End If
     On Error GoTo 0
     JJ& = 0
     Do While (.EOF = False)
       CIXI% = !COD_INTE
       CODD$ = !Cod_Exte
       codex$ = CODEXT$(CIXI%)
       If CODD$ <> codex$ Then
         HUBOERRO% = 1
         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
           ' recupera
           CIXI1% = CODINT%(CODD$)
           If CIXI1% > 0 Then
               ' CAMBIA EL CODIGO INTERNO
               .Edit
               !COD_INTE = CIXI1%
               .Update
             Else
               ' NO EXISTE EL CODIGO EXTERNO
               DESCRINUE$ = !DescripItem
               TIMA% = 1
               UNID$ = !Unidad
               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
               CIXI1% = NUMAS%
               .Edit
               !COD_INTE = CIXI1%
               .Update
           End If
         End If
       End If
     .MoveNext
     Loop
   End With
   '
30  SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + " WHERE cdbl(FechMov) >= " & FF1D & " "
   SQLX$ = SQLX$ + " AND cdbl(FechMov) <= " & FF2H & " "
   Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With MOVISTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       Call MENSERR(24, "No Existen Movimientos de Stock\en Rango de Fechas Seleccionado.")
       On Error GoTo 0
       GoTo 39
     End If
     On Error GoTo 0
     JJ& = 0
     Do While (.EOF = False)
       CIXI% = !COD_INTE
       CODD$ = !Cod_Exte
       codex$ = CODEXT$(CIXI%)
       If CODD$ <> codex$ Then
         HUBOERRO% = 1
         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
           ' recupera
           CIXI1% = CODINT%(CODD$)
           If CIXI1% > 0 Then
               ' CAMBIA EL CODIGO INTERNO
               .Edit
               !COD_INTE = CIXI1%
               .Update
             Else
               ' NO EXISTE EL CODIGO EXTERNO
               DESCRINUE$ = !Descrip
               TIMA% = 1
               UNID$ = "U."
               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
               CIXI1% = NUMAS%
               .Edit
               !COD_INTE = CIXI1%
               .Update
           End If
         End If
       End If
     .MoveNext
     Loop
   End With
   '
39 Movisto.Close
   Bolrecep.Close
   '
   ' controla OCOMDETA
   Call APERBASDAT("COMPRAS")
   '
   SQLX$ = " SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + " WHERE cdbl(Femi_Ocom) >= " & FF1D & " "
   SQLX$ = SQLX$ + " AND cdbl(Femi_Ocom) <= " & FF2H & " "
   Set BRECEPRATMP = compras.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With BRECEPRATMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       Call MENSERR(24, "No Existen Ordenes de Compra\en Rango de Fechas Seleccionado.")
       On Error GoTo 0
       GoTo 50
     End If
     On Error GoTo 0
     JJ& = 0
     Do While (.EOF = False)
       CIXI% = !COD_INTE
       CODD$ = !Cod_Exte
       codex$ = CODEXT$(CIXI%)
       If CODD$ <> codex$ Then
         HUBOERRO% = 1
         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
           ' recupera
           CIXI1% = CODINT%(CODD$)
           If CIXI1% > 0 Then
               ' CAMBIA EL CODIGO INTERNO
               .Edit
               !COD_INTE = CIXI1%
               .Update
             Else
               ' NO EXISTE EL CODIGO EXTERNO
               DESCRINUE$ = !Descritem
               TIMA% = 1
               UNID$ = !Uni_Mas
               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
               CIXI1% = NUMAS%
               .Edit
               !COD_INTE = CIXI1%
               .Update
           End If
         End If
       End If
     .MoveNext
     Loop
   End With
   '
50 OcomEnca.Close
   '
   If HUBOERRO% = 0 Then
       Call COMUNI("Fin de Control de Movimientos\sin Errores Recuperables.")
     Else
       Call COMUNI("Fin del Control")
   End If
   '
End Sub
    '
 Sub Command3_Click()
    Command3.Enabled = False
    '//OT-07-0420
5   RTA% = ALTERNA%("ABM Familias \Estructura por Familia")
    Select Case RTA%
    Case Is = 1
      GoTo 10
    Case Is = 2
      GoTo 20
    Case Else
      GoTo 99
    End Select
    '//OT-07-0420
10  If DERACCE("TAFAMFIL", "Actualizar Tabla de Familias y Valores") < 2 Then GoTo 99
    '
    Call TRALOCAL("TAFAMFIL", "")
    For U& = 1 To ULTREG&("!TAFAMFIL")
        X$ = REGLEIDO$("!TAFAMFIL", U&)
        CI% = CVI(Mid$(X$, 41, 2))
        PRECAL# = PRECALCU(CI%)
        Call REPLA(X$, MKS$(PRECAL#), 43, 4)
        Call GRAREG("!TAFAMFIL", X$, U&)
    Next U&

15  VTB% = VENTABU("TAFAMFIL")
    If VTB% >= 0 Then
      FIN& = ULTREG&("!TAFAMFIL")
      PRENUEVO% = 0
      For U& = 1 To FIN&
        X$ = REGLEIDO$("!TAFAMFIL", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        Call REPLA(X$, CD$, 1, 2)
        Call GRAREG("!TAFAMFIL", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(CD$) < 2 Then
          Call MENSERR(24, "Código de Familia Menor a Dos Caracteres")
          GoTo 15
        End If
        ' Valida que la familia posea descripcion
        DESC$ = TRIM$(Mid$(X$, 3, 27))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Codigo de Familia sin Descripción")
          GoTo 15
        End If
        '\\\OT-07-451-FL-22/11/07///
        ' Valida que se haya ingresado un codigo base
        CODIBAS% = CVI(Mid$(X$, 41, 2))
        If CODIBAS% < 1 Or CODIBAS% > NUMAS% Then
           Call MENSERR(24, "Falta Código de Producto Base")
           GoTo 15
        End If
        '\\\OT-07-451-FL-FIN///
        ' Valida quel codigo no este repetido
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TAFAMFIL", U2&)
          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
            Call MENSERR(24, "Código de Familia Repetido")
            GoTo 15
          End If
        Next U2& '
        Y$ = Mid$(X$, 1, 32) 'EN y$ guardo los datos que quiero conseguir del registro seleccionado que en este caso no son todos
        Call GRAREG("!TAFAMIL", Y$, U&) 'grabo el registro en la tabla local tafamil la posicion U& el registro Y$ ya recortado
        '
        If Abs(CVD(Mid$(X$, 33, 8)) - CVS(Mid$(X$, 43, 4))) >= 0.005 Then
           PRENUEVO% = 1
        End If
      Next U&
      '
      Call SETULTREG("!TAFAMFIL", FIN&)
      Call SETULTREG("!TAFAMIL", FIN&)
      '
      If PRENUEVO% = 1 Then
         If COMALTER%("Los Precios de Venta Calculados Difieren\de los Unitarios Registrados para Alguna Familia.\   \¿ Desea Actualizar los Precios Registrados ?\\Si, Actualizar\No Mantener") = 1 Then
            For U& = 1 To FIN
               X$ = REGLEIDO$("!TAFAMFIL", U&)
               PRECA# = CVS(Mid$(X$, 43, 4))
               Call REPLA(X$, MKD$(PRECA#), 33, 8)
               Call GRAREG("!TAFAMFIL", X$, U&)
            Next U&
            GoTo 15
         End If
      End If
      '
      Call TRACENTRAL("TAFAMIL", "") 'transfiero lo grabado en la tabla local a la tafamil del servidor
      Call TRACENTRAL("TAFAMFIL", "")
    End If
    '
    GoTo 5 ' VOLVEMOS A LA VENTANA DE OPCIONES
    '
    '//OT-07-0420
20  Call SELECHO("TAFAMFIL") 'Abre una ventana con
    n$ = VALACT1$("TAFAMFIL")
    Call FILTERFILE("STRUCFAM", "!STRUCFAM", 1, n$)
    '
    TITUPANTA$ = "Estructura Familia " + n$ + " - " + VALACT2$("TAFAMFIL")
    VTB% = VENTABU("STRUCFAM/TITUPAN=" + TITUPANTA$)

    If VTB% >= 0 Then
       For U& = 1 To ULTREG&("!STRUCFAM")
         XX$ = REGLEIDO$("!STRUCFAM", U&)
         Call REPLA(XX$, n$, 1, 2)
         Call GRAREG("!STRUCFAM", XX$, U&)
         MODOCALC$ = TRIM$(UCase$(Mid$(XX$, 13, 32)))
       Next U&
       '
       Call CIERRARCH("!STRUCFAM")
       Call FILTERUPDATE("STRUCFAM", "", 1, n$)
    End If
  '//OT-07-0420
    GoTo 5 ' VOLVEMOS A LA VENTANA DE OPCIONES
    '
99  Command3.Enabled = True
    '
End Sub


Sub Command4_Click()
  '\\\OT-06-0252-WAR-22/06/06///
    Command4.Enabled = False
    '
    Call TRALOCAL("TABLATMU", "")
    '
10  VTB% = VENTABU("TABLATMU")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TABLATMU")
        X$ = REGLEIDO$("!TABLATMU", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 4)))
        Call REPLA(X$, CD$, 1, 4)
        Call GRAREG("!TABLATMU", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a 1 caracteres
        If Len(CD$) < 1 Then
          Call MENSERR(24, "Código de TMU no Válido")
          GoTo 10
        End If
        ' Valida que el TMU posea descripcion
        DESC$ = TRIM$(Mid$(X$, 5, 32))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "TMU sin Descripción")
          GoTo 10
        End If
        ' Valida que el TMU no esté repetido
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TABLATMU", U2&)
          If TRIM$(UCase$(Mid$(Y$, 1, 4))) = CD$ Then
            Call MENSERR(24, "Código de TMU Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TABLATMU", "")
    End If
    '
99  Command4.Enabled = True
    '
  '\\\OT-06-0252-WAR-FIN///
End Sub

Private Sub Command5_Click()
   Call Command2_Click
End Sub

'
Private Sub Form_Load()
  '
  mnunewaspecto.Checked = False
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &H8000000F
      Picture2.Visible = True
      Frame1.BackColor = &H8000000F
      Frame2.BackColor = &H8000000F
      Frame3.BackColor = &H8000000F
      Frame4.BackColor = &H8000000F
      Frame5.BackColor = &H8000000F

      Frame6.BackColor = &H8000000F
      Frame9.BackColor = &H8000000F
      Frame11.BackColor = &H8000000F
      For I% = 0 To 9
         Option1(I%).BackColor = &H8000000F
      Next I%
      Check1.BackColor = &H8000000F
      Option2.BackColor = &H8000000F
      Option3.BackColor = &H8000000F
      Option4.BackColor = &H8000000F
      Option5.BackColor = &H8000000F
      Option6.BackColor = &H8000000F
      Option7.BackColor = &H8000000F
      mnunewaspecto.Checked = True
      
  End If

  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "TACRITEC"
  TABLA$(2) = "TACLATEC"
  TABLA$(3) = "GRUCOVEN"       ' grupos contables por articulo
  TABLA$(4) = "TASIVA"
  TABLA$(5) = "TADEPOSI"
  TABLA$(6) = "SECPROD"
  TABLA$(7) = "PADMAQUI"
  TABLA$(8) = "CATOPER"
  TABLA$(9) = "UMEDIDA"
  '
  If CONTROL$("", "MOREPOSI") = "PROHABIT" Then
    Option3.Value = True
  End If
  '
  If CONTROL$("", "IDSECUN") = "ALIAS" Then
    Option4.Value = True
  End If
  '
  If CONTROL$("AMASTO", "FORMULA") = "CANTIDAD" Then
    Option7.Value = True
  End If
  '
  If CONTROL$("AMASTO", "PORMILES") = "SI" Then
    Check1.Value = 1
  End If
  '
  Text12 = CONTROL$("AMASTO", "ATRIPAR1")
  Text3 = CONTROL$("AMASTO", "ATRIPAR2")
  '
End Sub

Private Sub MnuAspecto_Click()
    
End Sub

Private Sub mnuFamilia_Click()
    Call Command3_Click
End Sub

Private Sub mnunewaspecto_Click()
    mnunewaspecto.Checked = False
    If CLACONTRA% = 1 Then
       If CONTROL("NEWASPEC", "CAMBIOAS") = "" Then
          Call GRACONTROL("NEWASPEC", "CAMBIOAS", "SI")
          mnunewaspecto.Checked = True
       Else
          Call GRACONTROL("NEWASPEC", "CAMBIOAS", "")
       End If
    End If
End Sub

Private Sub mnutbu_Click()
   Call Command4_Click
End Sub

'
 Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 9
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
    Text1.TEXT = TABLA(Index)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    Command16.Enabled = True
    'Command16.SetFocus 'No se puede asignar el foco a
                        'un objeto no visible.
                        '(cuando lo llamo desde la barra
                        ' de menu)
    '
End Sub

Sub CARTABLA(Index)
   '
   If Index = 0 Then
       Call CARTATIMA
    'ElseIf Index = 1 Then
       'Call CARCRITEC
    'ElseIf Index = 2 Then
       'Call CARTACLAS
    ElseIf Index = 3 Then
       Call CARGRUCOVE
    ElseIf Index = 4 Then
       Call CARTASIVA
    ElseIf Index = 5 Then
       Call CARTADEPOSI
    ElseIf Index = 6 Then
       Call CARSECPROD
    ElseIf Index = 7 Then
       Call CARMAQUI
    ElseIf Index = 8 Then
       Call CARCATOPER
    ElseIf Index = 9 Then
       Call CARUMEDIDA
   End If
   '
End Sub
'
Private Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_gotfocus()
   Command16.SetFocus
End Sub

Sub CARTASIVA()
20 Call TRALOCAL("GRUPIVA", "")
21 VTB% = VENTABU%("GRUPIVA/NC/TITUPAN=Tasa IVA por Categoria")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUPIVA")
         X$ = REGLEIDO$("!GRUPIVA", U&)
         Call REPLA(X$, Chr$(U& - 1), 1, 1)
         Call GRAREG("GRUPIVA", X$, U&)
      Next U&
      FIN& = ULTREG&("!GRUPIVA")
      Call SETULTREG("GRUPIVA", FIN&)
   End If
   If ULTREG&("GRUPIVA") > 4 Then
      '\\\OT-06-0536-RG-12/12/06///
      OPT% = COMALTER%("FLEXOFT Está Diseñado para Procesar\Hasta 4 Tasas de IVA.\   \Los Registros Excedentes Pueden No Ser\Considerados en Algunos Procesos.\\Corregir\Aprobar")
      If OPT% <> 2 Then
         GoTo 21
      End If
      '\\\OT-06-0536-RG-FIN///
   End If
      'Call SETULTREG("GRUPIVA", 4)
      'Call COMUNI("Truncado a Máximo de 4 Categorías de IVA")
   Call CIERRARCH("GRUPIVA")
End Sub

Sub CARTADEPOSI()
    '
    Call BLANARCH("!TADEPOSI")
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call REGAPP("!TADEPOSI", Y$)
              End If
           End If
        End If
    Next U&
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, J&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, J&)
               Call GRAREG("TADEPOSI", X$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", Y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub

Sub CARCATOPER()
   '
   Call TRALOCAL("CATOPER", "CATOPER")
   VTB% = VENTABU%("CATOPER")
   If VTB% >= 0 Then
      PUCONTROL$ = String$(255, 0)
      JJ& = 0
      For U& = 1 To ULTREG&("!CATOPER")
         X$ = REGLEIDO$("!CATOPER", U&)
         COCAT% = Asc(Left$(X$, 1))
         If COCAT% > 0 Then
           If COCAT% < 256 Then
             If Mid$(PUCONTROL$, COCAT%, 1) = Chr$(0) Then
               Mid$(PUCONTROL$, COCAT%, 1) = Chr$(1)
               JJ& = JJ& + 1
               Call GRAREG("CATOPER", X$, JJ&)
             End If
           End If
         End If
      Next U&
      Call SETULTREG("CATOPER", JJ&)
      Call CIERRARCH("CATOPER")
   End If
   '
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         X$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(X$, 4)) <> "" Then
           If Len(TRIM$(Left$(X$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           Y$ = Left$(X$, 2) + Mid$(X$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", X$, JJ&)
           Call GRAREG("UMEDIDA", Y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub


Sub CARMAQUI()
    '
    J& = 0
    For U& = 1 To ULTREG&("PADMAQ")
        X$ = REGLEIDO$("PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("!PADMAQ", X$, J&)
        End If
    Next U&
    Call SETULTREG("!PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
    VTB% = VENTABU%("CARMAQUI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!PADMAQ")
        X$ = REGLEIDO$("!PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("PADMAQ", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
End Sub






