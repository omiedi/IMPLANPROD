VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form OPMAST07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestro de Items de Stock - Configuración"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   60
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Conversión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   90
      TabIndex        =   52
      Top             =   6360
      Width           =   1455
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "F. por Millar"
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
         Left            =   80
         TabIndex        =   54
         Top             =   270
         Width           =   1320
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Form.Patron"
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
         Left            =   80
         TabIndex        =   53
         Top             =   550
         Width           =   1320
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2440
      TabIndex        =   48
      Top             =   6360
      Width           =   1360
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         MaxLength       =   2
         TabIndex        =   49
         Top             =   435
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -60
         TabIndex        =   51
         Top             =   465
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   50
         Top             =   465
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Unidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1560
      TabIndex        =   45
      Top             =   6360
      Width           =   860
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "KG/U."
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
         Left            =   80
         TabIndex        =   47
         Top             =   210
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "M2/U."
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
         Left            =   80
         TabIndex        =   46
         Top             =   465
         Width           =   750
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Descrip."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3840
      TabIndex        =   42
      Top             =   6360
      Width           =   1155
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Larga"
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
         Left            =   210
         TabIndex        =   44
         Top             =   270
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Doble"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   43
         Top             =   525
         Width           =   885
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4170
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2940
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fórmula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   3315
      TabIndex        =   35
      Top             =   7215
      Width           =   2760
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Requiere Aprobar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   55
         Top             =   240
         Width           =   1170
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cant."
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
         Left            =   210
         TabIndex        =   37
         Top             =   465
         Width           =   885
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Uso"
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
         Left            =   210
         TabIndex        =   36
         Top             =   210
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   1200
         X2              =   1200
         Y1              =   120
         Y2              =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1680
      TabIndex        =   27
      Top             =   7215
      Width           =   1575
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nro. Plano"
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   28
         Top             =   465
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   6320
      Left            =   5280
      ScaleHeight     =   6255
      ScaleWidth      =   1020
      TabIndex        =   20
      Top             =   0
      Width           =   1080
      Begin VB.CommandButton Command7 
         Caption         =   "Mod. Cód."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   41
         ToolTipText     =   "Agregar 'P' a Papel"
         Top             =   4530
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Subst"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   40
         Top             =   3870
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Impo Timat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   39
         Top             =   3210
         Visible         =   0   'False
         Width           =   640
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -120
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command4 
         Caption         =   "TMU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         TabIndex        =   38
         ToolTipText     =   "A-B-M Tabla de TMU's"
         Top             =   2495
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Fam"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         TabIndex        =   25
         ToolTipText     =   "A-B-M Familias de Articulos"
         Top             =   1800
         Visible         =   0   'False
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   5640
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPMAST07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   120
         Width           =   640
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPMAST07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar "
         Top             =   990
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPMAST07.frx":0454
         Top             =   5880
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reposición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   90
      TabIndex        =   17
      Top             =   7215
      Width           =   1530
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Prov. Habit."
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
         Left            =   90
         TabIndex        =   19
         Top             =   465
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterio"
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
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6100
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   4995
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Atributos Paramétricos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   330
         TabIndex        =   30
         Top             =   4880
         Width           =   4455
         Begin VB.CheckBox Check5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "No Repite"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3150
            TabIndex        =   59
            Top             =   700
            Width           =   1200
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "No Repite"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3150
            TabIndex        =   58
            Top             =   360
            Width           =   1200
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
            Height          =   315
            Left            =   1200
            TabIndex        =   33
            Top             =   700
            Width           =   1875
         End
         Begin VB.TextBox Text12 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   31
            Top             =   350
            Width           =   1875
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo 'B':"
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
            Left            =   105
            TabIndex        =   34
            Top             =   800
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Atributo 'A':"
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
            Index           =   7
            Left            =   120
            TabIndex        =   32
            Top             =   450
            Width           =   1020
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterios y Tablas Generales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   330
         TabIndex        =   7
         Top             =   1250
         Width           =   4455
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tabla de Familias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   435
            TabIndex        =   62
            Top             =   1295
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Venta según Tasa I.V.A."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   435
            TabIndex        =   26
            Top             =   795
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tabla de Unidades de Medida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   435
            TabIndex        =   14
            Top             =   1050
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tabla de Tipos de Material"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   435
            TabIndex        =   9
            Top             =   285
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos Contables de Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   435
            TabIndex        =   8
            Top             =   540
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
         Left            =   3615
         TabIndex        =   1
         Top             =   300
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
         Left            =   3930
         Picture         =   "OPMAST07.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   825
         Width           =   3465
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   3
         Text            =   " "
         Top             =   300
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Configuración Industrial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   330
         TabIndex        =   2
         Top             =   3000
         Width           =   4455
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Herramentales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   450
            TabIndex        =   57
            Top             =   1240
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Plantas Productivas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   450
            TabIndex        =   56
            Top             =   220
            Width           =   3795
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Categorias de Operarios Productivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   450
            TabIndex        =   13
            Top             =   1440
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Máquinas y Equipos de Producción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   450
            TabIndex        =   12
            Top             =   970
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Secciones y Centros de Costos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   450
            TabIndex        =   11
            Top             =   730
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Depósitos y Almacenes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   450
            TabIndex        =   10
            Top             =   475
            Width           =   3795
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
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
         Left            =   -375
         TabIndex        =   6
         Top             =   405
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPMAST07.frx":2680
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "OPMAST07.frx":28B4
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPMAST07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Sub ABMDEPOMAIL()
    
10  Call SELECHO("TADEPOSI") ' MUESTRA TABLA DE DEPOSITOS
    NUDEP% = XVALO(VALACT1$("TADEPOSI")) ' DEPOSITO SELECCIONADO
    If NUDEP% < 1 Then Exit Sub
    NUMEDEP$ = TRIM$(Str$(NUDEP%)) '
    MAILX$ = CONTROL("DIREMAIL", "DEPO-" & NUMEDEP$)
20  MAIX$ = TRIM$(InputBox$("Ingrese el Mail Para Depósito " & NUMEDEP$, , MAILX$))
    '
    'GUARDA LA DIRECCION DE CORREO
    If MAIX$ = "" Then GoTo 10
    If InStr(MAIX$, "@") > 1 Then
        Call GRACONTROL("DIREMAIL", "DEPO-" & NUMEDEP$, MAIX$)
    Else
        Call COMUNI("Mail no Válido")
        MAILX$ = MAIX$
        GoTo 20
    End If
    '
   GoTo 10

End Sub

Sub CARFAMILIA()
   Call TRALOCAL("TAFAMIL", "")

10  VTB% = VENTABU("TAFAMIL")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAFAMIL")
        X$ = REGLEIDO$("!TAFAMIL", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        Call REPLA(X$, CD$, 1, 2)
        Call GRAREG("!TAFAMIL", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(CD$) < 2 Then
          Call MENSERR(24, "Código de Familia Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        DESC$ = TRIM$(Mid$(X$, 3, 27))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Codigo de Familia sin Descripción")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TAFAMIL", U2&)
          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
            Call MENSERR(24, "Código de Familia Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TAFAMIL", "")
    End If
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command16_Click()
    '
    Command16.Enabled = False
    For Index = 0 To 12
      If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
      End If
   Next Index
99 Command16.Enabled = True
   '
End Sub

Private Sub Command2_Click()
  Screen.MousePointer = 11
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
  'OPCION DE VALIDAR REPETICION DE ATRIBUTOS
  OPT$ = "NO"
  If Check4.Value = 1 Then OPT$ = "SI"
  Call GRACONTROL("AMASTO", "REPATRIA", OPT$)
  '
  OPT$ = "NO"
  If Check5.Value = 1 Then OPT$ = "SI"
  Call GRACONTROL("AMASTO", "REPATRIB", OPT$)
  '
  
  IDSECUN$ = "NUPLANO": INCLUALIAS% = (-1)
  If Option4.Value = True Then
    IDSECUN$ = "ALIAS"
    INCLUALIAS% = 1
  End If
  Call GRACONTROL("", "IDSECUN", IDSECUN$)
  '
  LOMANT% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
  If LOMANT% < 15 Or LONCODI% > 24 Then LOMANT% = 15
  
  LOMA% = XVALO(TRIM$(Text4))
  If LOMA% >= 15 And LOMA% <= 24 Then
     Call GRACONTROL("MASTER", "LOMAXCOD", Str(LOMA%))
     Text4 = LOMA%
     If LENFIELD%("ECOMAST", 1) <> LOMA% Then
        Call SETLENFIELD("ECOMAST", 1, LOMA%)
        UPDATMASTER% = 1
        Call GENECOMAS
        Call FINAL("")
     End If
     If LOMA% <> LOMANT% Then
        UPDATMASTER% = 1
        Call GENECOMAS
     End If
   Else
     Call MENSERR(24, "Ingreso No Válido (Solo Valores entre 15 y 24)")
     Text4.SetFocus
  End If
  '
  'Graba Control si utiliza descripcion larga o descripcion de codigo y alias
  DESCRILA$ = ""
  If Option9.Value = True Then DESCRILA$ = "SI"
  Call GRACONTROL("MASTER", "DESLARGA", DESCRILA$)
  '
  'Graba control si presenta o no formula Patron
  PATTRONN$ = ""
  If Check2.Value = 1 Then PATTRONN$ = "SI"
  Call GRACONTROL("MASTER", "FORPATRO", PATTRONN$)
  '
  MEPORUN$ = ""
  If Option11.Value = True Then
    MEPORUN$ = "SI"
  End If
  Call GRACONTROL("MASTER", "M2PORUN", MEPORUN$)
  '
  APRO$ = "NO"
  If Check3.Value = 1 Then APRO$ = "SI"
  Call GRACONTROL("FORMULAS", "REQAPRO", APRO$)
  '
  'SI ES NO NO PRESENTA EL DOMICILIO FISCAL EN SOLAPA DE ENTEGA CUANDO ESTE NO TIENE INFORMACION
  DOMIFI$ = "SI"
  If Check6.Value = 0 Then DOMIFI$ = "NO"
  Call GRACONTROL("MASTER", "DOMIFISC", DOMIFI$)
  
  Call COMUNI("Reconfiguración Realizada - Re-Inicie la Aplicación")
  Call FINAL("")
  Screen.MousePointer = 1
End Sub


'jandy sql fue quitada en la version 7 sql
'Private Sub Command25_Click()
'   Command25.Enabled = False
'   PUNCODINT$ = String$(32767, 0)
'   PUNCODDOS$ = String$(32767, 0)
'   HUBORECU% = 0: HUBOERROR% = 0
'   Call APERBASDAT("MASTER")
'   With Master
'     On Error Resume Next
'     .MoveFirst
'     If Err Then GoTo 99
'     On Error GoTo 0
'
'     Do While Not .EOF
'       CIXI% = !CODINT
'       If CIXI% > 0 Then
'         Mid$(PUNCODDOS$, CIXI%, 1) = Chr$(1)
'       End If
'     .MoveNext
'     Loop
'
'     .MoveFirst
'     Do While Not .EOF
'       CIXI% = !CODINT
'       If CIXI% > 0 Then
'         If Asc(Mid$(PUNCODINT$, CIXI%, 1)) > 0 Then
'           NUCOI% = InStr(PUNCODDOS$, Chr$(0))
'           HUBOERROR% = 1
'           TTXX$ = "Error Recuperable de Indices en Código '" + TRIM$(!CODIGO)
'           TTXX$ = TTXX$ + "\   \Desea Recuperar Indices de este Item ?\\Si, Recuperar\No, Continuar Control"
'           If COMALTER%(TTXX$) = 1 Then
'             .Edit
'             !CODINT = NUCOI%
'             .Update
'             Mid$(PUNCODDOS$, NUCOI%, 1) = Chr$(1)
'             Call COMUNI("Recuerde Revisar Fórmulas y Movimientos\del Código Recuperado.")
'             HUBORECU% = 1
'           End If
'           GoTo 19
'         End If
'         Mid$(PUNCODINT$, CIXI%, 1) = Chr$(1)
'       End If
'19   .MoveNext
'     Loop
'   End With
'
'   If HUBORECU% > 0 Then
'     UPDATMASTER% = 1
'     Call GENECOMAS
'   End If
'
'90 On Error GoTo 0
'   TTXX$ = "Fin del Control."
'   If HUBOERROR% < 1 Then
'     TTXX$ = TTXX$ + "\No se Detectaron Errores Recuperables."
'   End If
'   Call COMUNI(TTXX$)
'
'   If HUBOERROR% < 1 Then
'     TTXX$ = "Desea Controlar la Integridad\de las Bases de Datos de Movimientos ?"
'     TTXX$ = TTXX$ + "\   \Este Proceso Solo Debe Realizarse\por Personal Debidamente Calificado !!!"
'     If COMALTER%(TTXX$ + "\\No, Finalizar Control\Si, Controlar") = 2 Then
'       If USNBR% < 101 Then
'         Call MENSERR(24, "La Opción Requiere Privilegios de Supervisor")
'         GoTo 99
'       End If
'       Call VERIMOST
'     End If
'   End If
'
'99 On Error GoTo 0
'   Screen.MousePointer = 1
'   Command25.Enabled = True
'
'End Sub

'jandy sql fue quitada en la version 7 sql
'Sub VERIMOST()
'   '
'10 Call DESHASFECHA(DES%, HAS%, PERIO$)
'   If PERIO$ = "" Then Exit Sub
'   '
'   FEDES = DES%
'   HOYY = HOY(HOS$)
'   If DIENTRE%(FEDES, HOYY) > 45 Then
'     Call MENSERR(24, "Solo Pueden Controlarse\los Ultimos 45 Días.")
'     GoTo 10
'   End If
'   '
'   FEHAS = HAS%
'   If FEHAS > HOYY Then
'     Call MENSERR(24, "Fecha 'Hasta' Posterior a Hoy")
'     GoTo 10
'   End If
'   '
'   If HAS% < DES% Then
'     Call MENSERR(24, "Rango de Fechas no Procesable")
'     GoTo 10
'   End If
'   '
'   HUBOERRO% = 0
'   ' controla BOLRECEP
'   Call APERBASDAT("STOCKS")
'   '
'   FF1D = CDbl(CVDAT(DES%))
'   FF2H = CDbl(CVDAT(HAS%))
'   '
'   SQLX$ = " SELECT * FROM BOLRECEP "
'   SQLX$ = SQLX$ + " WHERE cdbl(FechRecep) >= " & FF1D & " "
'   SQLX$ = SQLX$ + " AND cdbl(FechRecep) <= " & FF2H & " "
'   Set BRECEPRATMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
'   '
'   With BRECEPRATMP
'     On Error Resume Next
'     .MoveFirst
'     If Err Then
'       Call MENSERR(24, "No Existen Boletas de Recepcion\en Rango de Fechas Seleccionado.")
'       On Error GoTo 0
'       GoTo 30
'     End If
'     On Error GoTo 0
'     JJ& = 0
'     Do While (.EOF = False)
'       CIXI% = !Cod_Inte
'       CODD$ = !COD_EXTE
'       codex$ = CODEXT$(CIXI%)
'       If CODD$ <> codex$ Then
'         HUBOERRO% = 1
'         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
'         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
'         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
'         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
'         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
'           ' recupera
'           CIXI1% = CODINT%(CODD$)
'           If CIXI1% > 0 Then
'               ' CAMBIA EL CODIGO INTERNO
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'             Else
'               ' NO EXISTE EL CODIGO EXTERNO
'               DESCRINUE$ = !DescripItem
'               TIMA% = 1
'               UNID$ = !Unidad
'               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
'               CIXI1% = NUMAS%
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'           End If
'         End If
'       End If
'     .MoveNext
'     Loop
'   End With
'   '
'30  SQLX$ = " SELECT * FROM MOVISTO "
'   SQLX$ = SQLX$ + " WHERE cdbl(FechMov) >= " & FF1D & " "
'   SQLX$ = SQLX$ + " AND cdbl(FechMov) <= " & FF2H & " "
'   Set MOVISTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
'   '
'   With MOVISTMP
'     On Error Resume Next
'     .MoveFirst
'     If Err Then
'       Call MENSERR(24, "No Existen Movimientos de Stock\en Rango de Fechas Seleccionado.")
'       On Error GoTo 0
'       GoTo 39
'     End If
'     On Error GoTo 0
'     JJ& = 0
'     Do While (.EOF = False)
'       CIXI% = !Cod_Inte
'       CODD$ = !COD_EXTE
'       codex$ = CODEXT$(CIXI%)
'       If CODD$ <> codex$ Then
'         HUBOERRO% = 1
'         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
'         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
'         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
'         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
'         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
'           ' recupera
'           CIXI1% = CODINT%(CODD$)
'           If CIXI1% > 0 Then
'               ' CAMBIA EL CODIGO INTERNO
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'             Else
'               ' NO EXISTE EL CODIGO EXTERNO
'               DESCRINUE$ = !Descrip
'               TIMA% = 1
'               UNID$ = "U."
'               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
'               CIXI1% = NUMAS%
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'           End If
'         End If
'       End If
'     .MoveNext
'     Loop
'   End With
'   '
'39 Movisto.Close
'   Bolrecep.Close
'   '
'   ' controla OCOMDETA
'   Call APERBASDAT("COMPRAS")
'   '
'   SQLX$ = " SELECT * FROM OCOMDETA "
'   SQLX$ = SQLX$ + " WHERE cdbl(Femi_Ocom) >= " & FF1D & " "
'   SQLX$ = SQLX$ + " AND cdbl(Femi_Ocom) <= " & FF2H & " "
'   Set BRECEPRATMP = Compras.OpenRecordset(SQLX$, dbOpenDynaset)
'   '
'   With BRECEPRATMP
'     On Error Resume Next
'     .MoveFirst
'     If Err Then
'       Call MENSERR(24, "No Existen Ordenes de Compra\en Rango de Fechas Seleccionado.")
'       On Error GoTo 0
'       GoTo 50
'     End If
'     On Error GoTo 0
'     JJ& = 0
'     Do While (.EOF = False)
'       CIXI% = !Cod_Inte
'       CODD$ = !COD_EXTE
'       codex$ = CODEXT$(CIXI%)
'       If CODD$ <> codex$ Then
'         HUBOERRO% = 1
'         TTXX$ = "Código '" + CODD$ + "' Cruzado con\'" + codex$ + " - " + DESCRIT0$(CIXI%) + "'"
'         TTXX$ = TTXX$ + "\  \Esto Puede Deberse a una Sustitución de Código"
'         TTXX$ = TTXX$ + "\o a una Restauración Defectuosa de la BD de Artículos."
'         TTXX$ = TTXX$ + "\  \Desea Recuperar el Código '" + CODD$ + "' ?"
'         If COMALTER%(TTXX$ + "\\No, Continuar Control\Si, Recuperar") = 2 Then
'           ' recupera
'           CIXI1% = CODINT%(CODD$)
'           If CIXI1% > 0 Then
'               ' CAMBIA EL CODIGO INTERNO
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'             Else
'               ' NO EXISTE EL CODIGO EXTERNO
'               DESCRINUE$ = !Descritem
'               TIMA% = 1
'               UNID$ = !Uni_Mas
'               Call AGRECOD(CODD$, DESCRINUE$, TIMA%, UNID$, OKX%)
'               CIXI1% = NUMAS%
'               .Edit
'               !Cod_Inte = CIXI1%
'               .Update
'           End If
'         End If
'       End If
'     .MoveNext
'     Loop
'   End With
'   '
'50 OcomEnca.Close
'   '
'   If HUBOERRO% = 0 Then
'       Call COMUNI("Fin de Control de Movimientos\sin Errores Recuperables.")
'     Else
'       Call COMUNI("Fin del Control")
'   End If
'   '
'End Sub
'
'
'''Private Sub Command3_Click()
'''    Command3.Enabled = False
'''    '
'''    Call TRALOCAL("TAFAMIL", "")
'''10
'''    VTB% = VENTABU("TAFAMIL")
'''    If VTB% >= 0 Then
'''      For U& = 1 To ULTREG&("!TAFAMIL")
'''        X$ = REGLEIDO$("!TAFAMIL", U&)
'''        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
'''        Call REPLA(X$, CD$, 1, 2)
'''        Call GRAREG("!TAFAMIL", X$, U&)
'''        '
'''        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
'''        If Len(CD$) < 2 Then
'''          Call MENSERR(24, "Código de Familia Menor a dos Caracteres")
'''          GoTo 10
'''        End If
'''        ' Valida que la familia posea descripcion
'''        DESC$ = TRIM$(Mid$(X$, 3, 27))
'''        If Len(DESC$) < 1 Then
'''          Call MENSERR(24, "Codigo de Familia sin Descripción")
'''          GoTo 10
'''        End If
'''        ' Valida que la familia posea descripcion
'''        For U2& = 1 To U& - 1
'''          Y$ = REGLEIDO$("!TAFAMIL", U2&)
'''          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
'''            Call MENSERR(24, "Código de Familia Repetido")
'''            GoTo 10
'''          End If
'''        Next U2&
'''      Next U&
'''      Call TRACENTRAL("TAFAMIL", "")
'''    End If
'''    '
'''99  Command3.Enabled = True
'''    '
'''End Sub

Private Sub Command4_Click()
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

    Command4.Enabled = False
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0

    DoEvents
    Screen.MousePointer = 11
    Data1.DatabaseName = UCase$(CommonDialog1.FileName)
    On Error Resume Next
    Data1.Refresh
    Data1.RecordSource = Data1.Database.TableDefs(0).Name
    Data1.Refresh
    On Error GoTo 0
    
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        CODIGO$ = ""
        On Error Resume Next
        CODIGO$ = .Recordset.Fields(0)
        CI% = CODINT(CODIGO$)
        If CI% < 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Código '" & CODIGO$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        On Error GoTo 0
        TIPMAT% = .Recordset.Fields(2)
        Call ACTUREGISTRO("MASTER", "Timate", "Codint = " & CI% & "", TIPMAT%, REGAF&, "NOMESS")
20      Screen.MousePointer = 11
       .Recordset.MoveNext
     Loop
   End With
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
99 Command4.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = True
    OPX% = ALTERNA%("Sustitución de Códigos\Re-Impresion de Etiquetas")
    If OPX% = 1 Then
        '
        Call CREACAMPO("STOCKS", "LOTINGRE", "COD_EXTE", 10, 32)
        '
        CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
        CommonDialog1.DialogTitle = "Archivo Origen de Datos"
        '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
    
        DoEvents
        Screen.MousePointer = 11
        Data1.DatabaseName = UCase$(CommonDialog1.FileName)
        On Error Resume Next
        Data1.Refresh
        Data1.RecordSource = Data1.Database.TableDefs(0).Name
        Data1.Refresh
        On Error GoTo 0
        
        With Data1
         .Recordset.MoveFirst
         Do While Not .Recordset.EOF
            CODIGO$ = ""
            CODIVIE$ = TRIM$(SAFETEXT(.Recordset.Fields(1)))
            CODINUE$ = TRIM$(SAFETEXT(.Recordset.Fields(3)))
            DESCRINUE$ = SAFETEXT$(.Recordset.Fields(4))
            '
            If CODIVIE$ = "" Then GoTo 20
            '
            If CODINUE$ <> "" Then
              If CODINT%(CODINUE$) <> 0 Then
                Call MENSERR(24, "Destino Ocupado para Código " & CODINUE$ & ".")
                GoTo 20
              End If
            End If
            '
            CIXI% = CODINT%(CODIVIE$)
            If CIXI% = 0 Then
              Call MENSERR(24, "Origen Inexistente para Código " & CODIVIE$ & ".")
              GoTo 20
            End If
            '
            CONDI$ = "Codint = " & CIXI% & ""
            Call ACTUREGISTRO("LOTINGRE", "COD_EXTE", CONDI$, CODIVIE$, REGAF&)
            '
            If CODINUE$ <> "" Then
                Call ACTUREGISTRO("MASTER", "Codigo", CONDI$, CODINUE$, REGAF&)
                Call ACTUREGISTRO("MASTER", "Descripcion", CONDI$, DESCRINUE$, REGAF&)
              Else
                Call ACTUREGISTRO("MASTER", "STATUS", CONDI$, -1, REGAF&)
            End If
            '
20          Screen.MousePointer = 11
           .Recordset.MoveNext
         Loop
        End With
        Screen.MousePointer = 1
        '
        UPDATMASTER% = 1
        Call GENECOMAS
        Call AMASTOK07.ACTECOMAS     'Actualiza listados de seleccion de articulos
        '
        Call MENSERR(24, "Proceso de Substitucion de Codigo y Descripcion Finalizado")
        '
    ElseIf OPX% = 2 Then
        '
        Call APERBASDAT("Stocks")
        '
        SQLX$ = "SELECT * FROM LOTINGRE order by UBICFIS "
        '
        Dim LOTINGTMP As ADODB.Recordset
        Set LOTINGTMP = FLEXCONN.Execute(FILTSQL$(FILTSQL$(SQLX$)))
        With LOTINGTMP
           If Not (.EOF And .BOF) Then
           CONT = 0
             Do While .EOF = False
               LOTT$ = SAFETEXT(!IdenLote)
               CII% = XVALO(!Cod_Inte)
               CODEXVIE$ = SAFETEXT(!COD_EXTE)
               CODEXNUE$ = CODEXT$(CII%)
               If CODEXVIE$ = CODEXNUE$ Then GoTo 50
               Call ACSALOTE(CII%, LOTT$)
               SALD = SaldoLote(CII%, LOTT$)
               If SALD > 0 Then
                 Call PRINSTIKAPRO2(CII%, LOTT$)
                 CONDI$ = "Cod_inte = " & CII% & "and idenlote = '" & LOTT$ & "'"
                 Call ACTUREGISTRO("LOTINGRE", "COD_EXTE", CONDI$, CODEXNUE$, REGAF&)
                 CONT = CONT + 1
                   If CONT Mod 10 = 0 Then
                     OPP% = COMALTER%("Se Han Re-Impreso 10 Etiquetas\\Continuar\Cancelar")
                     If OPP% = 2 Then GoTo 80
                   End If
               End If
50           .MoveNext
             Loop
           End If
           Call MENSERR(24, "Proceso de Re-Ipresion de Etiquetas Finalizado")
        End With
80      LOTINGTMP.Close
        Set LOTINGTMP = Nothing
        
    End If
99  Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   Screen.MousePointer = 11
   Call ABREMASTER
   '
   
   SQLX$ = "SELECT * FROM MASTER "
   SQLX$ = SQLX$ + " Where TIMATE = 1"
   SQLX$ = SQLX$ + " AND SUBSTRING(CODIGO,1,1) <> 'P'"
   Dim MASTERTMP As ADODB.Recordset
   Set MASTERTMP = New ADODB.Recordset
   MASTERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With MASTERTMP
     Do While Not .EOF
        CODIGOEXT$ = SAFETEXT(!CODIGO)
        !NPLANO = CODIGOEXT$
        !CODIGO = "P" & CODIGOEXT$
        .Update
       .MoveNext
     Loop
   End With
   MASTERTMP.Close
   Set MASTERTMP = Nothing
   '
   UPDATMASTER% = 1
   Call GENECOMAS
   Call MENSERR(24, "Proceso Terminado")
   Command7.Enabled = True
   Screen.MousePointer = 1

End Sub

'
Private Sub Form_Load()
  '
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
  TABLA$(10) = "PLANTPRO"
  TABLA$(11) = "HERRAMEN"
  TABLA$(12) = "TAFAMIL"
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
  EPAC$ = TRIM$(UCase$(EMPREAC$)) 'BOTON PARA AGREGAR UNA 'P' AL CODIGO
  If InStr(EPAC$, "ACHERNAR") Then ' SOLICITADO POR ACHERNAR
    Command7.Visible = True
  End If

  Text12 = CONTROL$("AMASTO", "ATRIPAR1")
  Text3 = CONTROL$("AMASTO", "ATRIPAR2")
  Text4 = TRIM$(CONTROL$("MASTER", "LOMAXCOD"))
  '
  If CONTROL("MASTER", "DESLARGA") <> "SI" Then
    Option8.Value = True
  End If
  '
  If CONTROL("MASTER", "FORPATRO") = "SI" Then
      Check2.Value = 1
  End If
  '
  If CONTROL("MASTER", "M2PORUN") = "SI" Then
    Option11.Value = True
  End If
  '
  'CONTROLES DE REPTICION DE ATRIBUTOS
  If CONTROL("AMASTO", "REPATRIA") = "SI" Then
     Check4.Value = 1
  End If
  If CONTROL("AMASTO", "REPATRIB") = "SI" Then
     Check5.Value = 1
  End If

  
  APRO$ = CONTROL("FORMULAS", "REQAPRO")
  If APRO$ = "SI" Then Check3.Value = 1
  '
  Check6.Value = 1
  DOMIFI$ = CONTROL("MASTER", "DOMIFISC")
  If DOMIFI$ = "NO" Then Check6.Value = 0
  '
  Call APLICACIONSKINS(Me, Picture1)
  '
End Sub

'
 Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 12
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
    Text1.TEXT = TABLA(Index)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    Command16.Enabled = True
    On Error Resume Next
    Command16.SetFocus 'No se puede asignar el foco a un objeto no visible.
    On Error GoTo 0
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
    ElseIf Index = 10 Then
       Call CARPLANTA
    ElseIf Index = 11 Then
       Call CARHERRAMENT
    ElseIf Index = 12 Then
       Call CARFAMILIA
   End If
   '
End Sub
'
Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
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
    OPX% = COMALTER%("Opciones de Depósitos\\Editar Tabla de Depósitos\ABM de Mail por Depósito")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 2 Then
       Call ABMDEPOMAIL
       Exit Sub
    End If
    Call BLANARCH("!TADEPOSI")
    j& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If j& < 62 Then
                 j& = j& + 1
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
10  VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    ' validacion de planta asignada a deposito
    If EXISTE%(LUDAT$ + "PLANTPRO" + ".RFS") > 0 Then
      If ULTREG&("PLANTPRO") > 0 Then ' si tiene datos la tabla
        For U& = 1 To ULTREG&("!TADEPOSI") ' recorro la tabla
          TX$ = REGLEIDO$("!TADEPOSI", U&)
          PLANTA$ = Mid$(TX$, 57, 4)
          If ECOARCH$("PLANTPRO", PLANTA$) = "" Then
             DEPO% = Asc(Mid$(TX$, 1, 1)) ' si no tiene planta presenta mensaje y vuelve
             Call COMUNI("Planta Incorrecta Para Depósito: " & DEPO%)
             GoTo 10
          End If
        Next U&
      End If
    End If

    j& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If j& < 62 Then
               j& = j& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, j&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, j&)
               Call GRAREG("TADEPOSI", X$, j&)
             End If
           End If
        End If
    Next U&
    '
    '
    Call SETULTREG("SUCURSAL", j&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", j&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    j& = j& + 1
    Call GRAREG("ECODEP", Y$, j&)
    Call SETULTREG("ECODEP", j&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    j& = 0
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
    j& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            j& = j& + 1
            Call GRAREG("SECPROD", X$, j&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", j&)
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
           '
           EKO$ = ECOARCH("UMEDIDA", Mid$(X$, 33, 4)) 'CARGA DE FACTOR A UNIDAD DE MEDIDA
           If EKO$ <> "" Then
             If CVS(Mid$(X$, 37, 4)) < 1 Then
                Call MENSERR(24, "Factor No Válido")
                GoTo 10
             End If
           End If
           '
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


Sub PRINSTIKAPRO2(CI%, LOTE$, Optional DEPO%)

    FORIPRE$ = "STKA-AC3" ' TRIM$(CONTROL$("", "STIKAPRO"))
    If FORIPRE$ = "" Then Exit Sub
    '
    Call ACSALOTE(CI%, LOTE$, "NOMESS")
    '
    PFAC$ = CONTROL$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "AND IdenLote = '" & LOTE$ & "' "
    If DEPO% > 0 Then
       SQLX$ = SQLX$ + "AND DepoUbic =  " & DEPO% & ""
    End If
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With LOTINGTEMP
       On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         
         CODIGO$ = CODEXT$(!Cod_Inte)
         DESXRIP$ = DESCRIT0$(!Cod_Inte)
         LOTEE$ = SAFETEXT$(!IdenLote)
         Deposito% = XVALO(!DEPOUBIC)
         UBIFISICA$ = SAFETEXT$(!UBICFIS)
         LARGO# = XVALO(!MetroIni)
         ancho# = XVALO(!AnchoMili)
           FF% = CVINT(!FechAlta)
         FECHARECEP$ = FECHATEX$(FF%)
         '
         'Cantidad# = !CanAlta
         ' esto se modifico a pedido de bianchi para que aparezca el saldo
         CANTINIC# = XVALO(!CanAlta)
         CANTICON# = XVALO(!CanConsu)
         CANTISAL# = CANTINIC# - CANTICON#: If CANTISAL# < 0 Then CANTISAL# = 0
         If CANTINIC# > 0 Then
            If ancho# > 0 Then LARGO# = 1000 * CANTINIC# / ancho# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
            LARGO# = LARGO# * CANTISAL# / CANTINIC# ' para que aparezca el saldo
         End If
         '
      End If
      On Error GoTo 0
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing

    If CONTROL("BOLRECEP", "FORMATO") = "TXT" Then
        TXTX$ = TRIM$(FORIPRE$)
        If TRIM$(FORIPRE$) <> "" Then
            GoSub PRINZEBRA
          Else
            Call MENSERR(24, "Falta Definir Archivo Transferencia de Impresión")
        End If
        GoTo 90
    End If
    '
    TIPODOC$ = "Etiqueta Identificador de Lote"
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODIGO$
    CODPARAM$(2) = "NRO-LOTE": DOCPARAM$(2) = LOTEE$
    CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = Deposito%
    CODPARAM$(4) = "UBI-MPRI": DOCPARAM$(4) = UBIFISICA$
    CODPARAM$(5) = "MEDIDA-1": DOCPARAM$(5) = FORMATNUM$(LARGO#, "F15.4")
    CODPARAM$(6) = "MEDIDA-2": DOCPARAM$(6) = FORMATNUM$(ancho#, "F15.4")
    CODPARAM$(7) = "CANTIDAD": DOCPARAM$(7) = FORMATNUM$(CANTISAL#, "F15.4")
    CODPARAM$(8) = "REF-MAT1": DOCPARAM$(8) = AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12)
    CODPARAM$(9) = "DES-MAT": DOCPARAM$(9) = DESXRIP$
    CODPARAM$(10) = "MEDIDA-3": DOCPARAM$(10) = AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12)
    CODPARAM$(11) = "FECH-REC": DOCPARAM$(11) = FECHARECEP$
    CODPARAM$(12) = "REF-MAT2": DOCPARAM$(12) = VALOBADA("MOVISTO", "COD_EXTE", "IDENLOTE = '" & LOTEE$ & "'")

    CAPARDOC = 12
    Call PRINTDOC("@" + FORIPRE$)
    '
90  Call SETPRINPORT("RESTORE")
    Exit Sub
    '
PRINZEBRA:
    '
    TXTX$ = LOADFILE(LUDAT$ + FORIPRE$ + ".TXT")
    Call REPLATEXT(TXTX$, Chr$(13) + Chr$(10), Chr$(10))
    Call REPLATEXT(TXTX$, "@CODI-MAT", CODIGO$)
    Call REPLATEXT(TXTX$, "@NRO-LOTE", LOTEE$)
    Call REPLATEXT(TXTX$, "@DEPOSITO", Str$(Deposito%))
    Call REPLATEXT(TXTX$, "@UBI-MPRI", UBIFISICA$)
    Call REPLATEXT(TXTX$, "@MEDIDA-1", FORMATNUM(LARGO#, "F10.1"))
    Call REPLATEXT(TXTX$, "@MEDIDA-2", FORMATNUM(ancho#, "F10.1"))
    Call REPLATEXT(TXTX$, "@CANTIDAD", FORMATNUM(Cantidad#, "F10.1"))
    Call REPLATEXT(TXTX$, "@DESC-MAT", DESXRIP$)
    Call REPLATEXT(TXTX$, "@REF-MAT1", AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12))
    '
    TXTX1$ = TXTX$
    ARCHOUT$ = LUCOM$ + FORIPRE$ + ".PRN"
    Call SAVEFILE(ARCHOUT$, TXTX$)
    'PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    'Call DOSCOMMAND("PRINT " + ARCHOUT$)
    '
75  On Error Resume Next
    Printer.Print TXTX1$
    If Err Then
       On Error GoTo 0
       OPX% = COMALTER%("Error de la Impresora\\Re-Intentar\Abandonar")
       If OPX% = 1 Then GoTo 75
    End If
    On Error GoTo 0
    '
Return
'
End Sub


Sub CARPLANTA()
  '
  Call TRALOCAL("PLANTPRO", "PLANTPRO")
  '
'  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1900))
'  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("PLANTPRO/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!PLANTPRO")
        X$ = REGLEIDO$("!PLANTPRO", U&)
        JJ& = JJ& + 1
        Call GRAREG("PLANTPRO", X$, JJ&)
     Next U&
     Call SETULTREG("PLANTPRO", JJ&)
     Call CIERRARCH("PLANTPRO")
  End If
  '
End Sub

Sub CARHERRAMENT()
    
    Call SETULTREG&("!HERRAMEN", 0)
    j& = 0
    SQLX$ = "SELECT * FROM HERRAMEN WITH(NOLOCK)"
    Dim HERRTMP As ADODB.Recordset
    Set HERRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With HERRTMP
      Do While Not .EOF
        CODIG$ = SAFETEXT$(!CODIGO)
        DESCRX$ = SAFETEXT$(!Descripcion)
        If CODIG$ <> "" Then
        Call REPLA$(X$, CODIG$, 1, 16)
        Call REPLA$(X$, DESCRX$, 17, 64)
        j& = j& + 1
        Call GRAREG("!HERRAMEN", X$, j&)
        End If
        .MoveNext
        
      Loop
    End With
    '
    VTB% = VENTABU%("HERRAMEN")
    If VTB% < 0 Then Exit Sub
    '
    j& = 0
    CONDI$ = "CODIGO <> ''"
    Call ACTUREGISTRO("HERRAMEN", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    
    SQLX$ = "SELECT * FROM HERRAMEN "
    Dim HERRAMENTMP As ADODB.Recordset
    Set HERRAMENTMP = New ADODB.Recordset
25  On Error Resume Next
    HERRAMENTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

    For U& = 1 To ULTREG&("!HERRAMEN")
        X$ = REGLEIDO$("!HERRAMEN", U&)
        CODIG$ = TRIM$(Left$(X$, 16))
        DESCRX$ = TRIM$(Mid$(X$, 17, 64))
        If CODIG$ <> "" Then
            With HERRAMENTMP
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                !CODIGO = CODIG$
                !Descripcion = DESCRX$
                !MARBORRA = 0
                .Update
            End With
        End If
    Next U&
    HERRAMENTMP.Close
    Set HERRAMENTMP = Nothing
    '
    CONDI$ = "MARBORRA = 1"
    Call BORRAREGISTROS("HERRAMEN", CONDI$, REGAF&, "NOMESS")
    Call TRACENTRAL("HERRAMEN", "")
    Call SETULTREG("!HERRAMEN", j&)
    Call CIERRARCH("!HERRAMEN")

End Sub


