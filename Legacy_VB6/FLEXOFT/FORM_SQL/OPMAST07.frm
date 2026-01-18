VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPMAST07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   0
      Pattern         =   "*.XYZ"
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   3840
      TabIndex        =   69
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   67
      Text            =   "OPMAST07.frx":0000
      ToolTipText     =   "Doble Clic para seleccionar Ruta"
      Top             =   9240
      Width           =   5880
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   65
      Text            =   "OPMAST07.frx":0002
      ToolTipText     =   "Doble Clic para seleccionar Ruta"
      Top             =   8400
      Width           =   5880
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00CDDADA&
      Caption         =   "REVISION"
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
      Left            =   5000
      TabIndex        =   60
      Top             =   6270
      Width           =   1100
      Begin VB.CheckBox Check6 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Actualiza "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   70
         TabIndex        =   61
         Top             =   240
         Width           =   950
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3840
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   57
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONVERSION"
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
      Left            =   120
      TabIndex        =   49
      Top             =   6270
      Width           =   1455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "F. por Millar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   270
         Width           =   1250
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Form.Patron"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   50
         Top             =   550
         Width           =   1250
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00CDDADA&
      Caption         =   "COD."
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
      Left            =   3120
      TabIndex        =   46
      Top             =   6270
      Width           =   720
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   120
         MaxLength       =   2
         TabIndex        =   47
         Top             =   435
         Width           =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "UNIDAD"
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
      Left            =   1650
      TabIndex        =   43
      Top             =   6270
      Width           =   1395
      Begin VB.OptionButton Option12 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "M/U."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   720
         TabIndex        =   64
         Top             =   360
         Width           =   620
      End
      Begin VB.OptionButton Option10 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Kg/U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   80
         TabIndex        =   45
         Top             =   210
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton Option11 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "M2/U."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   80
         TabIndex        =   44
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00CDDADA&
      Caption         =   "DESCRIP."
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
      Left            =   3900
      TabIndex        =   40
      Top             =   6270
      Width           =   1035
      Begin VB.OptionButton Option9 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Larga"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   260
         Value           =   -1  'True
         Width           =   850
      End
      Begin VB.OptionButton Option8 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Doble"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   41
         Top             =   510
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
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMULA"
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
      Left            =   3360
      TabIndex        =   33
      Top             =   7215
      Width           =   2735
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Requiere Aprobar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1320
         TabIndex        =   52
         Top             =   240
         Width           =   1170
      End
      Begin VB.OptionButton Option7 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   240
         TabIndex        =   35
         Top             =   465
         Width           =   885
      End
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Uso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   240
         TabIndex        =   34
         Top             =   210
         Value           =   -1  'True
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "COMPLEMENTO"
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
      Left            =   1720
      TabIndex        =   25
      Top             =   7215
      Width           =   1575
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Nro. Plano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   27
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Código Alias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   26
         Top             =   465
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   5280
      ScaleHeight     =   6105
      ScaleWidth      =   1050
      TabIndex        =   20
      Top             =   0
      Width           =   1080
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   62
         Top             =   5520
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   12000
         End
      End
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   23
         ToolTipText     =   "A-B-M Familias de Articulos"
         Top             =   1800
         Visible         =   0   'False
         Width           =   640
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
         Picture         =   "OPMAST07.frx":0004
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
         Left            =   120
         Picture         =   "OPMAST07.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar "
         Top             =   960
         Width           =   640
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPOSICION"
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
      Left            =   120
      TabIndex        =   17
      Top             =   7215
      Width           =   1530
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Prov. Habit."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   19
         Top             =   465
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Criterio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         TabIndex        =   18
         Top             =   220
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "TABLAS DE VALIDACION"
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
      Top             =   120
      Width           =   4875
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "ATRIBUTOS PARAMETRICOS"
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
         Left            =   210
         TabIndex        =   28
         Top             =   4880
         Width           =   4455
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "No Repite"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   400
            Left            =   3150
            TabIndex        =   56
            Top             =   700
            Width           =   1200
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "No Repite"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   400
            Left            =   3150
            TabIndex        =   55
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
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
            Top             =   700
            Width           =   1875
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
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
            TabIndex        =   29
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   32
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   450
            Width           =   1020
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "CRITERIOS Y TABLAS GENERALES"
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
         Left            =   210
         TabIndex        =   7
         Top             =   1250
         Width           =   4455
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Familias"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   12
            Left            =   435
            TabIndex        =   59
            Top             =   1295
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Grupos de Venta según Tasa I.V.A."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   4
            Left            =   435
            TabIndex        =   24
            Top             =   795
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Unidades de Medida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   9
            Left            =   435
            TabIndex        =   14
            Top             =   1050
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Tabla de Tipos de Material"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   435
            TabIndex        =   9
            Top             =   240
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Grupos Contables de Venta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
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
         Left            =   3495
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
         Height          =   765
         Left            =   3920
         Picture         =   "OPMAST07.frx":0458
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   210
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3465
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   3
         Text            =   " "
         Top             =   300
         Width           =   1695
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONFIGURACION INDUSTRIAL"
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
         Left            =   210
         TabIndex        =   2
         Top             =   3000
         Width           =   4455
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Herramentales"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   450
            TabIndex        =   54
            Top             =   1240
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Plantas Productivas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   10
            Left            =   450
            TabIndex        =   53
            Top             =   220
            Width           =   3795
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Categorias de Operarios Productivos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   8
            Left            =   450
            TabIndex        =   13
            Top             =   1440
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Máquinas y Equipos de Producción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   7
            Left            =   450
            TabIndex        =   12
            Top             =   970
            Width           =   3690
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Secciones y Centros de Costos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   450
            TabIndex        =   11
            Top             =   730
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Depósitos y Almacenes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
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
         Caption         =   "Tabla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -495
         TabIndex        =   6
         Top             =   415
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
      OleObjectBlob   =   "OPMAST07.frx":0762
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "OPMAST07.frx":0996
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta de Archivo  Dropbox Instrucciones Tabla Master (Quilmes)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   -120
      TabIndex        =   68
      Top             =   9000
      Width           =   6285
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta de Archivo  Dropbox Instrucciones Tabla Master (Cambaceres)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   -120
      TabIndex        =   66
      Top             =   8120
      Width           =   6285
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTablasAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuAltaPorFamilia 
         Caption         =   "Alta por Familia"
      End
      Begin VB.Menu EliminaItemBaja 
         Caption         =   "Eliminaciòn de Items dados de Baja (Irreversible)"
      End
      Begin VB.Menu SincroMaster 
         Caption         =   "Editar Ruta de Archivo de Drop Para Sincronizar Master"
         Visible         =   0   'False
      End
      Begin VB.Menu actMaster 
         Caption         =   "Actualizar Tabla Master Segun Archivos "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCodRango 
         Caption         =   "Configuración de Códigos Rangos"
      End
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
    MAILX$ = Control("DIREMAIL", "DEPO-" & NUMEDEP$)
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

Sub APAREA_TABLA_FAMILIASQL()
   If EXISTE_TABLA%("DATASQL") = 1 And InStr(1, EMPREAC$, "ENVAPLAST") > 0 Then
      'MARCO CON MARBORRA
      DESCRITABL$ = "MATERIALES ENVAPLAST"
      IDLI& = XVALO(VALOBADA("DATASQL", "ID_LISTA", "LISTA = '" & DESCRITABL$ & "'", "NOMESS"))
      Call ACTUREGISTRO("DATASQL", "MARBORRA", "ID_LISTA = " & IDLI&, 1, REGAF&)
      
      For U& = 1 To ULTREG&("TAFAMIL")
        X$ = REGLEIDO$("TAFAMIL", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        DESC$ = TRIM$(Mid$(X$, 3, 27))
        VALOX$ = "" & IDLI& & " , '" & DESCRITABL$ & "'"
        VALOX$ = VALOX$ & ", " & CD$ & " , '" & DESC$ & "'"
        SQLX$ = "INSERT INTO DATASQL "
        PRICAMP$ = "ID_LISTA, LISTA, Campo1, Campo2"
        SQLX$ = SQLX$ + " (" + PRICAMP$ + RET_CAMPOS$ + ") "
        SQLX$ = SQLX$ + " VALUES (" & VALOX$ & ")"
        FLEXCONN.Execute SQLX$
      Next U&
      Call BORRAREGISTROS("DATASQL", "ID_LISTA = " & IDLI& & " AND MARBORRA = 1", REGAF&, "NOMESS")
   End If
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
      Call APAREA_TABLA_FAMILIASQL
      
    End If
End Sub

Private Sub actMaster_Click()
   Call ActualizarMasterDesdeArchivos
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
  opt$ = "NO"
  If Check4.Value = 1 Then opt$ = "SI"
  Call GRACONTROL("AMASTO", "REPATRIA", opt$)
  '
  opt$ = "NO"
  If Check5.Value = 1 Then opt$ = "SI"
  Call GRACONTROL("AMASTO", "REPATRIB", opt$)
  '
  
  IDSECUN$ = "NUPLANO": INCLUALIAS% = (-1)
  If Option4.Value = True Then
    IDSECUN$ = "ALIAS"
    INCLUALIAS% = 1
  End If
  Call GRACONTROL("", "IDSECUN", IDSECUN$)
  '
  LOMANT% = XVALO(Control$("MASTER", "LOMAXCOD"))
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
  MEPORUN$ = ""
  If Option12.Value = True Then MEPORUN$ = "SI"
  Call GRACONTROL("MASTER", "METLINEA", MEPORUN$)

  APRO$ = "NO"
  If Check3.Value = 1 Then APRO$ = "SI"
  Call GRACONTROL("FORMULAS", "REQAPRO", APRO$)
  '
  APRO$ = "NO"
  If Check6.Value = 1 Then APRO$ = "SI"
  Call GRACONTROL("MASTER", "ACTUREVI", APRO$)
  '
  Call GRACONTROL("MASTER", "DROPMAST", Text5.TEXT)
  Call GRACONTROL("MASTER", "DROPMAS2", Text6.TEXT)

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
        Codigo$ = ""
        On Error Resume Next
        Codigo$ = .Recordset.Fields(0)
        Ci% = CODINT%(Codigo$)
        If Ci% < 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente.\Dar de Alta y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        On Error GoTo 0
        TIPMAT% = .Recordset.Fields(2)
        Call ACTUREGISTRO("MASTER", "Timate", "Codint = " & Ci% & "", TIPMAT%, REGAF&, "NOMESS")
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
            Codigo$ = ""
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
               LOTT$ = SAFETEXT(!idenlote)
               CII% = XVALO(!cod_inte)
               CODEXVIE$ = SAFETEXT(!Cod_Exte)
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
        CODIGOEXT$ = SAFETEXT(!Codigo)
        !NPLANO = CODIGOEXT$
        !Codigo = "P" & CODIGOEXT$
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

Private Sub Command8_Click()
End Sub

Private Sub EliminaItemBaja_Click()
    OPX% = COMALTER%("Atenciòn Este Proceso Elimina Fisicamente\los Items que se Encuentran Dados de Baja\Esta Eliminaciòn Es Irreversible\\Cancelar\Generar Proceso de Eliminaciòn (Hasta 1000 Còdigos por Ejecuciòn)")
    If OPX% < 2 Then Exit Sub
    Call ProcesoEliminacionItemsDadosdeBajaAchernar
End Sub
Sub ProcesoEliminacionItemsDadosdeBajaAchernar()
    'RECORRER LOS ITEMS DADOS DE BAJA
    
    If DERACCE%("ELIMMAST", "ELIMINA FISICAMENTE ITEMS DADOS DE BAJA") < 2 Then Exit Sub
    CI1% = 0
    SQLX$ = "SELECT TOP 1000 CODINT  FROM MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STATUS < 0 "
    SQLX$ = SQLX$ + " ORDER BY CODINT "
    Set RS = READSET(SQLX$)
    With RS
    Do While Not .EOF
       CI1% = XVALO(!CODINT)
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       Text2 = CI1%: Text2.Refresh
       
       CONDI$ = "Cod_Inte = " & CI1%
       SQLX$ = "DELETE FROM BOLRECEP WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM DOCUREL WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM LOTINGRE WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM MACONSIG WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM MOVISTO WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM OCOMDETA WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM ORDEFABR WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM PLADIASTO WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       
       CONDI$ = "Codi_Conj = " & CI1%
       SQLX$ = "DELETE FROM RENCONSIG WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       
       CONDI$ = "Codi_Elem = " & CI1%
       SQLX$ = "DELETE FROM RENCONSIG WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       
       CONDI$ = "CodiElem = " & CI1%
       SQLX$ = "DELETE FROM FORMULAS WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)

       CONDI$ = "CodIConj =  " & CI1%
       SQLX$ = "DELETE FROM Formulas WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM SECOPER WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
      
       CONDI$ = "CODINT = " & CI1%
       SQLX$ = "DELETE FROM NOVEDADES WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       SQLX$ = "DELETE FROM MASTER WHERE " + CONDI$: FLEXCONN.Execute (SQLX$)
       FLEXCONN.CommitTrans
      .MoveNext
    Loop
    End With
    If CI1% < 1 Then
       Call COMUNI("No se Han Encontrado Còdigos Dados de Baja")
       GoTo 99
    End If
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION (ULTIMO CODIGO: " & CI1% & " )" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    End If
    
    Call COMUNI("TRANSACCIÓN REALIZADA (ULTIMO CODIGO: " & CI1% & ")")
99  Text2 = ""
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
  If Control$("", "MOREPOSI") = "PROHABIT" Then
    Option3.Value = True
  End If
  '
  If Control$("", "IDSECUN") = "ALIAS" Then
    Option4.Value = True
  End If
  '
  If Control$("AMASTO", "FORMULA") = "CANTIDAD" Then
    Option7.Value = True
  End If
  '
  If Control$("AMASTO", "PORMILES") = "SI" Then
    Check1.Value = 1
  End If
  '
  EPAC$ = TRIM$(UCase$(EMPREAC$)) 'BOTON PARA AGREGAR UNA 'P' AL CODIGO
  If InStr(EPAC$, "ACHERNAR") Then ' SOLICITADO POR ACHERNAR
    Command7.Visible = True
  End If
  If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
     SincroMaster.Visible = True
     actMaster.Visible = True
  End If
   
  Text12 = Control$("AMASTO", "ATRIPAR1")
  Text3 = Control$("AMASTO", "ATRIPAR2")
  Text4 = TRIM$(Control$("MASTER", "LOMAXCOD"))
  '
  If Control("MASTER", "DESLARGA") <> "SI" Then
    Option8.Value = True
  End If
  '
  If Control("MASTER", "FORPATRO") = "SI" Then
      Check2.Value = 1
  End If
  '
  If Control("MASTER", "M2PORUN") = "SI" Then
    Option11.Value = True
  End If
  '
  If Control("MASTER", "METLINEA") = "SI" Then
    Option12.Value = True
  End If
  '
  'CONTROLES DE REPTICION DE ATRIBUTOS
  If Control("AMASTO", "REPATRIA") = "SI" Then
     Check4.Value = 1
  End If
  If Control("AMASTO", "REPATRIB") = "SI" Then
     Check5.Value = 1
  End If
  
  If Control("MASTER", "ACTUREVI") = "SI" Then
     Check6.Value = 1
  End If
  
  APRO$ = Control("FORMULAS", "REQAPRO")
  If APRO$ = "SI" Then Check3.Value = 1
  '
  Text5.TEXT = Control("MASTER", "DROPMAST") 'ruta de drop box para dejar alli archivos.
  Text6.TEXT = Control("MASTER", "DROPMAS2") 'ruta de drop box para que en quilmes lea los archivos y ejecute los script
  
  Call APLICACIONSKINS(Me, Picture1)
  '
End Sub

Sub ActualizarMasterDesdeArchivos(Optional MODO$)
        '
        '________________________________________________________________________________
        'ESTA MISMA RUTINA ESTA DUPLICADA EN PROYECTO FTBOLQUI DE QUILMES
        '________________________________________________________________________________
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "QUILMES") < 1 Then Exit Sub
        '__________________________________________________________________________________________
        'ESTO ES PARADO EN Quilmes para recorrer los archivos y actualizar el Master
        '__________________________________________________________________________________________
        
        CARPETASINCRO$ = TRIM$(Control("MASTER", "DROPMAS2"))
        If Dir(CARPETASINCRO$, vbDirectory) = "" Then
           Call COMUNI("Falta Configuar Carpeta de Sincronizaciòn")
           Exit Sub
        End If
        '
        If CARPETASINCRO$ <> "" Then
            'ASIGNO PROPIEDADES AL FILE1
                On Error Resume Next
                File1.Path = "C:\"
                On Error GoTo 0
                File1.Path = CARPETASINCRO$
                File1.Pattern = "*.txt"
                TX$ = ""
                List3.Clear
                'RECORRO EL FILE1 Y LE PONGO LA FECHA DEL ARCHIVO POR DELANTE PARA QUE ORDENE CORRECTAMENTE
                For n = 0 To File1.ListCount - 1
                  NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List3.AddItem TX$
                Next n
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For i& = 0 To List3.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List3.List(i&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List3.List(i&), POS1% + 2))
                  If Left$(TRIM$(UCase$(NombreArchivo$)), 11) = "ALTACODIGO_" Then
                    OK% = 0
                    Call EjecutarALtaMasterQuilmes(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt", OK%)
                    Cant_ItemsActualizados& = Cant_ItemsActualizados& + OK% 'ok devuelve 1 si actualizo, si no devuelve 0
                  End If
                Next i&

        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            Exit Sub
        End If
        
        If MODO$ <> "NOMESS" Then
           Call COMUNI("Proceso Terminado\Se Han Actualizado " & Cant_ItemsActualizados& & " Items")
        End If
        '

End Sub


Private Sub mnuAltaPorFamilia_Click()

    Dim ta As New DatosAuxiliares

    Call ta.MostrarTablaAuxilar("TAFORALFA")

End Sub

Private Sub mnuCodRango_Click()
        If DERACCE("CONCODRA", "CONFIGURACION DE CODIGOS RANGOS") < 2 Then Exit Sub
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "CONFIGURACIÓN CODIGOS RANGO"
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER;Timate  = 49"
        VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER;Timate  <> 49"
        VENTANA.Inicializar = 1
        Campos$ = "Cód.Rango/A16;" '1
        Campos$ = Campos$ + "Descripción/A64;"      '2
        Campos$ = Campos$ + "Cód.M.P./A16;"           '3
        Campos$ = Campos$ + "DESCRIPCION/A64;"      '4
        Campos$ = Campos$ + "CIR/A0;"       '5
        Campos$ = Campos$ + "CIMP/A0;"       '6
        Campos$ = Campos$ + "ID/A0" '7
        VENTANA.Campox = Campos$
        Dim obj As New RECORXET
        Set RS1 = obj.RS_CODRANGO()
        Call Carga_MSF_Recordset(RS1, Campos$, VENTABNEW.MSF)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)
'
20      VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub


        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        CONDI$ = "CODIGORANGO_ID > 0"
        Call ACTUREGISTRO("CODIGORANGO", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        SQLX$ = "SELECT * FROM CODIGORANGO"
        SQLX$ = SQLX$ + " WHERE CODIGORANGO_ID < 1"
        Dim MSF As msFlexGrid
        Set MSF = VENTABNEW.MSF
        
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With RS
              For REG& = 1 To MSF.Rows - 1
                  If TRIM$(MSF.TextMatrix(REG&, 1)) = "" Then GoTo 50
                  .AddNew
                  !CODRANGO = MSF.TextMatrix(REG&, 1)
                  !Descripcion = MSF.TextMatrix(REG&, 2)
                  !CODMATPRIMA = MSF.TextMatrix(REG&, 3)
                  !DESCRIMATPRIMA = MSF.TextMatrix(REG&, 4)
                  CIRGO = CODINT%(MSF.TextMatrix(REG&, 1))
                  !CIRANGO = CIRGO
                  MSF.TextMatrix(REG&, 5) = CIRGO ' SI NO LO TENIA LO ESCRIBE EN LA GRILLA PARA VALIDAR ABAJO
                  CIMP = CODINT%(MSF.TextMatrix(REG&, 3))
                  !CIMATPRIMA = CIMP
                  MSF.TextMatrix(REG&, 6) = CIMP ' SI NO LO TENIA LO ESCRIBE EN LA GRILLA PARA VALIDAR ABAJO
                  !MARBORRA = 0
                  .Update
50            Next REG&
        End With
        
        RS.Close
        Set RS = Nothing
        
        '
        'VALIDA AQUI SI COINCIDEN LAS U.MEDIDA (LO HAGO AQUI PARA NO REPETIR ANTES ASIGNAR EL COD INTERNO.)
        For REG& = 1 To MSF.Rows - 1
           CIRANGOX% = XVALO(MSF.TextMatrix(REG&, 5))
           CIMPX% = XVALO(MSF.TextMatrix(REG&, 6))
           If CIRANGOX% >= 0 And CIMPX% >= 0 Then
               If Unimed$(CIRANGOX%) <> Unimed$(CIMPX%) Then
                 Call COMUNI("Atención En Fila: " & REG& & " No Coincide Tipo de Material\" & MSF.TextMatrix(REG&, 1) & " = " & Unimed$(CIRANGOX%) & "\" & MSF.TextMatrix(REG&, 3) & " = " & Unimed$(CIMPX%))
                 FLEXCONN.RollbackTrans
                 On Error GoTo 0
                 GoTo 20
               End If
           End If
        Next REG&
        
        'VALIDO QUE EL CODIGO DE MATERIA PRIMA NO ESTE REPETIDO
        For REG& = 1 To MSF.Rows - 1
           CIRANGOX% = XVALO(MSF.TextMatrix(REG&, 5))
           CIMPX% = XVALO(MSF.TextMatrix(REG&, 6))
           For REG2& = 1 To MSF.Rows - 1
             CIMPX2% = XVALO(MSF.TextMatrix(REG2&, 6))
             If REG2& <> REG& And CIMPX2% = CIMPX% Then
               Call COMUNI("Atención En Fila: " & REG& & " Código: " & CODEXT$(CIMPX%) & " Repetido en Fila: " & REG2&)
               FLEXCONN.RollbackTrans
               On Error GoTo 0
               GoTo 20
             End If
           Next REG2&
        Next REG&
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        Call BORRAREGISTROS("CODIGORANGO", CONDI$, REGAF&, "NOMESS")
        
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
90     FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW
    
   
End Sub



Private Sub mnuTablasAuxiliares_Click()

    Dim ta As New DatosAuxiliares

    Call ta.ListarTodasLasTablasAuxiliares
        
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

Private Sub S1_Click()

End Sub

Private Sub SincroMaster_Click()
   Me.Height = 10560
   Call Centrar(Me)
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
      opt% = COMALTER%("FLEXOFT Está Diseñado para Procesar\Hasta 4 Tasas de IVA.\   \Los Registros Excedentes Pueden No Ser\Considerados en Algunos Procesos.\\Corregir\Aprobar")
      If opt% <> 2 Then
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
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 99 Then
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

    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 99 Then
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
    '
    '
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


Sub PRINSTIKAPRO2(Ci%, LOTE$, Optional DEPO%)

    FORIPRE$ = "STKA-AC3" ' TRIM$(CONTROL$("", "STIKAPRO"))
    If FORIPRE$ = "" Then Exit Sub
    '
    Call ACSALOTE(Ci%, LOTE$, "NOMESS")
    '
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & Ci% & " "
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
         
         Codigo$ = CODEXT$(!cod_inte)
         DESXRIP$ = DESCRIT0$(!cod_inte)
         LOTEE$ = SAFETEXT$(!idenlote)
         DEPOSITO% = XVALO(!DEPOUBIC)
         UBIFISICA$ = SAFETEXT$(!UbicFis)
         Largo# = XVALO(!MetroIni)
         ANCHO# = XVALO(!AnchoMili)
           FF% = CVINT(!FechAlta)
         FECHARECEP$ = FECHATEX$(FF%)
         '
         'Cantidad# = !CanAlta
         ' esto se modifico a pedido de bianchi para que aparezca el saldo
         CANTINIC# = XVALO(!CanAlta)
         CANTICON# = XVALO(!CanConsu)
         CANTISAL# = CANTINIC# - CANTICON#: If CANTISAL# < 0 Then CANTISAL# = 0
         If CANTINIC# > 0 Then
            If ANCHO# > 0 Then Largo# = 1000 * CANTINIC# / ANCHO# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
            Largo# = Largo# * CANTISAL# / CANTINIC# ' para que aparezca el saldo
         End If
         '
      End If
      On Error GoTo 0
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing

    If Control("BOLRECEP", "FORMATO") = "TXT" Then
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
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = Codigo$
    CODPARAM$(2) = "NRO-LOTE": DOCPARAM$(2) = LOTEE$
    CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSITO%
    CODPARAM$(4) = "UBI-MPRI": DOCPARAM$(4) = UBIFISICA$
    CODPARAM$(5) = "MEDIDA-1": DOCPARAM$(5) = FORMATNUM$(Largo#, "F15.4")
    CODPARAM$(6) = "MEDIDA-2": DOCPARAM$(6) = FORMATNUM$(ANCHO#, "F15.4")
    CODPARAM$(7) = "CANTIDAD": DOCPARAM$(7) = FORMATNUM$(CANTISAL#, "F15.4")
    CODPARAM$(8) = "REF-MAT1": DOCPARAM$(8) = AJUSTI$(Codigo$, 16) + AJUSTI$(LOTEE$, 12)
    CODPARAM$(9) = "DES-MAT": DOCPARAM$(9) = DESXRIP$
    CODPARAM$(10) = "MEDIDA-3": DOCPARAM$(10) = AJUSTI$(Codigo$, 16) + AJUSTI$(LOTEE$, 12)
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
    Call REPLATEXT(TXTX$, "@CODI-MAT", Codigo$)
    Call REPLATEXT(TXTX$, "@NRO-LOTE", LOTEE$)
    Call REPLATEXT(TXTX$, "@DEPOSITO", Str$(DEPOSITO%))
    Call REPLATEXT(TXTX$, "@UBI-MPRI", UBIFISICA$)
    Call REPLATEXT(TXTX$, "@MEDIDA-1", FORMATNUM(Largo#, "F10.1"))
    Call REPLATEXT(TXTX$, "@MEDIDA-2", FORMATNUM(ANCHO#, "F10.1"))
    Call REPLATEXT(TXTX$, "@CANTIDAD", FORMATNUM(Cantidad#, "F10.1"))
    Call REPLATEXT(TXTX$, "@DESC-MAT", DESXRIP$)
    Call REPLATEXT(TXTX$, "@REF-MAT1", AJUSTI$(Codigo$, 16) + AJUSTI$(LOTEE$, 12))
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
    J& = 0
    SQLX$ = "SELECT * FROM HERRAMEN WITH(NOLOCK)"
    Dim HERRTMP As ADODB.Recordset
    Set HERRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With HERRTMP
      Do While Not .EOF
        CODIG$ = SAFETEXT$(!Codigo)
        DESCRX$ = SAFETEXT$(!Descripcion)
        If CODIG$ <> "" Then
        Call REPLA$(X$, CODIG$, 1, 16)
        Call REPLA$(X$, DESCRX$, 17, 64)
        J& = J& + 1
        Call GRAREG("!HERRAMEN", X$, J&)
        End If
        .MoveNext
        
      Loop
    End With
    '
    VTB% = VENTABU%("HERRAMEN")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
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
                !CODIEMPR = CodiEmp%
                On Error GoTo 0
                !Codigo = CODIG$
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
    Call SETULTREG("!HERRAMEN", J&)
    Call CIERRARCH("!HERRAMEN")

End Sub


Private Sub Text5_DblClick()
     SELCARP07.Caption = "Seleccione Carpeta "
     SELCARP07.Show 1
     Text5 = SELCARP07.Label1

End Sub


Private Sub Text6_DblClick()
     SELCARP07.Caption = "Seleccione Carpeta "
     SELCARP07.Show 1
     Text6 = SELCARP07.Label1
End Sub


