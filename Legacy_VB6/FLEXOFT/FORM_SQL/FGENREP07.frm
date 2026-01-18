VERSION 5.00
Begin VB.Form FGENREP07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes          FLEXOFT"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Operadores Sql"
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
      Left            =   240
      TabIndex        =   77
      Top             =   4800
      Width           =   6135
      Begin VB.CheckBox chkUnionAll 
         Caption         =   "Union All"
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
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox CMBJUNTAR 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   75
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton Command14 
      Caption         =   ")"
      Height          =   480
      Left            =   4530
      TabIndex        =   74
      Top             =   6720
      Width           =   210
   End
   Begin VB.CommandButton Command15 
      Caption         =   "("
      Height          =   480
      Left            =   0
      TabIndex        =   73
      Top             =   6720
      Width           =   210
   End
   Begin VB.CommandButton Command12 
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   72
      Top             =   7440
      Width           =   210
   End
   Begin VB.CommandButton Command11 
      Height          =   560
      Left            =   8750
      Picture         =   "FGENREP07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Eliminar (Borrar) Consulta Activa"
      Top             =   105
      Width           =   500
   End
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
      Height          =   885
      Left            =   10815
      Picture         =   "FGENREP07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2370
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   9330
      ScaleHeight     =   120
      ScaleWidth      =   2385
      TabIndex        =   67
      Top             =   3360
      Width           =   2380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Presentacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   9310
      TabIndex        =   21
      Top             =   0
      Width           =   2460
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   210
         TabIndex        =   44
         Top             =   1095
         Width           =   1170
      End
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   79
         Text            =   "Combo6"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   70
         Top             =   3045
         Width           =   1065
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Distintos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   69
         Top             =   2825
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   64
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   1095
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1710
         TabIndex        =   46
         Top             =   1095
         Width           =   630
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Agrupa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   42
         Top             =   2605
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Totaliza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   41
         Top             =   2385
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         Caption         =   "Pos"
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
         Left            =   1160
         TabIndex        =   36
         Top             =   1425
         Width           =   1185
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
            Height          =   250
            Left            =   465
            TabIndex        =   39
            Top             =   560
            Width           =   600
         End
         Begin VB.TextBox Text3 
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
            Height          =   250
            Left            =   465
            TabIndex        =   37
            Top             =   260
            Width           =   600
         End
         Begin VB.Label Label11 
            Caption         =   "V:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   40
            Top             =   585
            Width           =   795
         End
         Begin VB.Label Label10 
            Caption         =   "H:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   38
            Top             =   285
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sort"
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
         Left            =   105
         TabIndex        =   32
         Top             =   1425
         Width           =   975
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Caption         =   "Des"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   35
            Top             =   620
            Width           =   645
         End
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            Caption         =   "Asc"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   34
            Top             =   420
            Width           =   645
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   33
            Top             =   220
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   645
         End
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   210
         TabIndex        =   25
         Top             =   550
         Width           =   2145
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         TabIndex        =   24
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Form:"
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
         Left            =   1605
         TabIndex        =   47
         Top             =   870
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Eco:"
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
         TabIndex        =   45
         Top             =   870
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Label:"
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
         TabIndex        =   23
         Top             =   320
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Form:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   22
         Top             =   265
         Width           =   795
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "X"
      Height          =   240
      Left            =   3400
      TabIndex        =   65
      ToolTipText     =   "Caracter Comodin"
      Top             =   720
      Width           =   260
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2130
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   4400
      Width           =   1575
   End
   Begin VB.ListBox LSTTABLASEC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      IntegralHeight  =   0   'False
      ItemData        =   "FGENREP07.frx":0614
      Left            =   1995
      List            =   "FGENREP07.frx":0616
      Sorted          =   -1  'True
      TabIndex        =   60
      Top             =   945
      Width           =   1665
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   4400
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "Configuracion Campo Auxiliar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6560
      TabIndex        =   52
      Top             =   3120
      Width           =   5220
      Begin VB.CommandButton Command7 
         Caption         =   "OK"
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
         Left            =   4515
         TabIndex        =   71
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         Sorted          =   -1  'True
         TabIndex        =   55
         Text            =   "CMBCAMPOS"
         Top             =   520
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   54
         Text            =   "CMBCAMPOS"
         Top             =   520
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Tabla:"
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
         TabIndex        =   58
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Campo:"
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
         Left            =   1680
         TabIndex        =   57
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label15 
         Caption         =   "Condicion:"
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
         TabIndex        =   56
         Top             =   885
         Width           =   1275
      End
   End
   Begin VB.ListBox CAMAUXI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      IntegralHeight  =   0   'False
      ItemData        =   "FGENREP07.frx":0618
      Left            =   3900
      List            =   "FGENREP07.frx":061F
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   50
      Top             =   3960
      Width           =   2505
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   240
      Left            =   105
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.TextBox TXTCANREGI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   7665
      Width           =   7050
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9915
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   29
      Top             =   5880
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   30
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      IntegralHeight  =   0   'False
      ItemData        =   "FGENREP07.frx":062C
      Left            =   220
      List            =   "FGENREP07.frx":062E
      TabIndex        =   26
      Top             =   6660
      Width           =   4280
   End
   Begin VB.CommandButton Command5 
      Height          =   540
      Left            =   8850
      Picture         =   "FGENREP07.frx":0630
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2310
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      Height          =   540
      Left            =   8850
      Picture         =   "FGENREP07.frx":093A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1155
      Width           =   330
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2120
      IntegralHeight  =   0   'False
      ItemData        =   "FGENREP07.frx":0C44
      Left            =   6555
      List            =   "FGENREP07.frx":0C46
      TabIndex        =   16
      Top             =   950
      Width           =   2145
   End
   Begin VB.CommandButton CMDBORRAR 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10920
      TabIndex        =   8
      Top             =   6285
      Width           =   850
   End
   Begin VB.CommandButton EJECUTAR 
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   9840
      TabIndex        =   7
      Top             =   6165
      Width           =   960
   End
   Begin VB.CommandButton GUARDACON 
      Caption         =   "Guardar Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8925
      TabIndex        =   6
      Top             =   5775
      Width           =   855
   End
   Begin VB.CommandButton CMDPEGAR 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10920
      TabIndex        =   5
      Top             =   6030
      Width           =   850
   End
   Begin VB.CommandButton CMDCOPIAR 
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10920
      TabIndex        =   4
      Top             =   5775
      Width           =   850
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   8280
      Picture         =   "FGENREP07.frx":0C48
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agregar Condición"
      Top             =   5760
      Width           =   540
   End
   Begin VB.TextBox TXTCONSULTA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4725
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6660
      Width           =   7050
   End
   Begin VB.ListBox LSTMOSTARCMP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      IntegralHeight  =   0   'False
      ItemData        =   "FGENREP07.frx":0F52
      Left            =   3900
      List            =   "FGENREP07.frx":0F54
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   945
      Width           =   2505
   End
   Begin VB.ListBox LSTTABLAS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      IntegralHeight  =   0   'False
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   945
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   5670
      Width           =   8595
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   300
         Left            =   7665
         TabIndex        =   18
         ToolTipText     =   "Caracter Comodin"
         Top             =   315
         Width           =   260
      End
      Begin VB.ComboBox CMBCAMPOS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   15
         Text            =   "CMBCAMPOS"
         Top             =   315
         Width           =   2790
      End
      Begin VB.ComboBox CMBOPERADOR 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FGENREP07.frx":0F56
         Left            =   3240
         List            =   "FGENREP07.frx":0F58
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "CMBOPERADOR"
         Top             =   315
         Width           =   1080
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
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
         Left            =   7275
         TabIndex        =   13
         ToolTipText     =   "Lista de Valores Posibles"
         Top             =   315
         Width           =   330
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   315
         Width           =   2715
      End
   End
   Begin VB.Label Label19 
      Caption         =   "Tipo.Join:"
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
      Left            =   240
      TabIndex        =   76
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label Label18 
      Caption         =   "Tabla Secundaria"
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
      Left            =   1770
      TabIndex        =   63
      Top             =   705
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   1785
      TabIndex        =   62
      Top             =   4425
      Width           =   330
   End
   Begin VB.Label Label2 
      Caption         =   "Tabla Principal:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   705
      Width           =   1920
   End
   Begin VB.Label Label14 
      Caption         =   "Auxiliares:"
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
      Left            =   3720
      TabIndex        =   51
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
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
      Left            =   1680
      TabIndex        =   49
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   960
      TabIndex        =   48
      Top             =   360
      Width           =   2325
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
      Height          =   555
      Left            =   3465
      TabIndex        =   28
      Top             =   105
      Width           =   5235
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Consulta"
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
      TabIndex        =   27
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   420
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Orden:"
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
      Left            =   5610
      TabIndex        =   17
      Top             =   705
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Campos:"
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
      Left            =   3150
      TabIndex        =   11
      Top             =   705
      Width           =   1275
   End
End
Attribute VB_Name = "FGENREP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KLIK%
Dim MODIFENCA%, ILIS1&

Private Sub CAMAUXI_Click()
  Call ARMASELFIELD
End Sub

Private Sub chkUnionAll_Click()

    LSTTABLASEC.ListIndex = -1
    UNIONALL$ = ""
    CMBJUNTAR.TEXT = ""
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    
    Call ARMA_CONSULTX
    
End Sub

Private Sub CMBJUNTAR_Click()
   Call ARMA_CONSULTX
End Sub


'
Private Sub CMDBORRAR_Click()
    TXTCONSULTA.TEXT = ""
End Sub

Private Sub Combo2_Click()
    Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & Combo2.TEXT))
    For i& = 1 To AUX.Fields.Count
        Combo3.AddItem AUX.Fields(i& - 1).Name
    Next i&
End Sub

Private Sub Combo6_Click()

    Text5 = Combo6.TEXT
    
End Sub

Private Sub Command10_Click()
'    Static CTX%
'    If CTX% < 1 Then
'       Call ARMALISTARCH(LUDAT$, "*.RFS")
'       CTX% = 1
'    End If
'    Call SELECHO("INDIRFS")   ' , "DIRECTORIO RFS")
'    VDEVU$ = TRIM$(VALACT1$("INDIRFS"))
'    If VDEVU$ <> "" Then
'       Text5 = VDEVU$
'    End If

End Sub

Private Sub Command11_Click()
    If Label1 <> "" Then
       If COMALTER%("¿ Confirma Eliminar la Consulta '" & TRIM$(Label1) & "' ?\  \(Esta Transaccion no Podrá Revertirse)\\No, Conservarla\Si, Eliminarla") = 2 Then
          CONDI$ = "COD_QUERY = '" & Label1.Caption & "'"
          Call BORRAREGISTROS("SQL_QUERIES", CONDI$, REGAF&, "NOMESS")
          
          Call ABRETACON
          strSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
          Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenDynaset)
          If Not (MIRS.EOF And MIRS.BOF) Then
             MIRS.Delete
             DEFTREP07.Label4 = ""
             DEFTREP07.Label9 = ""
             Call DEFTREP07.VerConsul
             MIRS.Close
             Unload Me
          End If
       End If
    End If
End Sub

Private Sub Command12_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 3) <> "OR " Then
          List2.List(i%) = "OR " & List2.TEXT
         End If
         Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command14_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Right(TRIM$(List2.TEXT), 1) <> ")" Then
          List2.List(i%) = List2.TEXT & ")"
        End If
        Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command15_Click()
    'INTERVENCION OR

    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 1) <> "(" Then
          List2.List(i%) = "(" & List2.TEXT
        End If
        Call ARMA_CONSULTX
    End If
End Sub



Private Sub Command7_Click()
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
      If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
        TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
        CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
        CONDIVALO$ = AJUSTI$(Text7, 512)
        'If TRIM$(TAVALO$) <> "" Then
        '  If TRIM$(CAVALO$) <> "" Then
            If TRIM$(CONDIVALO$) <> "" Then
              CAMPOSE$ = AJUSTI$(List1.TEXT, 168) + TAVALO$ + CAVALO$ + CONDIVALO$
              List1.RemoveItem (i&)
              List1.AddItem CAMPOSE$, i&
              List1.Selected(i&) = True
            End If
        '  End If
        'End If
      End If
    End If
    '
    For i& = 1 To LSTMOSTARCMP.ListCount
       If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          LSTMOSTARCMP.RemoveItem (i& - 1)
          LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
          LSTMOSTARCMP.Selected(i& - 1) = True
       End If
    Next i&
    '
    For i& = 1 To CAMAUXI.ListCount
       If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          CAMAUXI.RemoveItem (i& - 1)
          CAMAUXI.AddItem CAMPOSE$, i& - 1
          CAMAUXI.Selected(i& - 1) = True
       End If
    Next i&
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command9_Click()
    '
    LSTTABLASEC.ListIndex = (-1)
    Combo4.Clear
    Combo5.Clear
    TAPRIN$ = LSTTABLAS.TEXT
    If TAPRIN$ = "" Then Exit Sub
    '
10  For U& = 1 To LSTMOSTARCMP.ListCount
       If Left$(LSTMOSTARCMP.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          LSTMOSTARCMP.RemoveItem (U& - 1)
          GoTo 10
       End If
    Next U&
    '
20  For U& = 1 To List1.ListCount
       If Left$(List1.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          List1.RemoveItem (U& - 1)
          GoTo 20
       End If
    Next U&
    '
    Call LSTTABLAS_Click     ' PARA REFRESCAR EL COMBO4
    '
End Sub

Private Sub label1_Change()
    Dim FISELE$(128)
    List2.Clear
    
    CANFI% = 0
    ESTAENSQL% = 0
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK) where COD_QUERY = '" & Label1.Caption & "'"
    Set RS = READSET(SQLX$)
    If Not (RS.EOF And RS.BOF) Then
       Label9 = SAFETEXT$(RS!DES_QUERY)
       TTXX$ = SAFETEXT$(RS!PAR_QUERY)
       ESTAENSQL% = 1
    End If
    If ESTAENSQL% = 1 Then
      RS.Close
      Set RS = Nothing
      GoTo 30
    End If
    
    strSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_QUERY
       TTXX$ = MIRS!PAR_QUERY
       '
30     While Len(TTXX$) > 0
          PPXX% = InStr(TTXX$, Chr$(124))
          If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
          COMAI$ = Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 1)
          '
          If Left$(COMAI$, 10) = "DATABASE =" Then
              Call CMBBASE_CLICK
              CMBBASE = UCase$(TRIM$(Mid$(COMAI$, 11)))
              GoTo 19
            ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
              TASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLAS.ListCount
                 If UCase$(LSTTABLAS.List(KKI% - 1)) = TASELE Then
                   LSTTABLAS.Selected(KKI% - 1) = True
                   Call LSTTABLAS_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "SELTASEC =" Then
              TASELSEC$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLASEC.ListCount
                 If UCase$(LSTTABLASEC.List(KKI% - 1)) = TASELSEC Then
                   LSTTABLASEC.Selected(KKI% - 1) = True
                   Call LSTTABLASEC_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "MATCHFI1 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo4.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "MATCHFI2 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo5.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "SELFIELD =" Then
              CANFI% = CANFI% + 1
              FISELE$(CANFI%) = TRIM$(Mid$(COMAI$, 11))
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
              For KKI% = 1 To LSTMOSTARCMP.ListCount
                 If UCase$(TRIM$(Left$(LSTMOSTARCMP.List(KKI% - 1), 64))) = CASELE$ Then
                   LSTMOSTARCMP.RemoveItem KKI% - 1
                   LSTMOSTARCMP.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   LSTMOSTARCMP.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
              For KKI% = 1 To CAMAUXI.ListCount
                 If UCase$(TRIM$(Left$(CAMAUXI.List(KKI% - 1), 64))) = CASELE$ Then
                   CAMAUXI.RemoveItem KKI% - 1
                   CAMAUXI.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   CAMAUXI.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
              List2.AddItem TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "TIPOJOIN =" Then
              CMBJUNTAR.TEXT = TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "UNIONALL =" Then
              chkUnionAll.Value = 1
              SegundaTabla$ = TRIM$(Mid$(COMAI$, 11))
              For idTbSec% = 0 To LSTTABLASEC.ListCount - 1
                    If TRIM$(LSTTABLASEC.List(idTbSec%)) = SegundaTabla$ Then
                        LSTTABLASEC.ListIndex = idTbSec
                        UNIONALL$ = SegundaTabla$
                        TablaUnionOk% = 1
                        Exit For
                    End If
              Next idTbSec%
              If TablaUnionOk% = 0 Then
                chkUnionAll.Value = 0
                UNIONALL$ = ""
              End If
            Else
          End If
19     Wend
       '
       List1.Clear
       For KKI% = 1 To CANFI%
          List1.AddItem FISELE$(KKI%)
       Next KKI%
       '
       Call ARMA_CONSULTX
       Exit Sub
       '
    End If
    '
    'SI NO ESTABA EN SQL ENTRO EN MDB Y CIERRA EL RECORDSET
    If ESTAENSQL% = 0 Then
      MIRS.Close
      Set MIRS = Nothing
    End If

    Call LSTTABLAS_Click
    '
End Sub

Private Sub Command1_Click()
   If CMBCAMPOS.TEXT = "" Or CMBOPERADOR = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   If Len(TRIM$(CMBCAMPOS.TEXT)) > 24 Then
      MsgBox "Nombre de Campo Demasiado Largo"
      Exit Sub
   End If
   SQLY$ = AJUSTI$(CMBCAMPOS.TEXT, 24) + " " + AJUSTI$(CMBOPERADOR.TEXT, 8)
   VALOCO$ = TRIM$(Combo1.TEXT)
   '
   TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
   If (TPC = 52 Or TPC = 62) And VALOCO$ = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   '
   ' ESTO ES PARA QUE PERMITA COMPARAR VALORES DEL RS ENTRE SÍ
   For KKU& = 1 To CMBCAMPOS.ListCount
      If UCase$(CMBCAMPOS.List(KKU& - 1)) = UCase$(VALOCO$) Then GoTo 10
   Next KKU&
   '
   If TPC <> 52 And TPC <> 62 Then VALOCO$ = "'" + VALOCO$ + "'"
   '
10 SQLY$ = SQLY$ + " " + VALOCO$

   List2.AddItem SQLY$
   Call ARMA_CONSULTX
   Exit Sub
   
      If KLIK% > 0 Then GoTo 25
      SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " Where " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      KLIK% = KLIK% + 1
      TXTCONSULTA.TEXT = SQLX$
      Exit Sub
      
25    SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " AND " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      TXTCONSULTA.TEXT = SQLX$
      
End Sub

Function TIPODECAMPO(NTABLA$, NCAMPO$)
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    strSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    strSQL = strSQL & "sysobjects.id = syscolumns.id inner join systypes "
    strSQL = strSQL & "on syscolumns.xtype = systypes.xtype "
    strSQL = strSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    strSQL = strSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO = MIRS!xtype
    End If

End Function

Private Sub CMDCOPIAR_Click()
   Clipboard.SetText Mid$(TXTCONSULTA.TEXT, TXTCONSULTA.SelStart + 1, Text1.SelLength)
End Sub

Private Sub CORTAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

Private Sub Command2_Click()
   If CMBCAMPOS.TEXT <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CMBCAMPOS.TEXT & " AS VALOR FROM " & LSTTABLAS.TEXT & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CMBCAMPOS.TEXT & " ASC"
      '
      Combo1.Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      On Error Resume Next
      Set VAPOSI = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Ejecutar Consulta - Guarde y Re-Intente")
         Exit Sub
      End If
      On Error GoTo 0
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1.AddItem !valor
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
      '
   End If
   
End Sub

Private Sub Command3_Click()
10  COMO$ = InputBox$("Leyenda (Prompt)", "Configuración Consultas SQL", CMBCAMPOS.TEXT)
    If COMO$ = "" Then GoTo 10
    Combo1.Clear
    Combo1.AddItem ("¿" + COMO$ + "?")
    Combo1.TEXT = "¿" + COMO$ + "?"
End Sub

Private Sub Command4_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& > 1 Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i& - 2
         List1.Selected(i& - 2) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
End Sub

Private Sub Command5_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& < List1.ListCount Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i&
         List1.Selected(i&) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
   Call ARMA_CONSULTX
End Sub

Private Sub Command6_Click()
    '
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
       TOTALI$ = "0": If Check1.Value = 1 Then TOTALI$ = "1"
       AGRUPA$ = "0": If Check2.Value = 1 Then AGRUPA$ = "1"
       OCULTA$ = "0": If Check4.Value = 1 Then OCULTA$ = "1"
       MODOSOR$ = "    "
       If Option2.Value = True Then MODOSOR$ = "ASC "
       If Option3.Value = True Then MODOSOR$ = "DESC"
       CAMPOSE$ = AJUSTI$(List1.List(i&), 168)
       While Len(CAMPOSE$) < 168: CAMPOSE$ = CAMPOSE$ + " ": Wend
       Mid$(CAMPOSE$, 65, 8) = AJUSTI$(Text1, 8)
       Mid$(CAMPOSE$, 73, 1) = OCULTA$
       Mid$(CAMPOSE$, 75, 1) = AGRUPA$
       Mid$(CAMPOSE$, 76, 1) = TOTALI$
       Mid$(CAMPOSE$, 77, 4) = MODOSOR$
       Mid$(CAMPOSE$, 81, 64) = AJUSTI$(Text2, 64)
       Mid$(CAMPOSE$, 145, 16) = AJUSTI$(Text5, 16)
       Mid$(CAMPOSE$, 161, 8) = AJUSTI$(Text6, 8)
       '
       If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
         TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
         CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
         CONDIVALO$ = AJUSTI$(Text7, 512)
         'If TRIM$(CONDIVALO$) <> "" Then
            While Len(CAMPOSE$) < (168 + 64 + 64 + 512): CAMPOSE$ = CAMPOSE$ + " ": Wend
            Mid$(CAMPOSE$, 169, 64) = TAVALO$
            Mid$(CAMPOSE$, 233, 64) = CAVALO$
            Mid$(CAMPOSE$, 297, 512) = CONDIVALO$
         'End If
       End If
       List1.RemoveItem (i&)
       List1.AddItem CAMPOSE$, i&
       List1.Selected(i&) = True
       '
       For i& = 1 To LSTMOSTARCMP.ListCount
          If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             LSTMOSTARCMP.RemoveItem (i& - 1)
             LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
             LSTMOSTARCMP.Selected(i& - 1) = True
          End If
       Next i&
       '
       For i& = 1 To CAMAUXI.ListCount
          If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             CAMAUXI.RemoveItem (i& - 1)
             CAMAUXI.AddItem CAMPOSE$, i& - 1
             CAMAUXI.Selected(i& - 1) = True
          End If
       Next i&
       '
    End If
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command8_Click()
   'ADAT$ = Combo2.TEXT
   ADAT$ = Label1.Caption
   Call CARGADEBASE(ADAT$, TEBUSCA$, OKXX%)
   
End Sub

Private Sub EJECUTAR_Click()
    SHOWREP07.Label2 = ""
    If TRIM$(Label9) <> "" Then
          SHOWREP07.Label3 = Label9        ' Descripcion de la consulta
       Else
          SHOWREP07.Label3 = "Consulta en Desarrollo"
    End If
    TEXMO$ = ARMA_CONSULTB
    Call DECOD_CONSULTA(TEXMO$)
    'If TRIM$(Combo2.TEXT) <> "" Then
    If TRIM$(Label1.Caption) <> "" Then
         SHOWREP07.Label2 = ""
         'SHOWREP07.Label2 = Combo2.TEXT
         SHOWREP07.Label2 = Label1.Caption
      Else
         SHOWREP07.Label2 = "_TEXT_"
    End If
    SHOWREP07.Show 1
    MODICONSU% = 0
End Sub

Private Sub GUARDACON_Click()
    '
    FGUAREP07.Text1 = TRIM$(Label1.Caption)
    CODMO$ = TRIM$(Label1.Caption)
    FGUAREP07.Text2 = Label9
    FGUAREP07.Text3 = ""
    FGUAREP07.Label11 = ""
    FGUAREP07.Text4 = ""
    FGUAREP07.Label18 = ""
    FGUAREP07.Show 1
    If FGUAREP07.Label1 = "OK" Then
10     CODMO$ = TRIM$(FGUAREP07.Text1)
       DESMO$ = TRIM$(FGUAREP07.Text2)
       TEXMO$ = ARMA_CONSULTB
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If
       
       Call ABRETACON
       strSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(strSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_QUERY = CODMO$
          !DES_QUERY = DESMO$
          !TIP_QUERY = 0        ' REPORTE TABULADO
          !PAR_QUERY = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       '
30     If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
           ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + CODMO$ + ".SQF"
           Call SAVEFILE(ASALI$, TEXMO$)
       End If
       '
       If FGUAREP07.Text3 <> "" Then
          If FGUAREP07.Label11 <> "" Then
             ASALI$ = FGUAREP07.Label11 + "\" + FGUAREP07.Text3 + ".SQF"
             Call SAVEFILE(ASALI$, TEXMO$)
          End If
       End If
       '
    End If
    '
    Label1.Caption = CODMO$
    '
End Sub




Private Sub Label8_DblClick()
    Unload Me
End Sub

Private Sub List1_Click()
   If MODIFENCA% = 2 Then Exit Sub
   If MODIFENCA% = 1 Then
      If COMALTER%("¿ Descarta Cambios ?\\No, Conservar\Si, Descartar") <> 2 Then
         MODIFENCA% = 2
         On Error Resume Next
         List1.ListIndex = ILIS1&
         On Error GoTo 0
         MODIFENCA% = 1
         Exit Sub
      End If
   End If
   '
   CAMPOSE$ = List1.TEXT
   Text1 = TRIM$(Mid$(CAMPOSE$, 65, 8))
   Check3.Value = 0: If Mid$(CAMPOSE$, 74, 1) = "1" Then Check3.Value = 1
   Check4.Value = 0: If Mid$(CAMPOSE$, 73, 1) = "1" Then Check4.Value = 1    ' OCULTA
   Check1.Value = 0: If Mid$(CAMPOSE$, 76, 1) = "1" Then Check1.Value = 1
   Check2.Value = 0: If Mid$(CAMPOSE$, 75, 1) = "1" Then Check2.Value = 1
   Text2 = TRIM$(Mid$(CAMPOSE$, 81, 64))
   Text5 = TRIM$(Mid$(CAMPOSE$, 145, 16))
   Text6 = TRIM$(Mid$(CAMPOSE$, 161, 8))
   '
   TAUXI$ = Mid$(CAMPOSE$, 169, 640)
   Combo2.TEXT = TRIM$(Mid$(TAUXI$, 1, 64))
   Combo3.TEXT = TRIM$(Mid$(TAUXI$, 65, 64))
   Text7.TEXT = TRIM$(Mid$(TAUXI$, 129, 512))
   '
   ORDEX$ = TRIM$(UCase$(Mid$(CAMPOSE$, 77, 4)))
   Option1.Value = True
   If ORDEX$ = "ASC" Then Option2.Value = True
   If ORDEX$ = "DESC" Then Option3.Value = True
   MODIFENCA% = 0
   ILIS1& = List1.ListIndex
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   i& = List2.ListIndex
   If i& >= 0 Then
      CMBCAMPOS.TEXT = TRIM$(Left$(TTXX$, 24))
      CMBOPERADOR.TEXT = TRIM$(Mid$(TTXX$, 26, 8))
        VALORX$ = TRIM$(REPLACAR$(Mid$(TTXX$, 35), "'", ""))
      Combo1.TEXT = VALORX$
      List2.RemoveItem i&
   End If
   Call ARMA_CONSULTX
End Sub

Private Sub LSTMOSTARCMP_Click()
  Call ARMASELFIELD
End Sub

Sub ARMASELFIELD()
    '
    For i& = 0 To LSTMOSTARCMP.ListCount - 1
      ESTADO = LSTMOSTARCMP.Selected(i&)
      CAMPO$ = TRIM$(LSTMOSTARCMP.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 20
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
10         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 10
              End If
           Next J&
      End If
20  Next i&
    '
    For i& = 0 To CAMAUXI.ListCount - 1
      ESTADO = CAMAUXI.Selected(i&)
      CAMPO$ = TRIM$(CAMAUXI.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 120
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
110         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 110
              End If
           Next J&
      End If
120 Next i&
    '
    Call ARMA_CONSULTX
    '
End Sub
'
Private Sub LSTTABLAS_Click()
    '
    Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Sub ARMALISTAS()

    TXTCONSULTA.TEXT = ""
    CMBCAMPOS.Clear
    CMBOPERADOR.TEXT = ""
    Combo1.Clear
    Combo4.Clear
    Combo5.Clear
    LSTMOSTARCMP.Clear
    CAMAUXI.Clear
    List1.Clear
    '
    TAPRIN$ = LSTTABLAS.TEXT
    PREFPRIN$ = ""
    TASEC$ = LSTTABLASEC.TEXT
    If TASEC$ <> "" Then PREFPRIN$ = TAPRIN$ + "."
    If TAPRIN$ = "" Then Exit Sub
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           LSTMOSTARCMP.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           Combo4.AddItem AUX.Fields(i& - 1).Name
        End If
    Next i&
    '
    If TASEC$ <> "" Then
       PREFSEC$ = TASEC$ + "."
       Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TASEC$))
       For i& = 1 To AUX.Fields.Count
          If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
            CMBCAMPOS.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            LSTMOSTARCMP.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            Combo5.AddItem AUX.Fields(i& - 1).Name
          End If
       Next i&
    End If
    '
    For i& = 1 To 9
        CAMAUXI.AddItem "ValoBada_" & i&
    Next i&
    '
End Sub

Private Sub CMBBASE_CLICK()
    LSTTABLAS.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Private Sub Form_Load()
  Call ABRECONEXION
  KLIK% = 0
  CMBOPERADOR.Clear
  
  CMBOPERADOR.AddItem " =  "
  CMBOPERADOR.AddItem " <  "
  CMBOPERADOR.AddItem " <= "
  CMBOPERADOR.AddItem " >  "
  CMBOPERADOR.AddItem " >= "
  CMBOPERADOR.AddItem " <> "
  CMBOPERADOR.AddItem "LIKE"
  '
  CMBJUNTAR.AddItem "INNER"
  CMBJUNTAR.AddItem "LEFT"
  CMBJUNTAR.AddItem "RIGHT"
  '
  UNIONALL$ = ""
  '
  Combo6.AddItem "DEUDOR"
  Combo6.AddItem "PROVED"
  Combo6.AddItem "MASTER"
  Combo6.AddItem "USERS"
  Combo6.AddItem "SUTOT"
 
  Call MUESTRATABLAS

End Sub

Sub MUESTRATABLAS()
    LSTTABLAS.Clear: LSTTABLASEC.Clear
    Combo2.Clear: Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  'SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    SQLX$ = "SELECT name From sys.objects WHERE type IN ('P', 'U', 'V') AND LEFT(name, 3) <> 'dt_' AND LEFT(name, 3) <> 'pk_' AND LEFT(name, 3) <> 'sys';"
    
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       If TipoObjectSql$(SAFETEXT$(AUX!Name)) <> "TABLA" Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
       LSTTABLASEC.AddItem UCase$(AUX!Name)
       Combo2.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
    'VER AHORA CON TipoObjectSql$ SI ES SOTORED COMO SE PUEDE TRATAR PARA QUE FUNCIONE EN GENREP
End Sub

'
Sub ARMA_CONSULTX()
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = TRIM$(Left$(List1.List(i&), 64))
       '
       If UCase$(Left$(CAMPO$, 8)) = "VALOBADA" Then GoTo 19
       '
       ALIASCAMPO$ = CAMPO$
       If Mid$(List1.List(i&), 75, 1) = "1" Then
          CAMPO1$ = CAMPO$
          If InStr(CAMPO1$, ".") > 0 Then
            PPXYZ% = InStr(CAMPO1$, ".")
            CAMPO1$ = Mid$(CAMPO1$, PPXYZ% + 1)
          End If
          ALIASCAMPO$ = "Sx" + CAMPO1$
          CAMPO$ = "SUM(" + CAMPO$ + ") AS " + ALIASCAMPO$
          AGRUPA% = 1
       End If
       If VUELTA& = 0 Then
           SQLX$ = SQLX$ + CAMPO$
          Else
           SQLX$ = SQLX$ + ", " + CAMPO$
       End If
       VUELTA& = VUELTA& + 1
       '
       ORDEX$ = UCase$(TRIM$(Mid$(List1.List(i&), 77, 4)))
       If ORDEX$ <> "" Then
          If VUELTB& > o Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = Mid$(List1.List(i&), 75, 1)
       If AGRUX$ <> "1" Then
          If VUELTC& > o Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If

19  Next i&
    '
    If VUELTA = 0 Then SQLX$ = "SELECT *"
    '
    TAPRIN$ = LSTTABLAS.TEXT
    SQLY$ = " FROM " + TAPRIN$
    '
    TASEC$ = LSTTABLASEC.TEXT
    
    If TASEC$ <> "" Then
       CAMPRIN$ = TRIM$(Combo4.TEXT)
       If CAMPRIN$ <> "" Then
          CAMSEC$ = TRIM$(Combo5.TEXT)
          If CAMSEC$ <> "" Then
             TIPOJOIN$ = TRIM$(CMBJUNTAR.TEXT)
             If TIPOJOIN$ = "" Then TIPOJOIN$ = "INNER"
             SQLY$ = SQLY$ + " " & TIPOJOIN$ & " JOIN"
             SQLY$ = SQLY$ + " " + TASEC$
             
             SQLY$ = SQLY$ + " ON " + TAPRIN$ + "." + CAMPRIN$
             SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
          End If
       End If
    End If
    '
    If List2.ListCount > 0 Then
       JJ& = 0
       For i& = 0 To List2.ListCount - 1
          SQLY$ = SQLY$ '+ vbCrLf
          'INTERVENCION OR
          If JJ& < 1 Then
              SQLY$ = SQLY$ + " WHERE "
          ElseIf UCase$((Left$(TRIM$(List2.List(i&)), 3))) = "OR " Then
              SQLY$ = SQLY$ + " "
          Else
               SQLY$ = SQLY$ + " AND "
          End If
          SQLY$ = SQLY$ + TRIM$(Left$(List2.List(i&), 24)) + " " + TRIM$(Mid$(List2.List(i&), 26, 8)) + " " + TRIM$(Mid$(List2.List(i&), 35))
          JJ& = JJ& + 1
       Next i&
    End If
    '
    TXTCONSULTA.TEXT = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " ORDER BY " + SQLZ$
    '
    '
    If chkUnionAll.Value = 1 Then
        If LSTTABLASEC.ListIndex > 0 Then
            PrimeraTabla$ = TRIM$(LSTTABLAS.List(LSTTABLAS.ListIndex))
            SegundaTabla$ = TRIM$(LSTTABLASEC.List(LSTTABLASEC.ListIndex))
            ConsultaSql$ = TXTCONSULTA.TEXT
            POSI% = InStr(1, ConsultaSql$, PrimeraTabla$)
            If POSI% > 0 Then
                ConsultaSql$ = Mid$(ConsultaSql$, 1, POSI% - 1) & SegundaTabla$ & Mid$(ConsultaSql$, POSI% + Len(PrimeraTabla$))
                TXTCONSULTA.TEXT = SQLX$ + SQLY$ & " UNION ALL " & ConsultaSql$
            End If
            
            UNIONALL$ = SegundaTabla$
        End If
    End If
       
    TXTCANREGI.TEXT = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
End Sub

Function ARMA_CONSULTB()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(FGUAREP07.Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(FGUAREP07.Text2) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   0" + Chr$(124)
    CONSUB$ = CONSUB$ + "SELTABLE = " + LSTTABLAS.TEXT + Chr$(124)
    '
      TASECU$ = TRIM$(LSTTABLASEC.TEXT)
    If TASECU$ <> "" And chkUnionAll.Value = 0 Then CONSUB$ = CONSUB$ + "SELTASEC = " + TASECU$ + Chr$(124)
      CAMPOMATCH1$ = TRIM$(Combo4.TEXT)
    If CAMPOMATCH1$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI1 = " + CAMPOMATCH1$ + Chr$(124)
      CAMPOMATCH2$ = TRIM$(Combo5.TEXT)
    If CAMPOMATCH2$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI2 = " + CAMPOMATCH2$ + Chr$(124)
    '
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = List1.List(i&)
       CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    Next i&
    '
    For i& = 0 To List2.ListCount - 1
       CONDI$ = List2.List(i&)
       CONSUB$ = CONSUB$ + "CONDITIO = " + CONDI$ + Chr$(124)
    Next i&
    '
    TIPODEJOIN$ = TRIM$(CMBJUNTAR.TEXT)
    CONSUB$ = CONSUB$ + "TIPOJOIN = " + TIPODEJOIN$ + Chr$(124)
    '
    If chkUnionAll.Value = 1 And LSTTABLASEC.ListIndex > 0 Then
        SegundaTabla$ = LSTTABLASEC.List(LSTTABLASEC.ListIndex)
        CONSUB$ = CONSUB$ + "UNIONALL = " + SegundaTabla$ + Chr$(124)
    End If
    
    ARMA_CONSULTB = CONSUB$
End Function

Private Sub CMDPEGAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Clipboard.GetText + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

'QUEDA PENIDIENTE SI ES UN TEXTO  EL QUE VA A COMPARAR VAYA ENTRE _
COMILLAS SIMPLES (HACERLO CON ISNUMERIC ) _
SI EL OPERADOR ES LIKE DEBERIA LLEVAR EL SIMBOLO DEL PORCENTAJE _
FALTA TAMBIEN EL TEMA DEL EDITOR DE TEXTO _
VALIDAR LOS CAMPOS DE CONDICION _
CAMBIAR LAS PROPIEDADES DE LOS COMBO Y TEXT DE LA CONDICION PAARA QUE _
NO SEAN EDITABLES _
Y EJECUTAR LA CONSULTA

Private Sub LSTTABLASEC_Click()
    '
    If chkUnionAll.Value = 0 And UNIONALL$ = "" Then Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Private Sub Option1_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option2_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option3_Click()
   MODIFENCA% = 1
End Sub

Private Sub Text1_Change()
   MODIFENCA% = 1
End Sub

Private Sub Text2_Change()
   MODIFENCA% = 1
End Sub

