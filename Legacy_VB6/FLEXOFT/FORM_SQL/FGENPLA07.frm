VERSION 5.00
Begin VB.Form FGENPLA07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes - Tipo Plan Mensual FLEXOFT"
   ClientHeight    =   6570
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
   ScaleHeight     =   6570
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   69
      Top             =   5760
      Width           =   210
   End
   Begin VB.CommandButton Command15 
      Caption         =   "("
      Height          =   600
      Left            =   0
      TabIndex        =   68
      Top             =   5040
      Width           =   200
   End
   Begin VB.CommandButton Command14 
      Caption         =   ")"
      Height          =   600
      Left            =   4530
      TabIndex        =   67
      Top             =   5040
      Width           =   200
   End
   Begin VB.Frame Frame4 
      Caption         =   "Valores"
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
      Left            =   4320
      TabIndex        =   64
      Top             =   3120
      Width           =   2415
      Begin VB.OptionButton Option5 
         Caption         =   "Saldo Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   66
         Top             =   240
         Width           =   940
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Resultado Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   30
         TabIndex        =   65
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Columnas Complementarias"
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
      Left            =   6840
      TabIndex        =   59
      Top             =   3120
      Width           =   4935
      Begin VB.CheckBox Check4 
         Caption         =   "Total"
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
         Left            =   3940
         TabIndex        =   63
         Top             =   360
         Width           =   900
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Promedio"
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
         Left            =   2600
         TabIndex        =   62
         Top             =   360
         Width           =   1200
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Máximo"
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
         Left            =   1380
         TabIndex        =   61
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mínimo"
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
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
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
      Left            =   2310
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   3630
      Width           =   1815
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
      Height          =   2445
      IntegralHeight  =   0   'False
      Left            =   2205
      Sorted          =   -1  'True
      TabIndex        =   55
      Top             =   1095
      Width           =   1905
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
      Left            =   210
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   3630
      Width           =   1815
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
      Height          =   2445
      IntegralHeight  =   0   'False
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1100
      Width           =   1905
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
      TabIndex        =   46
      Top             =   4095
      Width           =   870
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
      TabIndex        =   45
      Top             =   4350
      Width           =   870
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
      TabIndex        =   44
      Top             =   4605
      Width           =   870
   End
   Begin VB.Frame Frame5 
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   4305
      TabIndex        =   27
      Top             =   840
      Width           =   7470
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
         Height          =   880
         Left            =   6270
         TabIndex        =   50
         Top             =   110
         Width           =   1065
         Begin VB.OptionButton Option3 
            Caption         =   "Des"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   51
            Top             =   600
            Width           =   645
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Asc"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   52
            Top             =   400
            Width           =   645
         End
         Begin VB.OptionButton Option1 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   53
            Top             =   200
            UseMaskColor    =   -1  'True
            Value           =   -1  'True
            Width           =   645
         End
      End
      Begin VB.ComboBox ComboValor 
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
         Index           =   1
         Left            =   420
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1700
         Width           =   2235
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
         Left            =   6720
         TabIndex        =   39
         Top             =   1155
         Width           =   630
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
         Left            =   3990
         TabIndex        =   38
         Top             =   1155
         Width           =   2115
      End
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
         Left            =   3990
         TabIndex        =   37
         Top             =   1770
         Width           =   2115
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
         Left            =   6720
         TabIndex        =   36
         Top             =   1785
         Width           =   630
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
         Height          =   280
         Left            =   2940
         TabIndex        =   34
         Top             =   1700
         Width           =   630
      End
      Begin VB.ComboBox ComboFecha 
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
         Left            =   420
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   2235
      End
      Begin VB.ComboBox ComboClase 
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
         Left            =   3990
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   420
         Width           =   2130
      End
      Begin VB.ComboBox ComboValor 
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
         Index           =   0
         Left            =   420
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1300
         Width           =   2235
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "+"
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
         Index           =   0
         Left            =   105
         TabIndex        =   48
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valores:"
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
         Index           =   2
         Left            =   210
         TabIndex        =   31
         Top             =   1050
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "+"
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
         Index           =   1
         Left            =   105
         TabIndex        =   49
         Top             =   1635
         Width           =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   3780
         X2              =   3780
         Y1              =   105
         Y2              =   2205
      End
      Begin VB.Label Label6 
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
         Left            =   6195
         TabIndex        =   43
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Encabezado:"
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
         Left            =   3885
         TabIndex        =   42
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Tabla Col. Eco:"
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
         Left            =   3885
         TabIndex        =   41
         Top             =   1545
         Width           =   1575
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
         Left            =   6195
         TabIndex        =   40
         Top             =   1220
         Width           =   750
      End
      Begin VB.Label Label14 
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
         Height          =   285
         Left            =   2415
         TabIndex        =   35
         Top             =   1420
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Columna (Fecha):"
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
         Index           =   3
         Left            =   210
         TabIndex        =   33
         Top             =   345
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Clase:"
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
         Index           =   1
         Left            =   3885
         TabIndex        =   32
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   240
      Left            =   105
      TabIndex        =   24
      Top             =   1575
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
      TabIndex        =   23
      Top             =   5985
      Width           =   7095
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9915
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   21
      Top             =   4200
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   22
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
      ItemData        =   "FGENPLA07.frx":0000
      Left            =   200
      List            =   "FGENPLA07.frx":0002
      TabIndex        =   18
      Top             =   4980
      Width           =   4320
   End
   Begin VB.CommandButton Command5 
      Height          =   540
      Left            =   6090
      Picture         =   "FGENPLA07.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7455
      Width           =   330
   End
   Begin VB.CommandButton Command4 
      Height          =   540
      Left            =   6510
      Picture         =   "FGENPLA07.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6840
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
      ItemData        =   "FGENPLA07.frx":0618
      Left            =   3780
      List            =   "FGENPLA07.frx":061A
      TabIndex        =   13
      Top             =   6840
      Width           =   1905
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
      Left            =   9870
      TabIndex        =   5
      Top             =   4485
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
      TabIndex        =   4
      Top             =   4095
      Width           =   855
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
      Left            =   7980
      Picture         =   "FGENPLA07.frx":061C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agregar Condición"
      Top             =   4080
      Width           =   750
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
      Top             =   4980
      Width           =   7095
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
      Height          =   2120
      IntegralHeight  =   0   'False
      ItemData        =   "FGENPLA07.frx":0926
      Left            =   735
      List            =   "FGENPLA07.frx":0928
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   6840
      Width           =   1905
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
      TabIndex        =   7
      Top             =   3990
      Width           =   8490
      Begin VB.CommandButton Command3 
         Caption         =   "?"
         Height          =   300
         Left            =   7245
         TabIndex        =   15
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
         Left            =   735
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "CMBCAMPOS"
         Top             =   315
         Width           =   2340
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
         ItemData        =   "FGENPLA07.frx":092A
         Left            =   3255
         List            =   "FGENPLA07.frx":092C
         TabIndex        =   11
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
         Left            =   6855
         TabIndex        =   10
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
         Left            =   4515
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   315
         Width           =   2235
      End
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
      Left            =   1995
      TabIndex        =   58
      Top             =   3585
      Width           =   330
   End
   Begin VB.Label Label15 
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
      Left            =   2100
      TabIndex        =   56
      Top             =   840
      Width           =   1695
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
      TabIndex        =   26
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
      Left            =   945
      TabIndex        =   25
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
      TabIndex        =   20
      Top             =   105
      Width           =   8310
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
      TabIndex        =   19
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
      Left            =   2160
      TabIndex        =   14
      Top             =   6720
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
      Left            =   105
      TabIndex        =   8
      Top             =   6720
      Width           =   1275
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
      Left            =   105
      TabIndex        =   6
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "FGENPLA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KLIK%
Dim MODIFENCA%, ILIS1&

'
Private Sub CMDBORRAR_Click()
TXTCONSULTA.TEXT = ""
End Sub


Private Sub Combo2_Change()
    Call ARMA_CONSULTX
End Sub

Private Sub Combo3_Change()
    Call ARMA_CONSULTX
End Sub

Private Sub ComboClase_Change()
    Call ARMA_CONSULTX
End Sub

Private Sub ComboClase_Click()
    Call ARMA_CONSULTX
End Sub

Private Sub ComboFecha_Change()
    Call ARMA_CONSULTX
End Sub

Private Sub ComboFecha_Click()
    Call ARMA_CONSULTX
End Sub

Private Sub ComboValor_Change(Index As Integer)
    Call ARMA_CONSULTX
End Sub

Private Sub ComboValor_Click(Index As Integer)
    Call ARMA_CONSULTX
End Sub

Private Sub label1_Change()
    Dim FISELE$(128)
    List2.Clear
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Option4.Value = True
    
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
    
    STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
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
              CMBBASE = UCase$(RTrim(Mid$(COMAI$, 11)))
              GoTo 19
            ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
              TASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLAS.ListCount
                 If UCase$(LSTTABLAS.LIST(KKI% - 1)) = TASELE Then
                   LSTTABLAS.Selected(KKI% - 1) = True
                   Call LSTTABLAS_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "SELTASEC =" Then
              TASELSEC$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLASEC.ListCount
                 If UCase$(LSTTABLASEC.LIST(KKI% - 1)) = TASELSEC Then
                   LSTTABLASEC.Selected(KKI% - 1) = True
                   Call LSTTABLASEC_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "MATCHFI1 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo2.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "MATCHFI2 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo3.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "CAMFECHA =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              ComboFecha.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "CAMCLASE =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              ComboClase.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
              Text1 = TRIM$(Mid$(CASELE$, 65, 10))
                ORDEX$ = TRIM$(UCase$(Mid$(CASELE$, 77, 4)))
              Text2 = TRIM$(Mid$(CASELE$, 81, 64))
              Text5 = TRIM$(Mid$(CASELE$, 145, 16))
              Option1.Value = True
                If ORDEX$ = "ASC" Then Option2.Value = True
                If ORDEX$ = "DESC" Then Option3.Value = True
              Text6 = TRIM$(Mid$(CASELE$, 161, 8))
            ElseIf Left$(COMAI$, 10) = "CAMVALOR =" Then
              CASELE$ = UCase$(TRIM(Mid$(COMAI$, 11)))
              LAB10$ = "+": If Mid$(CASELE$, 77, 1) = "-" Then LAB10 = "-"
              On Error Resume Next
              ComboValor(0).TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
              Label10(0) = LAB10$
              Text7 = TRIM$(Mid$(CASELE$, 65, 10))
            ElseIf Left$(COMAI$, 10) = "CAMVALO2 =" Then
              CASELE$ = UCase$(TRIM(Mid$(COMAI$, 11)))
              LAB10$ = "+": If Mid$(CASELE$, 77, 1) = "-" Then LAB10 = "-"
              On Error Resume Next
              ComboValor(1).TEXT = TRIM$(Left$(CASELE$, 64))
              Label10(1) = LAB10$
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
              List2.AddItem TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "COL_MIN = " Then
              If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
                Check1.Value = 1
              End If
            ElseIf Left$(COMAI$, 10) = "COL_MAX = " Then
              If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
                Check2.Value = 1
              End If
            ElseIf Left$(COMAI$, 10) = "COL_PRO = " Then
              If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
                Check3.Value = 1
              End If
            ElseIf Left$(COMAI$, 10) = "COL_TOT = " Then
              If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
                Check4.Value = 1
              End If
            ElseIf Left$(COMAI$, 10) = "FOR_VAL = " Then
              If TRIM$(Mid$(COMAI$, 11, 8)) = "SALFINAL" Then
                 Option5.Value = True
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
    'SI NO ESTABA EN SQL ENTRO EN MDB Y CIERRA EL RECORDSET
    If ESTAENSQL% = 0 Then
      MIRS.Close
      Set MIRS = Nothing
    End If
    '
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
   TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
   If (TPC = 52 Or TPC = 62) And VALOCO$ = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   '
   If TPC <> 52 And TPC <> 62 Then VALOCO$ = "'" + VALOCO$ + "'"
   SQLY$ = SQLY$ + " " + VALOCO$
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
      TCAMPO$ = CMBCAMPOS.TEXT
      TTABLA$ = LSTTABLAS.TEXT
      PPXX% = InStr(TCAMPO$, ".")
      If PPXX% > 1 Then
         TTABLA$ = Left$(TCAMPO$, PPXX% - 1)
         TCAMPO$ = Mid$(TCAMPO$, PPXX% + 1)
      End If
      '
      SQLX$ = "SELECT DISTINCT " & TCAMPO$ & " AS VALOR FROM " & TTABLA$ & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & TCAMPO$ & " ASC"
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
            Combo1.AddItem !VALOR
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
         CAMPOSE$ = List1.LIST(i& - 1)
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
         CAMPOSE$ = List1.LIST(i& - 1)
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
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
       TOTALI$ = "0": If Check1.Value = 1 Then TOTALI$ = "1"
       AGRUPA$ = "0": If Check2.Value = 1 Then AGRUPA$ = "1"
       MODOSOR$ = "    "
       If Option2.Value = True Then MODOSOR$ = "ASC "
       If Option3.Value = True Then MODOSOR$ = "DESC"
       CAMPOSE$ = AJUSTI$(List1.LIST(i&), 64) + AJUSTI$(Text1, 10) + AGRUPA$ + TOTALI$ + MODOSOR$ + AJUSTI$(Text2, 64) + AJUSTI$(Text5, 16) + AJUSTI$(Text6, 8)
       List1.RemoveItem (i&)
       List1.AddItem CAMPOSE$, i&
       List1.Selected(i&) = True
    End If
    For i& = 1 To LSTMOSTARCMP.ListCount
       If TRIM$(Left$(LSTMOSTARCMP.LIST(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          LSTMOSTARCMP.RemoveItem (i& - 1)
          LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
          LSTMOSTARCMP.Selected(i& - 1) = True
       End If
    Next i&
    Call ARMA_CONSULTX
End Sub

Private Sub Command8_Click()
   'ADAT$ = Combo2.TEXT
   ADAT$ = Label1.Caption
   Call CARGADEBASE(ADAT$, TEBUSCA$, OKXX%)
   
End Sub

Private Sub EJECUTAR_Click()
    SHOWPLA07.Label2 = ""
    If TRIM$(Label9) <> "" Then
          SHOWPLA07.Label3 = Label9        ' Descripcion de la consulta
       Else
          SHOWPLA07.Label3 = "Consulta en Desarrollo"
    End If
    TEXMO$ = ARMA_CONSULTB
    Call DECOD_CONSULTA(TEXMO$)
    'If TRIM$(Combo2.TEXT) <> "" Then
    If TRIM$(Label1.Caption) <> "" Then
         SHOWPLA07.Label2 = ""
         'SHOWPLA07.Label2 = Combo2.TEXT
         SHOWPLA07.Label2 = Label1.Caption
      Else
         SHOWPLA07.Label2 = "_TEXT_"
    End If
    If IMPOSEJEC% < 1 Then SHOWPLA07.Show 1
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
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 1, TEXMO$, OK%)
       End If
       '
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
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
          !TIP_QUERY = 1        ' Estadistica Mensual
          !PAR_QUERY = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       '
       If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
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

Private Sub Label10_DblClick(Index As Integer)
    If Label10(Index).Caption = "+" Then
         Label10(Index).Caption = "-"
      Else
         Label10(Index).Caption = "+"
    End If
End Sub

Private Sub Label11_DblClick(Index As Integer)
   If Index = 2 Then
      ComboValor(0).Clear
      ComboValor(1).Clear
      Dim AUX As ADODB.Recordset
      Set AUX = New ADODB.Recordset
      Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & LSTTABLAS.TEXT))
      For i& = 1 To AUX.Fields.Count
         ComboValor(0).AddItem AUX.Fields(i& - 1).Name
         ComboValor(1).AddItem AUX.Fields(i& - 1).Name
      Next i&
      Call ARMA_CONSULTX
   End If
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
   Text1 = TRIM$(Mid$(CAMPOSE$, 65, 10))
   Check1.Value = 0: If Mid$(CAMPOSE$, 76, 1) = "1" Then Check1.Value = 1
   Check2.Value = 0: If Mid$(CAMPOSE$, 75, 1) = "1" Then Check2.Value = 1
   Text2 = TRIM$(Mid$(CAMPOSE$, 81, 64))
   Text5 = TRIM$(Mid$(CAMPOSE$, 145, 16))
   Text6 = TRIM$(Mid$(CAMPOSE$, 161, 8))
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
      CAMPO$ = TRIM$(LSTMOSTARCMP.LIST(i&))
      If ESTADO = True Then
           For j& = 1 To List1.ListCount
              If TRIM$(Left$(List1.LIST(j& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 20
           Next j&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
10         For j& = 1 To List1.ListCount
              If TRIM$(Left$(List1.LIST(j& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (j& - 1)
                 If j& > 1 Then List1.Selected(j& - 2) = True
                 GoTo 10
              End If
           Next j&
      End If
20  Next i&
    Call ARMA_CONSULTX
    '
End Sub
'
Private Sub LSTTABLAS_Click()
    Call ARMALISTAS
    Call ARMA_CONSULTX
End Sub

Private Sub LSTTABLASEC_Click()
    Call ARMALISTAS
    Call ARMA_CONSULTX
End Sub

Sub ARMALISTAS()
    TXTCONSULTA.TEXT = ""
    CMBCAMPOS.Clear
    CMBOPERADOR.TEXT = ""
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    LSTMOSTARCMP.Clear
    List1.Clear
    '
    TAPRIN$ = LSTTABLAS.TEXT
    PREFPRIN$ = ""
    TASEC$ = LSTTABLASEC.TEXT
    If TASEC$ <> "" Then PREFPRIN$ = TAPRIN$ + "."
    If TAPRIN$ = "" Then Exit Sub
    '
    ComboFecha.Clear
    ComboClase.Clear
    ComboValor(0).Clear
    ComboValor(1).Clear
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        ComboFecha.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        ComboClase.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        ComboValor(0).AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        ComboValor(1).AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        Combo2.AddItem AUX.Fields(i& - 1).Name
    Next i&
    '
    If TASEC$ <> "" Then
       PREFSEC$ = TASEC$ + "."
       Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TASEC$))
       For i& = 1 To AUX.Fields.Count
           CMBCAMPOS.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
           ComboFecha.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
           ComboClase.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
           ComboValor(0).AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
           ComboValor(1).AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
           Combo3.AddItem AUX.Fields(i& - 1).Name
       Next i&
    End If
    '
End Sub

Private Sub CMBBASE_CLICK()
    LSTTABLAS.Clear: LSTTABLASEC.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
       LSTTABLASEC.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Sub MUESTRATABLAS()
    LSTTABLAS.Clear: LSTTABLASEC.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
       LSTTABLASEC.AddItem UCase$(AUX!Name)
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
  Call MUESTRATABLAS
  '
End Sub
Private Sub Command15_Click()
    'INTERVENCION OR

    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 1) <> "(" Then
          List2.LIST(i%) = "(" & List2.TEXT
        End If
        Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command12_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 3) <> "OR " Then
          List2.LIST(i%) = "OR " & List2.TEXT
         End If
         Call ARMA_CONSULTX
    End If
End Sub
Private Sub Command14_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Right(TRIM$(List2.TEXT), 1) <> ")" Then
          List2.LIST(i%) = List2.TEXT & ")"
        End If
        Call ARMA_CONSULTX
    End If
End Sub

Sub ARMA_CONSULTX()
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For i& = 1 To 4
       If i& = 1 Then CAMPO$ = ComboFecha.TEXT
       If i& = 2 Then CAMPO$ = ComboClase.TEXT
       If i& = 3 Then CAMPO$ = ComboValor(0).TEXT
       If i& = 4 Then CAMPO$ = ComboValor(1).TEXT
       ALIASCAMPO$ = CAMPO$
       If TRIM$(CAMPO$) <> "" Then
          If VUELTA& = 0 Then
              SQLX$ = SQLX$ + CAMPO$
             Else
              SQLX$ = SQLX$ + ", " + CAMPO$
          End If
          VUELTA& = VUELTA& + 1
       End If
       '
       ORDEX$ = "    "
       If i& = 1 Then ORDEX$ = "ASC "
       If i& = 2 Then
          If Option2.Value = True Then ORDEX$ = "ASC "
          If Option3.Value = True Then ORDEX$ = "DESC"
       End If
       If TRIM$(ORDEX$) <> "" Then
          If VUELTB& > O Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
    Next i&
    '
    If VUELTA = 0 Then SQLX$ = "SELECT *"
    '
    TAPRIN$ = LSTTABLAS.TEXT
    SQLY$ = " FROM " + TAPRIN$
    '
    TASEC$ = LSTTABLASEC.TEXT
    If TASEC$ <> "" Then
       CAMPRIN$ = TRIM$(Combo2.TEXT)
       If CAMPRIN$ <> "" Then
          CAMSEC$ = TRIM$(Combo3.TEXT)
          If CAMSEC$ <> "" Then
             SQLY$ = SQLY$ + " INNER JOIN " + TASEC$
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
          'If JJ& < 1 Then SQLY$ = SQLY$ + " WHERE " Else SQLY$ = SQLY$ + " AND "
          'INTERVENCION OR
          If JJ& < 1 Then
              SQLY$ = SQLY$ + " WHERE "
          ElseIf UCase$((Left$(TRIM$(List2.LIST(i&)), 3))) = "OR " Then
              SQLY$ = SQLY$ + " "
          Else
               SQLY$ = SQLY$ + " AND "
          End If

          SQLY$ = SQLY$ + TRIM$(Left$(List2.LIST(i&), 24)) + " " + TRIM$(Mid$(List2.LIST(i&), 26, 8)) + " " + TRIM$(Mid$(List2.LIST(i&), 35))
          JJ& = JJ& + 1
       Next i&
    End If
    '
    TXTCONSULTA.TEXT = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " ORDER BY " + SQLZ$
    TXTCANREGI.TEXT = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
End Sub

Function ARMA_CONSULTB()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(FGUAREP07.Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(FGUAREP07.Text2) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   1" + Chr$(124)
    '
    CONSUB$ = CONSUB$ + "SELTABLE = " + LSTTABLAS.TEXT + Chr$(124)
    '
    TASECU$ = TRIM$(LSTTABLASEC.TEXT)
    If TASECU$ <> "" Then
       CONSUB$ = CONSUB$ + "SELTASEC = " + TASECU$ + Chr$(124)
    End If
    '
    CAMPOMATCH1$ = TRIM$(Combo2.TEXT)
    If CAMPOMATCH1$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI1 = " + CAMPOMATCH1$ + Chr$(124)
    '
    CAMPOMATCH2$ = TRIM$(Combo3.TEXT)
    If CAMPOMATCH2$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI2 = " + CAMPOMATCH2$ + Chr$(124)
    '
    CAMPOSE$ = AJUSTI$(ComboFecha.TEXT, 64) + AJUSTI$("D8", 10) + "0" + "0" + "ASC " + AJUSTI$("", 64) + AJUSTI$("", 16) + AJUSTI$("", 8)
    CONSUB$ = CONSUB$ + "CAMFECHA = " + CAMPOSE$ + Chr$(124)
    '
    ORDEX$ = "    "
    If Option2.Value = True Then ORDEX$ = "ASC "
    If Option3.Value = True Then ORDEX$ = "DESC"
    CAMPOSE$ = AJUSTI$(ComboClase.TEXT, 64) + AJUSTI$(Text1, 10) + "0" + "0" + ORDEX$ + AJUSTI$(Text2, 64) + AJUSTI$(Text5, 16) + AJUSTI$(Text6, 8)
    CONSUB$ = CONSUB$ + "CAMCLASE = " + CAMPOSE$ + Chr$(124)
    '
    CAMPOSE$ = AJUSTI$(ComboValor(0), 64) + AJUSTI$(Text7, 10) + "0" + "1" + AJUSTI$(Label10(0), 4) + AJUSTI$("", 64) + AJUSTI$("", 16) + AJUSTI$("", 8)
    CONSUB$ = CONSUB$ + "CAMVALOR = " + CAMPOSE$ + Chr$(124)
    '
    CAMPOSE$ = AJUSTI$(ComboValor(1), 64) + AJUSTI$(Text7, 10) + "0" + "1" + AJUSTI$(Label10(1), 4) + AJUSTI$("", 64) + AJUSTI$("", 16) + AJUSTI$("", 8)
    CONSUB$ = CONSUB$ + "CAMVALO2 = " + CAMPOSE$ + Chr$(124)
    '
    For i& = 0 To List2.ListCount - 1
       CONDI$ = List2.LIST(i&)
       CONSUB$ = CONSUB$ + "CONDITIO = " + CONDI$ + Chr$(124)
    Next i&
    '
    If Check1.Value = 1 Then 'MINIMO
       CONSUB$ = CONSUB$ + "COL_MIN = SI" + Chr$(124)
    End If
    '
    If Check2.Value = 1 Then 'MAXIMO
       CONSUB$ = CONSUB$ + "COL_MAX = SI" + Chr$(124)
    End If
    '
    If Check3.Value = 1 Then 'PROMEDIO
       CONSUB$ = CONSUB$ + "COL_PRO = SI" + Chr$(124)
    End If
    '
    If Check4.Value = 1 Then 'TOTAL
       CONSUB$ = CONSUB$ + "COL_TOT = SI" + Chr$(124)
    End If
    '
    If Option4.Value = True Then
       CONSUB$ = CONSUB$ + "FOR_VAL = RESUMEN" + Chr$(124)
    ElseIf Option5.Value = True Then
       CONSUB$ = CONSUB$ + "FOR_VAL = SALFINAL" + Chr$(124)
    End If
    '
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

