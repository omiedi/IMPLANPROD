VERSION 5.00
Begin VB.Form REGMEDI07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de  medición"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   " "
      Top             =   2265
      Width           =   2880
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   10250
      TabIndex        =   30
      Top             =   4080
      Width           =   585
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   8400
      TabIndex        =   0
      Top             =   840
      Width           =   1560
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   645
      Index           =   0
      Left            =   3720
      TabIndex        =   7
      Top             =   6000
      Width           =   830
   End
   Begin VB.CommandButton Command6 
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
      Left            =   10650
      TabIndex        =   3
      Top             =   285
      Width           =   165
   End
   Begin VB.TextBox Text11 
      Height          =   300
      Left            =   9720
      TabIndex        =   2
      Text            =   " "
      Top             =   285
      Width           =   945
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00006000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   11040
      ScaleHeight     =   8940
      ScaleWidth      =   870
      TabIndex        =   47
      Top             =   0
      Width           =   930
      Begin VB.CommandButton Command7 
         Height          =   650
         Left            =   105
         Picture         =   "REGMEDI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Tabla de Instrumentos"
         Top             =   3000
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         Height          =   650
         Left            =   105
         Picture         =   "REGMEDI07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Imprime Registro de Medición"
         Top             =   2280
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Height          =   650
         Left            =   105
         Picture         =   "REGMEDI07.frx":07B4
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Formulario"
         Top             =   1560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   650
         Left            =   120
         Picture         =   "REGMEDI07.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Graba e Imprime Informe"
         Top             =   6960
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   48
         Top             =   7920
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
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
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
         Height          =   630
         Left            =   105
         Picture         =   "REGMEDI07.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   840
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   640
         Left            =   105
         Picture         =   "REGMEDI07.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "REGMEDI07.frx":1354
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8200
         Width           =   1515
      End
   End
   Begin VB.TextBox text9 
      Height          =   1185
      Left            =   2595
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2730
      Width           =   8280
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9960
      TabIndex        =   1
      Top             =   840
      Width           =   165
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registro de Medición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   31
      Top             =   4320
      Width           =   10935
      Begin VB.TextBox text19 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   3
         Left            =   10080
         TabIndex        =   29
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   1
         Left            =   3720
         TabIndex        =   13
         Top             =   2330
         Width           =   830
      End
      Begin VB.TextBox text19 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   2
         Left            =   10080
         TabIndex        =   23
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox text19 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   1
         Left            =   10080
         TabIndex        =   17
         Top             =   2325
         Width           =   795
      End
      Begin VB.TextBox text19 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   0
         Left            =   10080
         TabIndex        =   11
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   3
         Left            =   7700
         TabIndex        =   28
         Text            =   " "
         Top             =   3585
         Width           =   795
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   2
         Left            =   7700
         TabIndex        =   22
         Text            =   " "
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   1
         Left            =   7700
         TabIndex        =   16
         Text            =   " "
         Top             =   2325
         Width           =   795
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   0
         Left            =   7680
         TabIndex        =   10
         Text            =   " "
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox text17 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   3
         Left            =   6920
         TabIndex        =   27
         Top             =   3600
         Width           =   795
      End
      Begin VB.TextBox text17 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   2
         Left            =   6920
         TabIndex        =   21
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox text17 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   1
         Left            =   6920
         TabIndex        =   15
         Top             =   2330
         Width           =   795
      End
      Begin VB.TextBox text17 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   0
         Left            =   6920
         TabIndex        =   9
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   3
         Left            =   4540
         TabIndex        =   26
         Text            =   " "
         Top             =   3585
         Width           =   795
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   2
         Left            =   4540
         TabIndex        =   20
         Text            =   " "
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   1
         Left            =   4560
         TabIndex        =   14
         Text            =   " "
         Top             =   2325
         Width           =   795
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   0
         Left            =   4560
         TabIndex        =   8
         Text            =   " "
         Top             =   1680
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   3
         Left            =   3720
         TabIndex        =   25
         Top             =   3580
         Width           =   830
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   645
         Index           =   2
         Left            =   3720
         TabIndex        =   19
         Top             =   3000
         Width           =   830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   3
         Left            =   1350
         TabIndex        =   24
         Text            =   " "
         Top             =   3580
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   2
         Left            =   1350
         TabIndex        =   18
         Text            =   " "
         Top             =   3000
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Text            =   " "
         Top             =   2330
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   645
         Index           =   0
         Left            =   1350
         TabIndex        =   6
         Text            =   " "
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   645
         Index           =   3
         Left            =   9285
         TabIndex        =   96
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   645
         Index           =   2
         Left            =   9285
         TabIndex        =   95
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   645
         Index           =   1
         Left            =   9285
         TabIndex        =   94
         Top             =   2325
         Width           =   795
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "55"
         Height          =   645
         Index           =   0
         Left            =   9285
         TabIndex        =   93
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         Height          =   645
         Index           =   0
         Left            =   8490
         TabIndex        =   92
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         Height          =   645
         Index           =   3
         Left            =   8490
         TabIndex        =   91
         Top             =   3585
         Width           =   795
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         Height          =   645
         Index           =   2
         Left            =   8490
         TabIndex        =   90
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "45"
         Height          =   645
         Index           =   1
         Left            =   8490
         TabIndex        =   89
         Top             =   2325
         Width           =   795
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         Height          =   645
         Index           =   3
         Left            =   6125
         TabIndex        =   88
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         Height          =   645
         Index           =   2
         Left            =   6125
         TabIndex        =   87
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         Height          =   645
         Index           =   1
         Left            =   6125
         TabIndex        =   86
         Top             =   2325
         Width           =   795
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         Height          =   645
         Index           =   0
         Left            =   6125
         TabIndex        =   85
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   645
         Index           =   0
         Left            =   5340
         TabIndex        =   84
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   645
         Index           =   3
         Left            =   5340
         TabIndex        =   83
         Top             =   3585
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   645
         Index           =   2
         Left            =   5340
         TabIndex        =   82
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         Height          =   645
         Index           =   1
         Left            =   5340
         TabIndex        =   81
         Top             =   2330
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,54"
         Height          =   645
         Index           =   3
         Left            =   2955
         TabIndex        =   80
         Top             =   3600
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,54"
         Height          =   645
         Index           =   2
         Left            =   2960
         TabIndex        =   79
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,77"
         Height          =   645
         Index           =   1
         Left            =   2960
         TabIndex        =   78
         Top             =   2330
         Width           =   795
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,77"
         Height          =   645
         Index           =   0
         Left            =   2950
         TabIndex        =   77
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,63"
         Height          =   645
         Index           =   0
         Left            =   2160
         TabIndex        =   76
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,26"
         Height          =   645
         Index           =   3
         Left            =   2160
         TabIndex        =   75
         Top             =   3580
         Width           =   795
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,26"
         Height          =   645
         Index           =   2
         Left            =   2160
         TabIndex        =   74
         Top             =   3000
         Width           =   795
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,63"
         Height          =   645
         Index           =   1
         Left            =   2160
         TabIndex        =   73
         Top             =   2330
         Width           =   795
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accesorio enovolvente # 1 "
         Height          =   645
         Left            =   120
         TabIndex        =   72
         Top             =   2955
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accesorio plano # 1 "
         Height          =   645
         Left            =   120
         TabIndex        =   71
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accesorio plano # 2"
         Height          =   645
         Left            =   120
         TabIndex        =   70
         Top             =   2325
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accesorio envolvente # 2 "
         Height          =   645
         Left            =   120
         TabIndex        =   69
         Top             =   3585
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Veredicto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10065
         TabIndex        =   68
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "          Máxima Permitida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9270
         TabIndex        =   67
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Mínima permitida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   8475
         TabIndex        =   66
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "              Medida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7680
         TabIndex        =   65
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Veredicto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6915
         TabIndex        =   64
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "          Máxima Permitido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6120
         TabIndex        =   63
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Mínima permitido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   5325
         TabIndex        =   62
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         Medido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4530
         TabIndex        =   61
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "        Veredicto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3735
         TabIndex        =   60
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "          Máxima Permitida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2940
         TabIndex        =   59
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "      Mínima permitida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2145
         TabIndex        =   58
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                      Frecuencia (Hz)"
         Height          =   645
         Left            =   7680
         TabIndex        =   57
         Top             =   360
         Width           =   3200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                              Ancho de pulso positivo  ( ms) "
         Height          =   645
         Left            =   4520
         TabIndex        =   56
         Top             =   360
         Width           =   3195
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accseorio bajo medición"
         Height          =   645
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "              Medida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1350
         TabIndex        =   54
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                                                          Intensidad de corriente           (valor eficaz) (A~)"
         Height          =   645
         Left            =   1350
         TabIndex        =   46
         Top             =   360
         Width           =   3200
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "              Magintud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label85 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   10470
      TabIndex        =   101
      ToolTipText     =   "Número de Revisión"
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label86 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   100
      Top             =   900
      Width           =   195
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Técnico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   99
      Top             =   2265
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "REGMEDI07.frx":3276
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   600
      Picture         =   "REGMEDI07.frx":3580
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   98
      Top             =   1905
      Width           =   2880
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registro de Medición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   97
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6150
      TabIndex        =   53
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Veredicto General:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7920
      TabIndex        =   52
      Top             =   4150
      Width           =   2220
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6150
      TabIndex        =   51
      ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
      Top             =   2265
      Width           =   4695
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8880
      TabIndex        =   50
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label100 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6150
      TabIndex        =   45
      ToolTipText     =   "Doble Click Para Seleccionar Instrumento"
      Top             =   1905
      Width           =   4695
   End
   Begin VB.Label Label99 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comentarios:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   44
      Top             =   2625
      Width           =   1110
   End
   Begin VB.Label Label97 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Instrumentos:"
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
      Left            =   6120
      TabIndex        =   43
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label Label92 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   42
      Top             =   900
      Width           =   1380
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   41
      Top             =   1905
      Width           =   1380
   End
   Begin VB.Label Label88 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   1440
      TabIndex        =   40
      Top             =   1320
      Width           =   4305
   End
   Begin VB.Label Label96 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   39
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label95 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   38
      Top             =   960
      Width           =   2880
   End
End
Attribute VB_Name = "REGMEDI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '\\\OT-08-0936-OD-08/10/08///
    Call CIERRARCH("*.*")
    Unload Me
    '\\\OT-08-0936-OD-FIN///
End Sub

Private Sub Command2_Click()
   '\\\OT-08-0936-OD-08/10/08///
    Command2.Enabled = False
    '
    'VALIDACION POR FECHA
    
    VDEVV$ = TRIM$(Text11.TEXT)
    If FECHANUM(VDEVV$) < 1 Then
      Call MENSERR(24, "Fecha Inválida")
      GoTo 99
    End If
    '
    'VALIDACION DE NUMERO DE SERIE
    '
    NUSERIE$ = TRIM$(Text3)
    CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
    If CANTRE& < 1 Then
       Call MENSERR(24, "Número de Serie Incorrecto")
       GoTo 99
    End If
    '
    If TRIM$(Label100) = "" Then
      Call MENSERR(24, "Falta Informar Instrumento (1)")
      GoTo 99
    End If
    If TRIM$(Label11) = "" Then
      Call MENSERR(24, "Falta Informar Instrumento (2)")
      GoTo 99
    End If
    '
    'VALIDACION DE VEREDICTO GENERAL
    '
    If TRIM$(Text7) = "" Then
      Call MENSERR(24, "Falta Informar Veredicto General")
      GoTo 99
    End If
    
    If TRIM$(Label19) = "" Then
      Call MENSERR(24, "Falta Informar Inspector")
      GoTo 99
    End If
    ' Grabo en ensar074
    '
    Screen.MousePointer = 11
    '
'    SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR074 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'    SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'    Dim ENSA074TMP As ADODB.Recordset
'    'Set ENSA074TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'    Set ENSA074TMP = READSET(SQLZ$)
'    ULTREVIS% = XVALO(ENSA074TMP!MAXREVI)
'    ENSA074TMP.Close
'    Set ENSA074TMP = Nothing
'    '
'    SQLX$ = "SELECT REVISION FROM ENSAR074 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- VALIDAR EL  VALOR DE CAMPO REVISION
'    SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'    SQLX$ = SQLX$ + " AND REVISION =" & ULTREVIS% & ""
'    '
'    Dim ENSA074TMP1 As ADODB.Recordset '\\\OT-08-1064-OD-21/11/08-MEDITEA-BUSCO SI ES LA PRIMERA REVISION ,CASO CONTRARIO ULTIMA REVISION + 1
'    'Set ENSA074TMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    Set ENSA074TMP1 = READSET(SQLX$)
'    With ENSA074TMP1
'     If Not .EOF Then
'       REVIS% = 1 + ULTREVIS%
'     Else
'       REVIS% = 0 'SI ES EL PRIMERO EL VALOR VA A SER 0
'     End If
'    End With
'    ENSA074TMP1.Close
'    Set ENSA074TMP1 = Nothing
    '
    CONDI$ = "NumeroSerie = '" & NUSERIE$ & "'"
    CANTRE& = CantRegistros&("ENSAR074", CONDI$, "NOMESS")
    If CANTRE& = 0 Then
       REVIS% = 0
    Else
       REVIS% = CANTRE& 'NO LE SUMO UNO POR QUE ARRANCA DE 0 (EJ. SI DEVUELVE 0 REGISTROS VA A DAR DE ALTA EL NUMERO 0 SI DEVUELVE 1 VA A DAR DE ALTA EL 1
                        'POR QUE EL REGISTRO 1 CORRESPONDE AL 0
    End If
    '
    SQLX$ = "SELECT * FROM ENSAR074"
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NUSERIE$ & "'"
    SQLX$ = SQLX$ + " AND REVISION = " & REVIS%
    Dim ENSA4TMP As ADODB.Recordset
    Set ENSA4TMP = New ADODB.Recordset
    ENSA4TMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With ENSA4TMP
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !NumeroSerie = TRIM$(Text3)
            !MODELO = CODINT%(Label95)
            !Inspector = Left$(TRIM$(Label19), 32)
            !Comentarios = Text9
            !VerGeneral = Left$(TRIM$(Text7), 6)
            !Inten_Med1 = XVALO(FORMATNUM(Text1(0), "F5.3"))
            !Inten_Med2 = XVALO(FORMATNUM(Text1(1), "F5.3"))
            !Inten_Med3 = XVALO(FORMATNUM(Text1(2), "F5.3"))
            !Inten_Med4 = XVALO(FORMATNUM(Text1(3), "F5.3"))
            !Inten_Ver1 = Left$(TRIM$(Text2(0)), 6)
            !Inten_Ver2 = Left$(TRIM$(Text2(1)), 6)
            !Inten_Ver3 = Left$(TRIM$(Text2(2)), 6)
            !Inten_Ver4 = Left$(TRIM$(Text2(3)), 6)
            !AnPul_Med1 = XVALO(FORMATNUM(Text16(0), "F5.2"))
            !AnPul_Med2 = XVALO(FORMATNUM(Text16(1), "F5.2"))
            !AnPul_Med3 = XVALO(FORMATNUM(Text16(2), "F5.2"))
            !AnPul_Med4 = XVALO(FORMATNUM(Text16(3), "F5.2"))
            !AnPul_Ver1 = Left$(TRIM$(text17(0)), 6)
            !AnPul_Ver2 = Left$(TRIM$(text17(1)), 6)
            !AnPul_Ver3 = Left$(TRIM$(text17(2)), 6)
            !AnPul_Ver4 = Left$(TRIM$(text17(3)), 6)
            !Frec_Med1 = XVALO(FORMATNUM(Text18(0), "F5.2"))
            !Frec_Med2 = XVALO(FORMATNUM(Text18(1), "F5.2"))
            !Frec_Med3 = XVALO(FORMATNUM(Text18(2), "F5.2"))
            !Frec_Med4 = XVALO(FORMATNUM(Text18(3), "F5.2"))
            !Frec_Ver1 = Left$(TRIM$(text19(0)), 6)
            !Frec_Ver2 = Left$(TRIM$(text19(1)), 6)
            !Frec_Ver3 = Left$(TRIM$(text19(2)), 6)
            !Frec_Ver4 = Left$(TRIM$(text19(3)), 6)
            !Fech_Ensa = Text11
            !Bloqueado = "SI"
            !TECNICO = Left$(TRIM$(Text4), 32) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-GRABA EL TECNICO
            !Revision = REVIS% ' GRABA CAMPO DE REVISION
            !INSTRUMENTO1 = Left$(Label100, 32)
            !INSTRUMENTO2 = Left$(Label11, 32)
            .Update
    End With
    '
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    CUERPO$ = "Registro de Medición" & vbCrLf
    CUERPO$ = CUERPO$ + "Técnico: " & UCase$(TRIM$(Text4))
    CUERPO$ = CUERPO$ + " - Resultado: " & UCase$(TRIM$(Text7))
    Call ANOTANUSE(NUSERIE$, CUERPO$)
    '
    ENSA4TMP.Close
    Set ENSA4TMP = Nothing
    '
    Call PRINTMEDI(1)
    '
    Image1.Visible = True
    Image3.Visible = False
    Command23.Enabled = True
    '
    Call BLANFORMU
    '
    Screen.MousePointer = 1

99  Command2.Enabled = True
   '\\\OT-08-0936-OD-fin///

End Sub


Private Sub Command23_Click()
   NUSERIE$ = TRIM$(Text3)
   CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
   If CANTRE& < 1 Then
      Call MENSERR(24, "Número de Serie Incorrecto")
      Exit Sub
   End If
   Call PRINTMEDI(0)
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   FSTPRDM.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command4_Click()
   '\\\OT-08-0936-OD-08/10/08///
   Call Text3_DblClick
   '\\\OT-08-0936-OD-FIN///
End Sub

Private Sub Command6_Click()
   '\\\OT-08-0936-OD-08/10/08///
    If Image3.Visible = False Then
      Call CommSel("SELFECHA", Command6, Text11)
    Else
      Command4.SetFocus
    End If
    '\\\OT-08-0936-OD-FIN///
End Sub


Private Sub Command7_Click()
   Command7.Enabled = False
   Call COFUTIE07.Command7_Click
   Command7.Enabled = True

End Sub

Private Sub Form_Load()
    '\\\OT-08-0936-OD-06/10/08///
    Call ABREORFAB
    FF% = HOY(HO$)
    Text11.TEXT = HO$
    Label19 = USERNAME$
    '\\\OT-08-0936-OD-FIN///
End Sub

Private Sub Labe46_Click(Index As Integer)

End Sub

Private Sub Labe48_Click(Index As Integer)

End Sub

Private Sub Image1_Click()
   If DERACCE("MODENS72", "Modificacion  de Registro de Medición") > 1 Then
        Image1.Visible = False
        Image3.Visible = True
        Command2.Enabled = False
        Command23.Enabled = True  '\\\OT-08-1064-OD-21/11/08-MEDITEA- HABILITA LA IMPRESORA
        If TRIM$(Text3) <> "" Then
          'PONGO LA FECHA DEL ULTIMO INFORME
          NUSERIE$ = TRIM$(Text3)
          CONDI$ = "NumeroSerie ='" & NUSERIE$ & "'"
          ULTIMOENSA% = XVALO(VALOBADA("ENSAR074", "MAX(REVISION)", CONDI$))
          CONDI$ = CONDI$ + " AND REVISION = " & ULTIMOENSA%
          Text11 = FECHATEX(CVINT(VALOBADA("ENSAR074", "Fech_Ensa", CONDI$, "NOMESS")))
        End If

   End If

End Sub

Private Sub Image3_Click()
   If DERACCE("MODENS74", "Modificacion  de Registro de Medición") > 1 Then
        Image1.Visible = True
        Image3.Visible = False
        Command2.Enabled = True
        Command23.Enabled = False
        'PONGO FECHA DEL DIA CUANDO VA A REALIZAR OTRA REVISION
        HOII% = HOY(HOS$)
        Text11 = HOS$

   End If

End Sub

Private Sub Label100_DblClick()
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label100 = TRIM$(ECOARCH$("INSTRUME", Chr$(CODINSTRU%)))
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("INSTRUME")
    CODINSTRU% = XVALO(VALACT1$("INSTRUME"))
    If CODINSTRU% < 1 Then Exit Sub
    Label11 = TRIM$(ECOARCH$("INSTRUME", Chr$(CODINSTRU%)))

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  '\\\OT-08-0936-OD-08/10/08///
  VALOR1 = ROUND#(XVALO(Text1(Index)), 3)
  
  If XVALO(VALOR1) < 1000 Then
    Text1(Index) = FORMATNUM(Text1(Index), "F7.3")
    '
    REFE1 = XVALO(Label38(Index))
    REFE2 = XVALO(Label9(Index))
    VALOR = XVALO(Text1(Index))
    Text2(Index) = ""
    If XVALO(Text1(Index)) > 0 Then '\\\OT-08-1066-OD-21/11/08-MEDITEA-CAMBIAR A NC SI NO ESTA DENTRO DE LOS PARAMETROS
        If VALOR >= REFE1 And VALOR <= REFE2 Then
          Text2(Index) = "A"
        Else
          Text2(Index) = "NC"
        End If
   End If
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 1000.\Debe Corregir Para Poder Continuar.")
    Text1(Index).SetFocus
  End If
  '\\\OT-08-0936-OD-FIN///

End Sub

Private Sub Text11_DblClick()
   Call Command6_Click
End Sub

Private Sub Text11_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text16_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text16_LostFocus(Index As Integer)
  '\\\OT-08-0936-OD-08/10/08///
  VALOR1 = ROUND#(XVALO(Text16(Index)), 3)
  
  If XVALO(VALOR1) < 1000 Then
    Text16(Index) = FORMATNUM(Text16(Index), "F7.3")
    REFE1 = XVALO(Label45(Index))
    REFE2 = XVALO(Label46(Index))
    VALOR = XVALO(Text16(Index))
    text17(Index) = ""
    If XVALO(Text16(Index)) > 0 Then  '\\\OT-08-1066-OD-21/11/08-MEDITEA-CAMBIAR A NC SI NO ESTA DENTRO DE LOS PARAMETROS
        If VALOR >= REFE1 And VALOR <= REFE2 Then
          text17(Index) = "A"
        Else
          text17(Index) = "NC"
        End If
    End If
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 1000.\Debe Corregir Para Poder Continuar.")
    Text16(Index).SetFocus
  End If
  '\\\OT-08-0936-OD-FIN///

End Sub

Private Sub text17_Change(Index As Integer)
   Call VeredGeneral '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR DE  LA COLUMNA VEREDICTO
End Sub

Private Sub text17_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text18_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus

End Sub

Private Sub Text18_LostFocus(Index As Integer)
  '\\\OT-08-0936-OD-08/10/08///
  VALOR1 = ROUND#(XVALO(Text18(Index)), 3)
  
  If XVALO(VALOR1) < 1000 Then
    Text18(Index) = FORMATNUM(Text18(Index), "F7.3")
    
    REFE1 = XVALO(Label47(Index))
    REFE2 = XVALO(Label48(Index))
    VALOR = XVALO(Text18(Index))
    text19(Index) = ""
    If XVALO(Text18(Index)) > 0 Then   '\\\OT-08-1066-OD-21/11/08-MEDITEA-CAMBIAR A NC SI NO ESTA DENTRO DE LOS PARAMETROS
        If VALOR >= REFE1 And VALOR <= REFE2 Then
          text19(Index) = "A"
        Else
          text19(Index) = "NC"
        End If
    End If
  Else
    Call MENSERR(24, "Sólo Se Aceptan Valores Menores a 1000.\Debe Corregir Para Poder Continuar.")
    Text18(Index).SetFocus
  End If
  '\\\OT-08-0936-OD-FIN///

End Sub


Private Sub text19_Change(Index As Integer)
   Call VeredGeneral  '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR DE  LA COLUMNA VEREDICTO
End Sub

Private Sub text19_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text2_Change(Index As Integer)
  Call VeredGeneral  '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR DE  LA COLUMNA VEREDICTO
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text3_Change()
   '\\\OT-08-0936-OD-08/10/08///
   Command2.Enabled = True
   Image3.Visible = False
   Image1.Enabled = False
   Image1.Visible = True
   Command23.Enabled = True
   
   Call LIMPIAINFORMACION
   '\\\OT-08-0936-OD-08/10/08///

End Sub


Private Sub Text3_DblClick()
   '\\\OT-08-0936-OD-08/10/08///
   Call SELECHO("INDINUSE")
   TTXX$ = TRIM$(VALACT1$("INDINUSE"))
   If TTXX$ <> "" Then
      Text3 = TTXX$
      Call Text3_KeyPress(13)
   End If
   
  '\\\OT-08-0936-OD-08/10/08///
End Sub
'
Sub PRINTMEDI(N%)
    '
    REVIS% = XVALO(Label85) ' SI VIENE DE IMPRESION SIMPLE TOMA DEL LABEL
    If N% > 0 Then
        REVIS% = 0 ' SI VIENE DE GRABACION BUSCA EL  VALOR MAXIMO
        NUSERIE$ = TRIM$(Text3)
        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR074 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        Dim MAXREVTMP As ADODB.Recordset
        'Set MAXREVTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
        Set MAXREVTMP = READSET(SQLZ$)
        ULTREVIS% = XVALO(MAXREVTMP!MAXREVI)
        MAXREVTMP.Close
        Set MAXREVTMP = Nothing
        REVIS% = ULTREVIS%
    End If
    
    TIPODOC$ = "Registro de Medicion" ' guardar el documento immpreso
    NUMEDOC$ = TRIM$(Text3) & "-" & REVIS%

    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = TRIM$(Text11) 'FECHA
    CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = Label88.Caption 'DESCRIPCION DE MATERIAL
    CODPARAM$(3) = "REF-MAT1":  DOCPARAM$(3) = TRIM$(Text3.TEXT) 'NRO SERIE
    CODPARAM$(4) = "REF-MAT2":  DOCPARAM$(4) = TRIM$(Text7.TEXT)  'VEREDICTO
    CODPARAM$(5) = "REF-MAT3":  DOCPARAM$(5) = TRIM$(Label19)  'Inspector
    CODPARAM$(6) = "OBSERVA":  DOCPARAM$(6) = TRIM$(Text9.TEXT)  'Observaciones
    CODPARAM$(7) = "REF-MAT4":  DOCPARAM$(7) = TRIM$(Label100)  'INSTRUMENTO 1
    CODPARAM$(8) = "REF-MAT5":  DOCPARAM$(8) = TRIM$(Label11)   'INSTRUMENTO 2
    '
    CAPARDOC% = 8
    '
    CODTABU1$(1) = "MEDIDA-1"
    CODTABU1$(2) = "N-PLANO"
    CODTABU1$(3) = "MEDIDA-2"
    CODTABU1$(4) = "R-PLANO"
    CODTABU1$(5) = "MEDIDA-3"
    CODTABU1$(6) = "F-PLANO"
    CACOLTAB1% = 6
    '
    CAITAB1% = 0
    For U& = 1 To 4
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Text1(U& - 1))    'MEDIDA
        TETABU1$(2, CAITAB1%) = UCase(TRIM$(Text2(U& - 1)))    ' VEREDICTO
        TETABU1$(3, CAITAB1%) = TRIM$(Text16(U& - 1))   'MEDIDA
        TETABU1$(4, CAITAB1%) = UCase(TRIM$(text17(U& - 1)))   ' VEREDICTO
        TETABU1$(5, CAITAB1%) = TRIM$(Text18(U& - 1))   'MEDIDA
        TETABU1$(6, CAITAB1%) = UCase(TRIM$(text19(U& - 1)))   ' VEREDICTO
    Next U&
    
    '
    FORIPRE$ = TRIM$(Control$("", "REGMEDI"))
    Call PRINTDOC("@" + FORIPRE$)
End Sub
Sub BLANFORMU()
    For I& = 1 To 4
        Text1(I& - 1) = ""
        Text2(I& - 1) = ""
        Text16(I& - 1) = ""
        text17(I& - 1) = ""
        Text18(I& - 1) = ""
        text19(I& - 1) = ""
    Next I&
    Text3 = ""
    Text4 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNICO
    Text7 = ""
    Text9 = ""
    Label85 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Label100 = ""
    Label11 = ""
    Text11 = FECHATEX$(HOY(HO$))
End Sub
Sub LIMPIAINFORMACION()
    For I& = 1 To 4
        Text1(I& - 1) = ""
        Text2(I& - 1) = ""
        Text16(I& - 1) = ""
        text17(I& - 1) = ""
        Text18(I& - 1) = ""
        text19(I& - 1) = ""
    Next I&
    Text4 = "" '\\\OT-08-1065-OD-21/11/08-MEIDTEA-LIMPIA EL TEXT DEL TECNICO
    Text7 = ""
    Text9 = ""
    Label85 = "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-LIMPIO EL LABEL DE REVISION
    Label88 = ""
    Label95 = ""
    Label100 = ""
    Label11 = ""

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       NUSERIE$ = TRIM$(Text3)
       If InStr(NUSERIE$, "'") < 1 Then
        'VALIDACION DE NUMERO DE SERIE
        '
        CANTRE& = CantRegistros&("TANUMSE", "NumeroSerie ='" & NUSERIE$ & "'")
        If CANTRE& < 1 Then
           Call MENSERR(24, "Número de Serie Incorrecto")
           Exit Sub
        End If

    
        NUSERIE$ = TRIM$(Text3)
        CI% = XVALO(VALOBADA("TANUMSE", "Cod_Inte", "NumeroSerie ='" & NUSERIE$ & "'"))
        If CI% > 0 Then
         Label95 = CODEXT$(CI%)
         Label88 = DESCRIT0$(CI%)
        End If
        '
'        REVIS% = 0
'        SQLZ$ = "SELECT MAX(REVISION)AS MAXREVI FROM ENSAR074 WITH(NOLOCK)" '\\\OT-08-1064-OD-21/11/08-MEDITEA- MAXIMO VALOR DE CAMPO REVISION
'        SQLZ$ = SQLZ$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
'        Dim ENSA074TMP As ADODB.Recordset
'        'Set ENSA074TMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
'        Set ENSA074TMP = READSET(SQLZ$)
'        ULTREVIS% = XVALO(ENSA074TMP!MAXREVI)
'        ENSA074TMP.Close
'        Set ENSA074TMP = Nothing
        Label85 = ULTREVIS%
        '
        BLQUEO$ = ""
        SQLX$ = "SELECT TOP 1 * FROM ENSAR074 WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NumeroSerie ='" & NUSERIE$ & "'"
        'SQLX$ = SQLX$ + " AND REVISION = " & ULTREVIS% & "" '\\\OT-08-1064-OD-21/11/08-MEDITEA-AGREGO EL NUMERO DE REVISION COMO PARTE DE LA CONSULTA
        SQLX$ = SQLX$ + " ORDER BY REVISION DESC  "
        Dim ENSA4TMP As ADODB.Recordset
        'Set ENSA4TMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Set ENSA4TMP = READSET(SQLX$)
        With ENSA4TMP
             If Not .EOF Then
                Label85 = XVALO(!Revision)
                Text3 = SAFETEXT(!NumeroSerie)
                CI% = !MODELO
                Label95 = CODEXT$(CI%)
                Label88 = DESCRIT0$(CI%)
                Label19 = SAFETEXT$(!Inspector)
                Text9 = SAFETEXT$(!Comentarios)
                Text1(0) = FORMATNUM(XVALO(!Inten_Med1), "F7.3")
                Text1(1) = FORMATNUM(XVALO(!Inten_Med2), "F7.3")
                Text1(2) = FORMATNUM(XVALO(!Inten_Med3), "F7.3")
                Text1(3) = FORMATNUM(XVALO(!Inten_Med4), "F7.3")
                Text2(0) = SAFETEXT$(!Inten_Ver1)
                Text2(1) = SAFETEXT$(!Inten_Ver2)
                Text2(2) = SAFETEXT$(!Inten_Ver3)
                Text2(3) = SAFETEXT$(!Inten_Ver4)
                Text16(0) = FORMATNUM(XVALO(!AnPul_Med1), "F7.3")
                Text16(1) = FORMATNUM(XVALO(!AnPul_Med2), "F7.3")
                Text16(2) = FORMATNUM(XVALO(!AnPul_Med3), "F7.3")
                Text16(3) = FORMATNUM(XVALO(!AnPul_Med4), "F7.3")
                text17(0) = SAFETEXT$(!AnPul_Ver1)
                text17(1) = SAFETEXT$(!AnPul_Ver2)
                text17(2) = SAFETEXT$(!AnPul_Ver3)
                text17(3) = SAFETEXT$(!AnPul_Ver4)
                Text18(0) = FORMATNUM(XVALO(!Frec_Med1), "F7.3")
                Text18(1) = FORMATNUM(XVALO(!Frec_Med2), "F7.3")
                Text18(2) = FORMATNUM(XVALO(!Frec_Med3), "F7.3")
                Text18(3) = FORMATNUM(XVALO(!Frec_Med4), "F7.3")
                text19(0) = SAFETEXT$(!Frec_Ver1)
                text19(1) = SAFETEXT$(!Frec_Ver2)
                text19(2) = SAFETEXT$(!Frec_Ver3)
                text19(3) = SAFETEXT$(!Frec_Ver4)
                Text7 = SAFETEXT$(!VerGeneral)
                Text4 = SAFETEXT$(!TECNICO) '\\\OT-08-1065-OD-21/11/08-MEIDTEA-CARGA EL TEXT DEL TECNICO
                Label100 = SAFETEXT$(!INSTRUMENTO1)
                Label11 = SAFETEXT$(!INSTRUMENTO2)
                fecccha$ = SAFETEXT$(!Fech_Ensa) ' si no tiene fecha pone la de hoy
                If fecccha$ <> "" Then fff% = CVINT(fecccha$): Text11 = FECHATEX(fff%)
                'si tiene fecha le damos formato "dd/mm/aa"
                If TRIM$(fecccha$) = "" Then FF% = HOY(HO$): Text11.TEXT = HO$
                BLQUEO$ = SAFETEXT(!Bloqueado)
             End If
        End With
        ENSA4TMP.Close
        Set ENSA4TMP = Nothing
       End If
      
      If BLQUEO$ = "SI" Then
            Image3.Visible = True
            Image1.Visible = False
            Command2.Enabled = False
            Image1.Enabled = True
      Else
            Command23.Enabled = False '\\\OT-08-1064-OD-21/11/08-MEIDTEA BLOQUEA LA IMPRESION
          
      End If

   End If
End Sub

Private Sub Text4_GotFocus()
  If Image3.Visible = True Then Command4.SetFocus '\\\OT-08-1066-OD-21/11/08-MEDITEA QUE NO SE PUEDA EDITAR SI TIENE EL CANDADO VISIBLE
End Sub


Private Sub Text7_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub

Private Sub Text9_GotFocus()
   If Image3.Visible = True Then Command4.SetFocus
End Sub
Sub VeredGeneral()
  '
  CONTAD% = 0
  '
  For I& = 0 To 3 '\\\OT-08-1066-OD-21/11/08-MEDITEA-VALIDAR SI HAY UN NC VEREDICTO SEA NC
    If UCase(TRIM$(Text2(I&))) = "NC" Then 'RECORRE LA PRIMERA COLUMNA DE VEREDICTO EN BUSCA DE NC
      Text7 = "NC"
      Exit Sub
    End If
  Next I&
  '
  For U& = 0 To 3
    If UCase(TRIM$(text17(U&))) = "NC" Then 'RECORRE LA SEGUNDA COLUMNA DE VEREDICTO EN BUSCA DE NC
      Text7 = "NC"
      Exit Sub
    End If
  Next U&
  '
  For j& = 0 To 3
    If UCase(TRIM$(text19(j&))) = "NC" Then 'RECORRE LA TERCERA COLUMNA DE VEREDICTO EN BUSCA DE NC
      Text7 = "NC"
      Exit Sub
    End If
  Next j&
  '
  For Z& = 0 To 3
    If UCase(TRIM$(Text2(Z&))) = "A" Then 'RECORRE LA PRIMERA COLUMNA DE VEREDICTO EN BUSCA DE A
    CONTAD% = CONTAD% + 1
    End If
  Next Z&
  '
  For X& = 0 To 3
    If UCase(TRIM$(text17(X&))) = "A" Then 'RECORRE LA SEGUNDA COLUMNA DE VEREDICTO EN BUSCA DE A
    CONTAD% = CONTAD% + 1
    End If
  Next X&
  '
  For W& = 0 To 3
    If UCase(TRIM$(text19(W&))) = "A" Then 'RECORRE LA TERCERA COLUMNA DE VEREDICTO EN BUSCA DE A
    CONTAD% = CONTAD% + 1
    End If
  Next W&
  '
  If CONTAD% = 12 Then 'VALIDAR SI TODOS TIENEN A  VEREDICTO SEA A
    Text7 = "A"
  Else
    Text7 = ""
  End If

End Sub

