VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CIERREOP10 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Produccion"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
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
   ScaleHeight     =   7545
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   165
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Materiales Consumidos"
      Height          =   2685
      Left            =   120
      TabIndex        =   56
      Top             =   2640
      Width           =   9375
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   2100
         Left            =   5235
         TabIndex        =   62
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   2100
         Left            =   7245
         TabIndex        =   61
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   2100
         Left            =   6720
         TabIndex        =   60
         Top             =   390
         Width           =   30
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2100
         Left            =   1860
         TabIndex        =   59
         Top             =   375
         Width           =   30
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   175
         Sorted          =   -1  'True
         TabIndex        =   57
         Top             =   645
         Width           =   9000
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo                        Descripcion                                     Id.Bulto            U-M           Cantidad"
         Height          =   270
         Left            =   175
         TabIndex        =   58
         Top             =   375
         Width           =   9000
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "No Conformes"
      Height          =   1000
      Left            =   4860
      TabIndex        =   27
      Top             =   1600
      Width           =   1845
      Begin VB.CommandButton Command2 
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
         Left            =   1530
         TabIndex        =   29
         Top             =   555
         Width           =   175
      End
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   210
         TabIndex        =   30
         Top             =   570
         Width           =   1350
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   210
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1350
      TabIndex        =   51
      Top             =   1740
      Width           =   600
   End
   Begin VB.CommandButton Command30 
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
      Left            =   1935
      TabIndex        =   50
      Top             =   1725
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Aprobado"
      Height          =   1000
      Left            =   3030
      TabIndex        =   44
      Top             =   1600
      Width           =   1860
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   47
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   46
         Top             =   570
         Width           =   1320
      End
      Begin VB.CommandButton Command25 
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
         Left            =   1530
         TabIndex        =   45
         Top             =   570
         Width           =   175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mermas y Devolucion de Materiales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   9405
      Begin VB.PictureBox Picture1 
         Height          =   1350
         Left            =   120
         ScaleHeight     =   1290
         ScaleWidth      =   9075
         TabIndex        =   35
         Top             =   525
         Width           =   9135
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  'None
            Height          =   3375
            Left            =   -1400
            ScaleHeight     =   3375
            ScaleWidth      =   10305
            TabIndex        =   64
            Top             =   -100
            Width           =   10300
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   8655
               TabIndex        =   118
               Top             =   285
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   8655
               TabIndex        =   117
               Top             =   600
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   8655
               TabIndex        =   116
               Top             =   915
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   8655
               TabIndex        =   115
               Top             =   1230
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   8655
               TabIndex        =   114
               Top             =   2175
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   8655
               TabIndex        =   113
               Top             =   1860
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   8655
               TabIndex        =   112
               Top             =   1545
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   8655
               TabIndex        =   111
               Top             =   2490
               Width           =   1305
            End
            Begin VB.TextBox text4 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   8655
               TabIndex        =   110
               Top             =   2805
               Width           =   1305
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
               Height          =   285
               Index           =   3
               Left            =   8325
               TabIndex        =   109
               Top             =   1230
               Width           =   175
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
               Height          =   285
               Index           =   4
               Left            =   8325
               TabIndex        =   108
               Top             =   1545
               Width           =   175
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
               Height          =   285
               Index           =   5
               Left            =   8325
               TabIndex        =   107
               Top             =   1845
               Width           =   175
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
               Height          =   285
               Index           =   6
               Left            =   8325
               TabIndex        =   106
               Top             =   2175
               Width           =   175
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
               Height          =   285
               Index           =   7
               Left            =   8325
               TabIndex        =   105
               Top             =   2490
               Width           =   175
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
               Height          =   285
               Index           =   8
               Left            =   8325
               TabIndex        =   104
               Top             =   2805
               Width           =   175
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
               Height          =   285
               Index           =   2
               Left            =   8325
               TabIndex        =   103
               Top             =   915
               Width           =   175
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
               Height          =   285
               Index           =   1
               Left            =   8310
               TabIndex        =   102
               Top             =   600
               Width           =   175
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
               Height          =   285
               Index           =   0
               Left            =   8310
               TabIndex        =   101
               Top             =   285
               Width           =   175
            End
            Begin VB.CommandButton Command18 
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
               Index           =   8
               Left            =   3195
               TabIndex        =   100
               Top             =   2805
               Width           =   175
            End
            Begin VB.CommandButton Command19 
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
               Index           =   8
               Left            =   4905
               TabIndex        =   99
               Top             =   2805
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   8
               Left            =   5880
               TabIndex        =   98
               Top             =   2805
               Width           =   1305
            End
            Begin VB.CommandButton Command18 
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
               Index           =   7
               Left            =   3210
               TabIndex        =   97
               Top             =   2490
               Width           =   175
            End
            Begin VB.CommandButton Command19 
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
               Index           =   7
               Left            =   4905
               TabIndex        =   96
               Top             =   2490
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   7
               Left            =   5880
               TabIndex        =   95
               Top             =   2490
               Width           =   1305
            End
            Begin VB.CommandButton Command18 
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
               Index           =   4
               Left            =   3210
               TabIndex        =   94
               Top             =   1545
               Width           =   175
            End
            Begin VB.CommandButton Command19 
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
               Index           =   4
               Left            =   4905
               TabIndex        =   93
               Top             =   1545
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   5880
               TabIndex        =   92
               Top             =   1545
               Width           =   1305
            End
            Begin VB.CommandButton Command18 
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
               Index           =   5
               Left            =   3210
               TabIndex        =   91
               Top             =   1860
               Width           =   175
            End
            Begin VB.CommandButton Command19 
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
               Index           =   5
               Left            =   4905
               TabIndex        =   90
               Top             =   1860
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   5880
               TabIndex        =   89
               Top             =   1860
               Width           =   1305
            End
            Begin VB.CommandButton Command18 
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
               Index           =   6
               Left            =   3210
               TabIndex        =   88
               Top             =   2175
               Width           =   175
            End
            Begin VB.CommandButton Command19 
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
               Index           =   6
               Left            =   4905
               TabIndex        =   87
               Top             =   2175
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   5880
               TabIndex        =   86
               Top             =   2175
               Width           =   1305
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   5880
               TabIndex        =   85
               Top             =   1230
               Width           =   1305
            End
            Begin VB.CommandButton Command19 
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
               Index           =   3
               Left            =   4905
               TabIndex        =   84
               Top             =   1230
               Width           =   175
            End
            Begin VB.CommandButton Command18 
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
               Index           =   3
               Left            =   3210
               TabIndex        =   83
               Top             =   1230
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   5880
               TabIndex        =   82
               Top             =   915
               Width           =   1305
            End
            Begin VB.CommandButton Command19 
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
               Index           =   2
               Left            =   4905
               TabIndex        =   81
               Top             =   915
               Width           =   175
            End
            Begin VB.CommandButton Command18 
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
               Index           =   2
               Left            =   3210
               TabIndex        =   80
               Top             =   915
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   5880
               TabIndex        =   79
               Top             =   600
               Width           =   1305
            End
            Begin VB.CommandButton Command19 
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
               Index           =   1
               Left            =   4905
               TabIndex        =   78
               Top             =   600
               Width           =   175
            End
            Begin VB.CommandButton Command18 
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
               Index           =   1
               Left            =   3210
               TabIndex        =   77
               Top             =   600
               Width           =   175
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   5880
               TabIndex        =   76
               Top             =   285
               Width           =   1305
            End
            Begin VB.CommandButton Command19 
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
               Index           =   0
               Left            =   4905
               TabIndex        =   75
               Top             =   285
               Width           =   175
            End
            Begin VB.CommandButton Command18 
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
               Index           =   0
               Left            =   3210
               TabIndex        =   74
               Top             =   285
               Width           =   175
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   0
               Left            =   10080
               TabIndex        =   73
               Top             =   385
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   1
               Left            =   10080
               TabIndex        =   72
               Top             =   700
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   2
               Left            =   10080
               TabIndex        =   71
               Top             =   1015
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   3
               Left            =   10080
               TabIndex        =   70
               Top             =   1330
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   4
               Left            =   10080
               TabIndex        =   69
               Top             =   1645
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   5
               Left            =   10080
               TabIndex        =   68
               Top             =   1960
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   6
               Left            =   10080
               TabIndex        =   67
               Top             =   2275
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   7
               Left            =   10080
               TabIndex        =   66
               Top             =   2590
               Width           =   200
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   200
               Index           =   8
               Left            =   10080
               TabIndex        =   65
               Top             =   2905
               Width           =   200
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   8
               Left            =   1530
               TabIndex        =   154
               ToolTipText     =   "Código"
               Top             =   2805
               Width           =   1755
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   8
               Left            =   3540
               TabIndex        =   153
               ToolTipText     =   "Código"
               Top             =   2805
               Width           =   1440
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   8
               Left            =   7275
               TabIndex        =   152
               Top             =   2805
               Width           =   1050
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   8
               Left            =   5280
               TabIndex        =   151
               Top             =   2805
               Width           =   420
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   7
               Left            =   1530
               TabIndex        =   150
               ToolTipText     =   "Código"
               Top             =   2490
               Width           =   1755
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   7
               Left            =   3540
               TabIndex        =   149
               ToolTipText     =   "Código"
               Top             =   2490
               Width           =   1440
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   7
               Left            =   7275
               TabIndex        =   148
               Top             =   2490
               Width           =   1050
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   7
               Left            =   5280
               TabIndex        =   147
               Top             =   2490
               Width           =   420
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   4
               Left            =   1530
               TabIndex        =   146
               ToolTipText     =   "Código"
               Top             =   1545
               Width           =   1755
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   4
               Left            =   3540
               TabIndex        =   145
               ToolTipText     =   "Código"
               Top             =   1545
               Width           =   1440
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   4
               Left            =   7275
               TabIndex        =   144
               Top             =   1545
               Width           =   1050
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   4
               Left            =   5280
               TabIndex        =   143
               Top             =   1545
               Width           =   420
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   5
               Left            =   1530
               TabIndex        =   142
               ToolTipText     =   "Código"
               Top             =   1860
               Width           =   1755
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   5
               Left            =   3540
               TabIndex        =   141
               ToolTipText     =   "Código"
               Top             =   1860
               Width           =   1440
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   5
               Left            =   7275
               TabIndex        =   140
               Top             =   1860
               Width           =   1050
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   5
               Left            =   5280
               TabIndex        =   139
               Top             =   1860
               Width           =   420
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   6
               Left            =   1530
               TabIndex        =   138
               ToolTipText     =   "Código"
               Top             =   2175
               Width           =   1755
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   6
               Left            =   3540
               TabIndex        =   137
               ToolTipText     =   "Código"
               Top             =   2175
               Width           =   1440
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   6
               Left            =   7275
               TabIndex        =   136
               Top             =   2175
               Width           =   1050
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   6
               Left            =   5280
               TabIndex        =   135
               Top             =   2175
               Width           =   420
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   3
               Left            =   5280
               TabIndex        =   134
               Top             =   1230
               Width           =   420
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   3
               Left            =   7275
               TabIndex        =   133
               Top             =   1230
               Width           =   1050
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   3
               Left            =   3540
               TabIndex        =   132
               ToolTipText     =   "Código"
               Top             =   1230
               Width           =   1440
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   3
               Left            =   1530
               TabIndex        =   131
               ToolTipText     =   "Código"
               Top             =   1230
               Width           =   1755
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   5280
               TabIndex        =   130
               Top             =   915
               Width           =   420
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   7275
               TabIndex        =   129
               Top             =   915
               Width           =   1050
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   3540
               TabIndex        =   128
               ToolTipText     =   "Código"
               Top             =   915
               Width           =   1440
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   2
               Left            =   1530
               TabIndex        =   127
               ToolTipText     =   "Código"
               Top             =   915
               Width           =   1755
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   126
               Top             =   600
               Width           =   420
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   1
               Left            =   7275
               TabIndex        =   125
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   1
               Left            =   3540
               TabIndex        =   124
               ToolTipText     =   "Código"
               Top             =   600
               Width           =   1440
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   1
               Left            =   1530
               TabIndex        =   123
               ToolTipText     =   "Código"
               Top             =   600
               Width           =   1755
            End
            Begin VB.Label Label42 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   5280
               TabIndex        =   122
               Top             =   285
               Width           =   420
            End
            Begin VB.Label Label41 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   7275
               TabIndex        =   121
               Top             =   285
               Width           =   1050
            End
            Begin VB.Label Label40 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   3540
               TabIndex        =   120
               ToolTipText     =   "Código"
               Top             =   285
               Width           =   1440
            End
            Begin VB.Label Label39 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Index           =   0
               Left            =   1530
               TabIndex        =   119
               ToolTipText     =   "Código"
               Top             =   285
               Width           =   1755
            End
         End
         Begin VB.VScrollBar VScroll3 
            Height          =   1320
            Left            =   8865
            Max             =   340
            TabIndex        =   36
            Top             =   -15
            Width           =   225
         End
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   164
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   163
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   162
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   161
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   160
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   159
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   158
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   157
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   156
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   155
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mermas:"
         Height          =   220
         Left            =   7230
         TabIndex        =   63
         Top             =   300
         Width           =   1110
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         Height          =   220
         Left            =   5925
         TabIndex        =   41
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución:"
         Height          =   220
         Left            =   4650
         TabIndex        =   40
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         Height          =   220
         Left            =   345
         TabIndex        =   39
         Top             =   330
         Width           =   675
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "     Id.Bulto:"
         Height          =   220
         Index           =   1
         Left            =   2010
         TabIndex        =   38
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         Height          =   220
         Left            =   3495
         TabIndex        =   37
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cantidad"
      Height          =   2505
      Left            =   6810
      TabIndex        =   15
      Top             =   105
      Width           =   2685
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1000
         TabIndex        =   23
         Top             =   2080
         Width           =   1500
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1000
         TabIndex        =   33
         Top             =   1725
         Width           =   1500
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1000
         TabIndex        =   32
         ToolTipText     =   "Proceso Control de Producto"
         Top             =   1380
         Width           =   1500
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1000
         TabIndex        =   31
         Top             =   1035
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PCP:"
         Height          =   300
         Left            =   -400
         TabIndex        =   26
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Prod.:"
         Height          =   225
         Left            =   -400
         TabIndex        =   25
         Top             =   1785
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U/Bultos:"
         Height          =   300
         Left            =   -400
         TabIndex        =   24
         Top             =   2145
         Width           =   1365
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1000
         TabIndex        =   20
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1000
         TabIndex        =   19
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
         Height          =   300
         Left            =   -405
         TabIndex        =   18
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyect:"
         Height          =   300
         Left            =   -400
         TabIndex        =   17
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acumul:"
         Height          =   300
         Left            =   -400
         TabIndex        =   16
         Top             =   735
         Width           =   1365
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
      Height          =   285
      Left            =   6510
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.CommandButton Command5 
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
      Left            =   3150
      TabIndex        =   0
      Top             =   195
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7605
      Left            =   9555
      ScaleHeight     =   7545
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command14 
         Height          =   640
         Left            =   120
         Picture         =   "CIERREOP10.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2460
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Conc."
         Height          =   555
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Conciliacion"
         Top             =   1770
         Width           =   645
      End
      Begin VB.CommandButton Command15 
         Height          =   600
         Left            =   120
         Picture         =   "CIERREOP10.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1050
         Width           =   645
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   1005
         Left            =   120
         Picture         =   "CIERREOP10.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar Cierre de O/F y Reingreso de Materiales a O/F"
         Top             =   6360
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
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   4575
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
            Left            =   -135
            TabIndex        =   6
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   120
         Picture         =   "CIERREOP10.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CIERREOP10.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4845
         Width           =   1515
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "CIERREOP10.frx":2AC2
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CIERREOP10.frx":2CF6
      TabIndex        =   166
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubic:"
      Height          =   225
      Left            =   2280
      TabIndex        =   55
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant:"
      Height          =   225
      Left            =   2160
      TabIndex        =   54
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito Entrada:"
      Height          =   420
      Left            =   240
      TabIndex        =   53
      Top             =   1605
      Width           =   1065
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   255
      TabIndex        =   52
      Top             =   2160
      Width           =   1890
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro S/F:"
      Height          =   300
      Left            =   120
      TabIndex        =   49
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   48
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   630
      Width           =   1800
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   14
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   210
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   300
      Left            =   135
      TabIndex        =   13
      Top             =   1110
      Width           =   1125
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   12
      Top             =   1050
      Width           =   5370
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4755
      TabIndex        =   11
      Top             =   690
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   240
      Left            =   3150
      TabIndex        =   10
      Top             =   765
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
      Height          =   510
      Left            =   4200
      TabIndex        =   9
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5480
      TabIndex        =   8
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Prod. Nº:"
      Height          =   300
      Left            =   105
      TabIndex        =   7
      Top             =   270
      Width           =   1125
   End
End
Attribute VB_Name = "CIERREOP10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%
Dim NUELOTESS$(8)

Sub ENVIAMAILAPRO(OP$)
    MAILPROG$ = TRIM$(CONTROL$("BOLRECEP", "PROGMAIL"))
    ASUNTO$ = "Cierre de Orden de Produccion " & OP$
    TEXTOMAI$ = "Se ha Informado un Cierre de Produccion para " & OP$
    MAI_DESTI$ = TRIM$(MAILPROG$)
    MAI_DIREL$ = MAI_DESTI$
    MAI_ASUNT$ = ASUNTO$
    MAI_TEXTO$ = TEXTOMAI$
    If MAILPROG$ = "" Or InStr(MAILPROG$, "@") < 1 Then 'TIENE DIRECCION VALIDA
       Exit Sub
    End If
    '
    Call GENERAMAIL
End Sub

Sub GRABAMERMAS()
    '
    FF% = FECHANUM(Label24)
    DOPRI$ = "OP-" & Label32
    DOSEC$ = Label26
    '
    'Se graban las MERMAS
    For YX& = 0 To 8
       CIXI% = CODINT%(Label39(YX&))
       CANTIDD# = XVALO(text4(YX&))
       If CIXI% > 0 And CANTIDD# > 0 Then
          LOTE$ = Label40(YX&)
          CMO$ = "CF"    ' CONSUMO DE FABRICACION
          CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
30        DEP% = VALOBADA("lotingre", "DEPOUBIC", "idenlote='" & LOTE$ & "'")
          '
          'Movimiento de entrada
          REFE$ = "Comp. Mermas " + Label32
          Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, 0, DEP%, CANTIDD#)
          '
          'Movimiento de salida
          CMO$ = "MP"
          REFE$ = "Mermas " + Label32
          Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, 0, DEP%, CANTIDD# * (-1))
          '
         End If
     Next YX&
     '
End Sub

Sub GRACONSCANTASIG()
    '
    FF% = FECHANUM(Label24)
    DOPRI$ = "OP." & Label32
    DOSEC$ = "OP." & Label32
    REFE$ = "Cons. Mat. Asignado " + Label32
    NUORIT% = 0
    '
    XXPO% = 65
    '
    'Se graban las Consumo de (Cantidad Asignada - MERMAS)
    For YX& = 0 To List1.ListCount
       'SE TOMA LA CANTIDAD ASIGNADA
       LINEA$ = List1.List(YX&)
       FFX$ = (Mid(LINEA$, 1, 16))
       CIXI% = CODINT(FFX$)
       CODIGO$ = CODEXT$(CIXI%)
       DESC$ = DESCRIT0(CIXI%)
       UNI$ = UNIMED$(CIXI%)
       LOTE$ = TRIM$(Mid(LINEA$, 49, 15))
       CANTASI# = XVALO(Mid(LINEA$, 70, 10))
       IDREQUE$ = Mid(LINEA$, 90, 16)
       '
       'SE VERIFICA SI HAY MERMAS DE ESE CODIGO
       For XA% = 0 To 8
         If Label39(XA%) <> "" Then
            'en caso de haber mermas se descuentan a la cant asig
            TCI% = CODINT(Label39(XA%))
            If TCI% = CIXI% Then
               CANTASI# = CANTASI# - XVALO(text4)
            End If
         End If
       Next XA%
       '
       '
       CMO$ = "CF"    ' CONSUMO DE FABRICACION
       CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
       FEFABRI% = CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$, "NOMESS"))
       FEVENCI% = CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$, "NOMESS"))
       LOTEPROV$ = VALOBADA("LOTINGRE", "LoteCol_Prov", CONDI$, "NOMESS")
       PROVEE$ = VALOBADA("LOTINGRE", "ORIGEN", CONDI$, "NOMESS")
       DOCINGRE$ = VALOBADA("LOTINGRE", "DOCUINGRE", CONDI$, "NOMESS")
       CANTIDD# = CANTASI#
30     DEP% = XVALO(VALOBADA("lotingre", "DEPOUBIC", "idenlote='" & LOTE$ & "'"))
       DOSEC$ = NUMDOCSEC$(FF%)
       '
       Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, 0, DEP%, CANTIDD# * (-1))
       '
       Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
       '
     Next YX&
        
End Sub

Sub LimpiarPCP()
  '
  Call BLANARCH("!PCP")
  Call BLANARCH("PCP")
  '
End Sub

Private Sub Command1_Click()
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Private Sub Command10_Click(Index As Integer)
    DEPO% = XVALO(Text6)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Label41(Index) = UBISELE$
    End If
End Sub

Private Sub Command14_Click()
  Command14.Enabled = False
  Call CARTACONPRO
  Command14.Enabled = True
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   FIKARDEX09.Show 1
   Command15.Enabled = True
End Sub



Private Sub Command2_Click()
    DEPO% = XVALO(Text6)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text8 = UBISELE$
    End If
End Sub

Private Sub Command20_Click(Index As Integer)

   '
   If Label32 <> "" Then

        NORFA$ = TRIM$(Label32)
        CONDI$ = "NUORPROD = '" & NORFA$ & "'"
        SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
        CONDI$ = "DESTINO = '" & SOLFAB & "' order by cod_inte"
        Call COPYSTRU("MASTER", "MASREQ")
        Call CARGALISEL("!MASREQ", "BALAREQMO", "COD_INTE/C16+MASTER/A48", CONDI$, "DISTINCT")
        Call FRESHECHO("!MASREQ")
        Call SELECHO("!MASREQ")
        VDEVU$ = VALACT1$("!MASREQ")
'        Call SELECHO("MASTER")
'        VDEVU$ = VALACT1$("MASTER")
        If VDEVU$ <> "" Then
          CICOMPO% = CODINT%(VDEVU$)
          Label39(Index) = VDEVU$
          Label42(Index) = UNIMED(CICOMPO%)
      End If
   End If
   '

End Sub

Private Sub Command21_Click(Index As Integer)
 If Label32 <> "" Then
     CIXI% = CODINT%(Label39(Index))
     If CIXI% > 0 Then
        Call ARMALILOT(CIXI%, 0, 0, 1)
'        Call FRESHECHO("!LILOTSE")
'        Call SELECHO("!LILOTSE")
'        Call SELECHO("!LILOTSE")
'        VDEVU$ = VALACT1$("!LILOTSE")
'        VDEV2$ = VALACT2$("!LILOTSE")
        NORFA$ = TRIM$(Label32)
        CONDI$ = "NUORPROD = '" & NORFA$ & "'"
        SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
        CONDI$ = "COD_INTE = " & CIXI% & " AND COD_DESTI = '" & SOLFAB & "' AND IDENLOTE <> '' order by COD_DESTI"
        Call CARGALISEL("!LILOTES", "LOTINGRE", "IDENLOTE/A16;CANALTA/F10.1;CANSALDO/F10.1", CONDI$)
        Call FRESHECHO("!LILOTES")
        Call SELECHO("!LILOTES/Devolucion de Lotes")
        '
        VDEV1$ = VALACT1$("!LILOTES")
        VDEV2$ = VALACT2$("!LILOTES")
        For KKU& = 0 To 8
            If KKU& <> Index Then
              If Label40(KKU&) = VDEV1$ Then
                 Call MENSERR(24, "No Válido Repetir Lotes")
                 Exit Sub
              End If
            End If
        Next KKU&
        '
        If VDEV1$ <> "" Then
           Label40(Index) = VDEV1$
        End If
    End If
  End If
End Sub

Private Sub Command25_Click()
    '
    DEPO% = XVALO(Text6)
    If ECOARCH$("TADEPOSI", Chr$(DEPO%)) = "" Then
        Call MENSERR(24, "Primero Debe Seleccionar el Deposito")
        Exit Sub
    End If
    Call GENTAUBI(DEPO%)
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text2 = UBISELE$
    End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If EXISTE%("DECAOT09.EXE") > 0 Then
        RECONCI09.Show 1
      Else
        Call MENSERR(24, "Imposible Realizar Conciliación de Materiales.\  \La Presente Instalación no Incluye\Módulo de Depósito Auto-Organizado (Caotico).")
   End If
   Command3.Enabled = True
End Sub

Private Sub Command6_Click(Index As Integer)
   Label39(Index) = ""
   Label40(Index) = ""
   Label41(Index) = ""
   Label42(Index) = ""
   Text1(Index) = ""
   text4(Index) = ""
End Sub

Private Sub Label18_Change()
   TOTPROD# = XVALO(Label18) + XVALO(Label19)
   Label20 = FORMATNUM(TOTPROD#, "F13." & CNTDC$(CODINT(Label28)))
End Sub

Private Sub Label19_Change()
   TOTPROD# = XVALO(Label18) + XVALO(Label19)
   Label20 = FORMATNUM(TOTPROD#, "F13." & CNTDC$(CODINT(Label28)))
End Sub

Private Sub Label19_DblClick()
    'Se completa con los Rubros existentes
   Call BLANARCH("!PCP")
   For I& = 1 To ULTREG&("TACONPRO")
       X$ = REGLEIDO$("TACONPRO", I&)
       CODI% = CVI(Mid$(X$, 1, 2))
       DES$ = Mid(X$, 3, 30)
       If Asc(Left$(X$, 1)) > 0 Then
           'Se completa la cantidad ingresada en los rubros, en caso de haber
           'ingresado cantidades previamente
           PCON$ = REGLEIDO$("PCP", I&)
           If Asc(Left$(PCON$, 1)) > 0 Then
              CANTI# = CVD(Mid$(PCON$, 34, 8))
           End If
           '
           Call REPLA(TXT$, MKI$(CODI%), 1, 2)
           Call REPLA(TXT$, DES$, 3, 30)
           Call REPLA(TXT$, MKD$(CANTI#), 34, 8)
           Call GRAREG("!PCP", TXT$, I&)
      End If
   Next I&
   Call SETULTREG("!PCP", I&)
   Call CIERRARCH("!PCP")
   AA = ULTREG&("TACONPRO")
   '
   VTB% = VENTABU%("CONTPROC")
   '
   For I& = 1 To ULTREG&("!PCP")
      X$ = REGLEIDO$("!PCP", I&)
      CANTI# = CVD(Mid$(X$, 34, 8))
      SUMA& = SUMA& + CANTI#
      Call GRAREG("PCP", X$, I&)
   Next I&
   '
   CIXI% = CODINT(Label28)
   Call SETULTREG("PCP", I&)
   Call CIERRARCH("PCP")
   Label19 = FORMATNUM(SUMA&, "F11." + CNTDC$(CIXI%))
   '
End Sub

Private Sub Label28_Change()
   CIXI% = CODINT%(Label28)
   'Text6 = TRIM$(Str(DEPOPRE%(CIXI%)))  'EL DEPOSITO LO DEBE SELECCIONAR EL USUARIO
End Sub

Private Sub Label39_Click(Index As Integer)
  Call MUESTRASEL1(Index)
End Sub

Private Sub Label40_Click(Index As Integer)
 Call MUESTRASEL1(Index)
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("TAUBICA")
    
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text2 = UBISELE$
    End If
End Sub



Private Sub Text3_LostFocus()
   Text3 = FORMATNUM(XVALO(Text3), "F13." & CNTDC$(CODINT(Label28)))
End Sub

Private Sub Text4_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 8
       If Label39(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(text4(KKI%))
         Else
           text4(KKI%) = ""
       End If
    Next KKI%
    Label50 = TRIM$(FORMATNUM$(SUMATRA, "F13." + CNTDC$))
'    If XVALO(Text4(Index)) < 0 Or TRIM$(Text4(Index)) = "" Then
'
'       Label39(Index) = ""
'
'       Label42(Index) = ""
'
'    End If
End Sub

Private Sub Text4_LostFocus(Index As Integer)
   DECI$ = CNTDC$(CODINT(Label39(Index)))
   text4(Index) = FORMATNUM(XVALO(text4(Index)), "F13." + DECI$)
End Sub

Private Sub Text5_Change()
   Label18 = XVALO(Text5) + XVALO(Text7)
   Label18 = FORMATNUM$(Label18, "F13." + CNTDC$(CODINT(Label28)))
End Sub

Private Sub Text5_LostFocus()
 Text5 = FORMATNUM(XVALO(Text5), "F13." + CNTDC$(CODINT(Label28)))
End Sub

Private Sub Text6_Change()
   If XVALO(Text6) > 0 Then
      Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
      Text2 = ""
     Else
      Label38 = ""
   End If
End Sub

Private Sub Text6_DblClick()
   Call Command30_Click
End Sub


Private Sub Command13_Click()
    '
    Command13.Enabled = False
    '
    If Label32 = "" Then
       Call MENSERR(24, "Número de O/F no Válido")
       GoTo 99
    End If
    '
    ' VALIDAR FECHA
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% < 1 Or FECHA% > HOII% Then
       Call MENSERR(24, "Fecha de Lanzamiento No Válida")
       GoTo 99
    End If
    '
    ' valida deposito de entrada
    If XVALO(Text6) < 1 Then
       Call MENSERR(24, "Falta Indicar Deposito de Entrada.")
       GoTo 99
    End If
    '
    DEPENTRA% = XVALO(Text6)
    If DEPENTRA% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPENTRA%)) = "" Then
       Call MENSERR(24, "Deposito de Entrada no Válido")
       GoTo 99
    End If
    '
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    If DEPENTRA% = DEPOTRANS% Then
       Call COMUNI("Depósito de Entrada No Válido\Configurado como Material en Transito ")
       GoTo 99
    End If

    'Validar Ubicación Cant Conforme
    If XVALO(Text5) > 0 Then
      POSI$ = TRIM$(Text2)
      CI% = CODINT(Label28)
      If PERMITEUBI%(CI%, "", POSI$, DEPENTRA%) <> 1 Then
         GoTo 99
      End If
    End If
    '
    'Validar Ubicación Cant NO Conforme
    If XVALO(Text7) > 0 Then
      POSI$ = TRIM$(Text8)
      If PERMITEUBI%(CI%, "", POSI$, DEPENTRA%) <> 1 Then
         GoTo 99
      End If
    End If
    '
    If EXISTE%("DECAOT09.EXE") < 1 And INTERPRE% < 1 Then
       If TRIM$(Text3) = "" Then
          Text3 = Label18
          GoTo 50
       End If
    End If
    '
    If XVALO(Label18) > 0 Then
      If XVALO(Text3) <= 0 Then
         Call MENSERR(24, "Falta Ingresar Unidades por Bulto")
         GoTo 99
      End If
    End If
    'No deja continuar si la cantidad de bultos es > a 99
    If XVALO(Label18) > 0 Then
        REALIZADO# = XVALO(Label18): UNXBIL# = XVALO(Text3)
        If (REALIZADO# / UNXBIL#) > 99 Then
            Call MENSERR(24, "Cantidad de Bultos Mayor a 99 No Admisible.\ Vuelva a Informar el Cierre de la O/P")
            GoTo 99
        End If
    End If
    '
    CONDI$ = "DOPRICOR = 'OP-" & Label32 & "' AND CODIMOVI = 'CF'"
    If CantRegistros("MOVISTO", CONDI$) < 1 Then
       Call COMUNI("ADVERTENCIA: No Hay Consumos Declarados\para Esta Orden de Produccion")
    End If
    '
    CONDI$ = "DESTINO ='" & Label26 & "' AND KILO_ASIG > 0 AND KILO_CONSU < KILO_ASIG "
    If CantRegistros("BALAREQMO", CONDI$) > 0 Then
       RESP% = COMALTER%("Falta Declarar Fraccionamiento / Consumo\de Material Asignado para Esta Orden de Produccion\\Cancelar Cierre\Continuar")
       If RESP% <> 2 Then GoTo 99
    End If
    '
    'Validacion de Items Sobrantes
50  For U& = 0 To Text1.Count - 1
       CODIGO$ = Label39(U&)
       If CODIGO$ <> "" Then
         CIXI% = CODINT%(CODIGO$)
         If CIXI% < 1 Then
            Call MENSERR(24, "Codigo no Válido en Linea " & U + 1 & " de Devoluciones")
            GoTo 99
         End If
         '
         If TRIM$(Label40(U&)) = "" Then
            Call MENSERR(24, "Componente sin Lote / Partida en Linea " & U + 1 & " de Devoluciones")
            GoTo 99
         End If
         '
         CANDEVU = XVALO(Text1(U&))
         If CANDEVU > 0 Then
           IDLOT$ = TRIM$(Label40(U&))
           CONDI$ = "COD_INTE = " & CIXI% & " AND idenlote='" & IDLOT$ & "'"
           CanAlta = VALOBADA("LOTINGRE", "CANALTA", CONDI$)
           If CanAlta < CANDEVU Then
              COMUNI ("Cantidad Devuelta Supera a Cantidad Original\del en Lote, en Fila Nro " & U& + 1)
              GoTo 99
           End If
           POSI$ = TRIM$(Label41(U&))
           ' If PERMITEUBI%(CIXI%, "", POSI$, DEPENTRA%) <> 1 Then
           DESTINX$ = TRIM$(Label26)   ' LA DEVOLUCION QUEDA DESTINADA A LA SF - EPB 6/10/11
           If PERMITEUBI%(CIXI%, DESTINX$, POSI$, DEPENTRA%) <> 1 Then
              GoTo 99
           End If
         End If
       End If
    Next U&
    '
    'Si se ingresan los datos correctos se graba el aprobado y no conforme
    '
    Call GRACIERROF(STRILOT$) 'DEVUELVE UN STRING CON LOS LOTES GENERADOS POR BULTO
    Call GRABADEVOL
    Call GRABAMERMAS
    '
    If EXISTE%("DECAOT09.EXE") < 1 And INTERPRE% < 1 Then GoTo 70
    '
    'Calculo de ETIQUETAS CONFORMES
    CICONJ% = CODINT%(Label28)
    For I& = 1 To Len(STRILOT$) Step 32
       LOTEID$ = TRIM$(Mid$(STRILOT$, I&, 32))
       Call PRINSTIKAPRO1(CICONJ%, LOTEID$)
    Next I&
    '
    'DEVOLUCIONES (*** QUEDA PARA VER POR EPB ***)
60  For I& = 0 To 8
      If CODINT%(Label39(I&)) > 0 Then
       If XVALO(Text1(I&)) > 0 Then
          CIXI% = CODINT%(Label39(I&))
          CANTIDEVU = XVALO(Text1(I&))
          DEP% = XVALO(Text6)
          LOTEESS$ = Label15(I&)
          Call PRINSTIKAPRO1(CIXI%, LOTEESS$, DEP%, CANTIDEVU)     ' IMPRIME ETIQUETAS DE DEVOLUCIONES
        End If
       End If
    Next I&
    '
70  SOLFA$ = Label26
    Call ABREORFAB
    CONDI$ = "IDSUBOR = '" & SOLFA$ & "'"
    CANPROYEC = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$))
    CANREALIZ = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$))
    CIERRA% = 0
    If CANREALIZ >= CANPROYEC Then
         CIERRA% = 1
      Else
         OPX% = COMALTER%("ATENCION: La OP." & Label32 & " \queda con Saldo Pendiente.\\Mantener Abierta\Cerrar")
         If OPX% = 2 Then CIERRA% = 1
    End If
    If CIERRA% = 1 Then
       Call ACTUREGISTRO("ORDEFABR", "STATUORF", CONDI$, 3, RAFE&)
       Call ACTUREGISTRO("ORDEFABR", "FechCierr", CONDI$, FETEXCOR1$(FECHANUM(Label24)), RAFE&)
    End If
    '
    If EXISTE%("DECAOT09.EXE") < 1 And INTERPRE% < 1 Then GoTo 98

    ' Envio de mail a programacion.
    Call ENVIAMAILAPRO(Label32)
    '
98  Screen.MousePointer = 1
    Call Label32_DblClick 'Se usa para limpiar campos
    Call LimpiarPCP       'Limpia archivo
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
End Sub

Function NUMDOCSEC$(FF%)
    '
    FEXX$ = FECHATEX$(FF%)
    '
    IDENORPROD$ = Mid$(FEXX$, 7, 2) & Mid$(FEXX$, 4, 2) & Mid$(FEXX$, 1, 2)
    '
    
    VALOCAMP$ = "MAX (CAST(SUBSTRING(DOSECCOR, 8,6)AS INTEGER ))"
    CONDI$ = "doseccor LIKE '" & IDENORPROD$ & "%'"
    '
    PRONUMORPRO% = 1 + XVALO(VALOBADA("Movisto", VALOCAMP$, CONDI$, "NOMESS"))
    '
    NUMORDPROD$ = TRIM$(Str$(PRONUMORPRO%))
    While Len(NUMORDPROD$) < 4: NUMORDPROD$ = "0" + NUMORDPROD$: Wend
    NUMDOCSEC$ = IDENORPROD$ & "/" & NUMORDPROD$
    '
End Function
'
Sub GRACIERROF(STRILOT$)
     STRILOT$ = ""
     '1.- Actualiza saldo pendiente de la Orden de Fabricacion
     SOLFA$ = Label26
     IDLOBA$ = Label32
     DEPX% = XVALO(Text6)
     If DEPX% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
     '
     CANTI = XVALO(Text5) + XVALO(Label19) ' APROBADO + PCP
     SCRAP = XVALO(Text7)
     CIERRA% = 0: If XVALO(Label18) + XVALO(Label37) >= XVALO(Label34) Then CIERRA% = 1
     If XVALO(Label18) + XVALO(Label37) < XVALO(Label34) Then
          OPX% = COMALTER%("ATENCION: La OP." & NORFA$ & " \queda con Saldo Pendiente.\\Mantener Abierta\Cerrar")
          If OPX% = 2 Then CIERRA% = 1
     End If
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM  ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & SOLFA$ & "'"
     Dim ABREORFFFAB As ADODB.Recordset
     Set ABREORFFFAB = New ADODB.Recordset
     ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     On Error Resume Next
     With ABREORFFFAB
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           REFEOF$ = !REFEREN
          !CanApro = !CanApro + CANTI
          !CanRech = !CanRech + SCRAP
          If CIERRA% = 1 Then
            !STATUORF = 3         ' MARCA COMO CERRADA
          End If
          !FechCierr = CVDAT(FECHANUM(Label24))
          .Update
        End If
        On Error GoTo 0
     End With
     ABREORFFFAB.Close
     Set ABREORFFFAB = Nothing
     '
     ' HASTA QUI CERRO LA ORDEN DE FABRICACION
     '
     If CANTI >= 0.05 Then
        NUORIT% = 0
        FECHA% = FECHANUM(Label24)
        CICONJ% = CODINT%(Label28)
        CMOV$ = "OP"
        DOPRI$ = "OP-" & Label32
        REFE$ = REFEOF$
        VUNI# = 0
        MONEII% = 1
        NC% = 0:
        UBIK$ = Left$(Text2, 32)
        '
        CONDI$ = "IDSUBOR = '" & SOLFA$ & "'"
        LOTEPRO$ = SAFETEXT(VALOBADA("ORDEFABR", "IDENLOTE", CONDI$, "NOMESS"))
        If TRIM$(LOTEPRO$) = "" Then LOTEPRO$ = Label32
        '
        'GUARDA EL LOTE POR CADA BULTO
        SALDO = XVALO(Text5) + XVALO(Label19)    ' AGREGA LOS DE CONTROL DE PROCESO
        BULTINI% = 0
        Do While SALDO > 0
           CANTI = XVALO(Text3)
           If CANTI > SALDO Then CANTI = SALDO
           SALDO = SALDO - CANTI
           '
15         BULTINI% = BULTINI% + 1
           IBUL$ = TRIM$(Str$(BULTINI%)): While Len(IBUL$) < 2: IBUL$ = "0" + IBUL$: Wend
           LOTEID$ = IDLOBA$ + "-" + IBUL$
           If CantRegistros("LOTINGRE", "IDENLOTE = '" & LOTEID$ & "'") > 0 Then GoTo 15
           '
           FEVEN% = DIASPOST(FECHA, 730)
           Call CREALOTIN(FECHA%, 0, 0, CICONJ%, LOTEID$, "", LOTEPRO$, CDbl(CANTI), 0, 0, DEPX%, UBIK$, "", "", "", "CUAR")
           STRILOT$ = STRILOT$ + AJUSTI$(LOTEID$, 32)
           DOSEC$ = NUMDOCSEC$(FECHA%)
           Call MOVISTOK(FECHA%, CICONJ%, LOTEID$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI)
        Loop
     End If
     '
     STRILOCONF$ = STRILOT$    ' CONTIENE UNICAMENTE LOS LOTES CONFORMES
     '
     If CICONJ% > 0 Then
        SQLX$ = "UPDATE  MASTER "
        SQLX$ = SQLX$ + " SET DepoEntra = " & DEPX%
        SQLX$ = SQLX$ + " WHERE CODINT = " & CICONJ% & ""
        FLEXCONN.Execute (FILTSQL$(SQLX$))
     End If
     '
     'Registro de lote para cantidades no conformes
     CANTNC = XVALO(Text7)
     If CANTNC > 0 Then
        LOTENC$ = Label32 + "/NC"
        UBIKNC$ = TRIM$(Text8)
        DOSEC$ = NUMDOCSEC$(FECHA%)
        BULTINI% = 0
        '
35      BULTINI% = BULTINI% + 1
        IBUL$ = TRIM$(Str$(BULTINI%)): While Len(IBUL$) < 2: IBUL$ = "0" + IBUL$: Wend
        LOTEID$ = LOTENC$ + "-" + IBUL$
        If CantRegistros("LOTINGRE", "IDENLOTE = '" & LOTEID$ & "'") > 0 Then GoTo 35
        '
        Call CREALOTIN(FECHA%, FECHA%, FECHA%, CICONJ%, LOTEID$, "", LOTEPRO$, CDbl(CANTNC), 0, 0, DEPX%, UBIKNC$, "", "", "", "RECH")
        STRILOT$ = STRILOT$ + AJUSTI$(LOTEID$, 32)
        Call MOVISTOK(FECHA%, CICONJ%, LOTEID$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFEOF$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTNC)
     End If
     '
     'ESTA PENDIENTE PARA VER ESTO COMO SE VA HACER   (**LO TIENE QUE VER EPB**)
     'Lee PCP y Graba movimiento de salida en Movisto
     For XA& = 1 To ULTREG("!PCP")
       REGISTRO$ = REGLEIDO$("!PCP", XA&)
       CODCON& = CVI(Mid(REGISTRO$, 1, 2))
       DES$ = Mid(REGISTRO$, 3, 30)
       CANX# = CVD(Mid(REGISTRO$, 34, 8))
       If CANX# > 0 Then
          REFE$ = "(" & (CODCON&) & ") " & DES$
          ' BUSCA DE QUE LOTE LO VA A DESCONTAR
          CANLOTES% = Len(STRILOCONF$) / 32
          For KU% = CANLOTES% To 1 Step -1
             ILOT$ = TRIM$(Mid$(STRILOCONF$, 32 * CANLOTES% - 31, 32))
             If SaldoLote(CICONJ%, ILOT$) >= CANX# Then
               'definir documento secundario
               DOSEC$ = NUMDOCSEC$(FECHA%)
               Call MOVISTOK(FECHA%, CICONJ%, ILOT$, "CP", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANX# * (-1))
               Call ACSALOTE(CICONJ%, ILOT$, "NOMESS")
               GoTo 39
             End If
          Next KU%
          '
          Call MENSERR(24, "Imposible Registrar Movimiento PCP\Correspondiente a: " + REFE$)
          '
       End If
39   Next XA&
     '
     SOFA$ = TRIM$(Label26)
     CANLOTES% = Len(STRILOCONF$) / 32
     For KU% = CANLOTES% To 1 Step -1
        ILOT$ = TRIM$(Mid$(STRILOCONF$, 32 * CANLOTES% - 31, 32))
        Call REEMPLAZALOTE(CICONJ%, SOFA$, ILOT$, DEPX%)
     Next KU%
     '
End Sub
'
Sub GRABADEVOL()
     '
     FF% = FECHANUM(Label24)
     DOPRI$ = "OP-" & Label32
     DOSEC$ = Label26
     REFE$ = "Devolucion " + Label32
     '
     'Graba Devoluciones
     For YX& = 0 To 8
        CIXI% = CODINT%(Label39(YX&))
        CANTIDD# = XVALO(Text1(YX&))
        If CIXI% > 0 And CANTIDD# > 0 Then
           UBIC$ = TRIM$(Label41(YX&))
           CMO$ = "DV"    ' CONSUMO DE FABRICACION
30         DEP% = XVALO(Text6)    ' VALOBADA("lotingre", "DEPOUBIC", "idenlote='" & LOTE$ & "'")
           LOTORI$ = Label40(YX&)
           LOTE$ = PROXILOTE$(LOTORI$)
           '
           CONDI$ = "IDENLOTE = '" & LOTORI$ & "'"
           CCIIN% = XVALO(VALOBADA("LOTINGRE", "COD_INTE", CONDI$, "NOMESS"))
           REMIPROX$ = SAFETEXT$(VALOBADA("LOTINGRE", "DocuIngre", CONDI$, "NOMESS"))
           FECHALTT% = CVINT(VALOBADA("LOTINGRE", "FECHALTA", CONDI$, "NOMESS"))
           ENVASX$ = SAFETEXT$(VALOBADA("LOTINGRE", "CodEnvase", CONDI$, "NOMESS"))
           FECHFABR% = CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$, "NOMESS"))
           FECHAVEN% = CVINT(VALOBADA("LOTINGRE", "FECHVENC", CONDI$, "NOMESS"))
           MARCAX$ = SAFETEXT$(VALOBADA("LOTINGRE", "Marca", CONDI$, "NOMESS"))
           ORIGENIN$ = SAFETEXT$(VALOBADA("LOTINGRE", "Origen", CONDI$, "NOMESS"))
           LOTEINGR$ = SAFETEXT$(VALOBADA("LOTINGRE", "LOTECOL_PROV", CONDI$, "NOMESS"))
           STALIB$ = SAFETEXT$(VALOBADA("LOTINGRE", "Libera", CONDI$, "NOMESS"))
           DESTIN$ = ""             ' EL FRACCIONADO QUEDA SIN DESTINO - epb 2/4/2010    ' SAFETEXT$(!Cod_Desti)
           DESTIN$ = TRIM$(Label26) ' LA DEVOLUCION QUEDA ASIGNADA A LA SF - EPB 6/10/11
           '
           Call CREALOTIN(FECHALTT%, FECHFABR%, FECHAVEN%, CIXI%, LOTE$, _
             MARCAX$, LOTEINGR$, CANTIDD#, 0, 0, DEP%, UBIC$, ORIGENIN$, _
             REMIPROX$, ENVASX$, STALIB$, 0, DESTIN$)
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, 0, DEP%, CANTIDD#)
           '
           Label15(YX&) = LOTE$
           '
         End If
     Next YX&
     '
End Sub
'
Sub PRIVALEMA()
     '
     FORIPRE$ = TRIM$(CONTROL$("", "FORCONSU"))
     If FORIPRE$ = "" Then
        Call MENSERR(24, "Falta Definir Formulario Vale de Materiales")
        Exit Sub
     End If
     '
     If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Consumo de Materiales:\  \'" + RMCC$ + "'.")
           Exit Sub
        End If
     End If
     '
     Screen.MousePointer = 11
     FECHDOC$ = Label24
     NUMEDOC$ = Label51
     TIPODOC$ = "Vale de Consumo de Materiales"
     '
     'CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
     'CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
     CAPARDOC% = 2
     '
     CODTABU1$(1) = "ORD-ITEM"
     CODTABU1$(2) = "CANTIDAD"
     CODTABU1$(3) = "COD-MAT"
     CODTABU1$(4) = "DES-MAT"
     CODTABU1$(5) = "NRO-LOTE"
     CODTABU1$(6) = "REF-MAT1"
     CACOLTAB1% = 6
     '
     DESTIDOC% = 0
     CAITAB1% = 0
     '
     For YX& = 0 To 2
        If NLOTX$ <> "" Then
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(4, CAITAB1%) = Label6
        End If
     Next YX&
     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(2, CAITAB1%) = Text1(YX&)
           TETABU1$(3, CAITAB1%) = CODEX$
           TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
           TETABU1$(5, CAITAB1%) = Label40(YX&)
           TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label40(YX&))
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(2, CAITAB1%) = Text1(YX&)
           TETABU1$(3, CAITAB1%) = CODEX$
           TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
           TETABU1$(5, CAITAB1%) = ""
           TETABU1$(6, CAITAB1%) = UBISTO(CIXI%)
        End If
     Next YX&
     '
     Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
     Screen.MousePointer = 1
     '
     Call CIERRARCH("*.*")
     '
End Sub






Private Sub Command18_Click(Index As Integer)
   '
   If Label32 <> "" Then

        NORFA$ = TRIM$(Label32)
        CONDI$ = "NUORPROD = '" & NORFA$ & "'"
        SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
        CONDI$ = "DESTINO = '" & SOLFAB & "' order by cod_inte"
        Call COPYSTRU("MASTER", "MASREQ")
        Call CARGALISEL("!MASREQ", "BALAREQMO", "COD_INTE/C16+MASTER/A48", CONDI$, "DISTINCT")
        Call FRESHECHO("!MASREQ")
        Call SELECHO("!MASREQ")
        VDEVU$ = VALACT1$("!MASREQ")
'        Call SELECHO("MASTER")
'        VDEVU$ = VALACT1$("MASTER")
        If VDEVU$ <> "" Then
          CICOMPO% = CODINT%(VDEVU$)
          Label39(Index) = VDEVU$
          Label42(Index) = UNIMED(CICOMPO%)
      End If
   End If
   '
End Sub

Private Sub Command19_Click(Index As Integer)
  If Label32 <> "" Then
     CIXI% = CODINT%(Label39(Index))
     If CIXI% > 0 Then
        Call ARMALILOT(CIXI%, 0, 0, 1)
'        Call FRESHECHO("!LILOTSE")
'        Call SELECHO("!LILOTSE")
'        Call SELECHO("!LILOTSE")
'        VDEVU$ = VALACT1$("!LILOTSE")
'        VDEV2$ = VALACT2$("!LILOTSE")
        NORFA$ = TRIM$(Label32)
        CONDI$ = "NUORPROD = '" & NORFA$ & "'"
        SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
        CONDI$ = "BALAREQMO.COD_INTE = " & CIXI% & " AND BALAREQMO.DESTINO = '" & SOLFAB$ & "' AND BALAREQMO.KILO_CONSU > 0"
        Call CARGALISEL("!LILOTES", "BALAREQMO+LOTINGRE/BALAREQMO.IDENLOTE = LOTINGRE.IDENLOTE", "BALAREQMO.IDENLOTE/A16;LOTINGRE.CANALTA/F10.1;LOTINGRE.CANSALDO/F10.1", CONDI$)
        Call FRESHECHO("!LILOTES")
        Call SELECHO("!LILOTES/Devolucion de Lotes")
        '
        VDEV1$ = VALACT1$("!LILOTES")
        If VDEV1$ <> "" Then
           VDEV2$ = VALACT2$("!LILOTES")
           For KKU& = 0 To 8
               If KKU& <> Index Then
                 If Label40(KKU&) = VDEV1$ Then
                    Call MENSERR(24, "No Válido Repetir Lotes")
                    Exit Sub
                 End If
               End If
           Next KKU&
           '
           Label40(Index) = VDEV1$
        End If
    End If
  End If

End Sub

Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Text6 = VALACT1$("TADEPOSI")
End Sub



Private Sub Command4_Click()
    Call CommSel("SELFECHA", Command4, Label24)
End Sub

Private Sub Command5_Click()
    '
    CONDI$ = "StatuOrf < 3 "   ' para que no este anulada ni cumplida
    'CONDI$ = CONDI$ + "AND CanApro < (CanProy - 0.005) "
    CONDI$ = CONDI$ + " and NuOrProd <> '' "
    CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    
    Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
    Call CommSel("ORDPEN/Indice de Ordenes de Produccion Pendientes", Command5, Label32)
    
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    '
End Sub

Sub TAMARCAS()
    '
    If CI% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      CI% = CODINT(COD1$)
20    If CI% > 0 Then
         Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(CI%))
         VTB% = VENTABU%("TAMARCAS")
         If VTB% >= 0 Then
            Call FILTERUPDATE("TAMARCAS", "!TAMARCAS", 4, MKI$(CI%))
         End If
      End If
    End If
    '
End Sub
'
'Private Sub Command6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
'     Call Command6_Click(Index)
'   End If
'End Sub




Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !Cod_Inte
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function



Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       'Frame9.Visible = True
    End If
    '
    Call FRESHALL
    '
    If Left$(APLINVO$, 7) = "OFABRIC" Then
       Label32 = TRIM$(Mid$(APLINVO$, 8))
       If Label32 <> "" Then
          If ECOARCH$("ECORDEP", Label32) <> "" Then
            HOII% = HOY(HOS$)
            Label24 = HOS$
          End If
       End If
    End If
    '
    Call LimpiarPCP
    '
    Picture2.Top = 0
    VScroll3.Min = 0
    VScroll3.Max = (Picture2.Height - Picture1.Height) / 10
    '

End Sub


Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
       NVM$ = TRIM$(ProNroComprobante("CF"))
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    Label51 = "CF-" + NVM$
    Label20 = ""
    HOII% = HOY(HOS$)
    Label24 = HOS$
    ' PONER EN B
    'LANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Label28 = ""
    Label30 = ""
    Text5 = ""
    Text6 = ""
    Text2 = ""
    Text3 = ""
    Label34 = ""
    Label37 = ""
    Label18 = ""
    Label26 = ""
    Label19 = ""
    Text7 = ""
    Text8 = ""
    'Limpia Devoluciones
    For I& = 0 To 8
      Label39(I&) = ""
      Label40(I&) = ""
      Label41(I&) = ""
      Label42(I&) = ""
      Text1(I&) = ""
    Next I&
    'Limpia Mermas
    For I& = 0 To 8
      text4(I&) = ""
    Next I&
    '
    List1.Clear
    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         'SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
         SQLX$ = SQLX$ + " WHERE NuOrProd = '" & NORFA$ & "'"
         '
         Dim ORDEFABTMP As ADODB.Recordset
         Set ORDEFABTMP = New ADODB.Recordset
         Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With ORDEFABTMP
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "Orden de Fabricación no Válida. Consulte.")
              Call FINAL("")
           End If
           On Error GoTo 0
           CIXI% = !Cod_Inte
           Label28 = CODEXT$(CIXI%)
           Label30 = DESCRIT0$(CIXI%)
              CAPENDE# = !CanProy - !CanApro: If CAPENDE < 0 Then CAPENDE = 0
              If !STATUORF >= 3 Then CAPENDE# = 0
           Label34 = FORMATNUM$(!CanProy, "F13." + CNTDC$(CIXI%))
           Label37 = FORMATNUM$(!CanApro, "F13." + CNTDC$(CIXI%))
           'Label18 = FORMATNUM$(CAPENDE#, "F10." + CNTDC$)
           Label26 = SAFETEXT(!IdSubOr)
           '
           List1.Clear
           Call CargarMatAsignados(Label26)
         End With
         '
    End If
99  Screen.MousePointer = 1
End Sub

Private Sub Label32_DblClick()
    ' USAR PARA LIMPIAR TODO
    Label32 = ""
End Sub

'
Sub LIMPLAB()

   'Label6.Caption = ""
   'text4.TEXT = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text12 = ""
  
End Sub


Private Sub Label39_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label40_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub
'
Private Sub Text1_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 8
       If Label39(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text1(KKI%))
         Else
           Text1(KKI%) = ""
       End If
    Next KKI%
    DECI$ = CNTDC$(CODINT(Label39(Index)))
    Label49 = TRIM$(FORMATNUM$(SUMATRA, "F13." + DECI$))
'    If XVALO(Text1(Index)) < 0 Or TRIM$(Text1(Index)) = "" Then
'       Label39(Index) = ""
'       Label42(Index) = ""
'       Label40(Index) = ""
'       Label41(Index) = ""
'    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    DECI$ = CNTDC$(CODINT(Label39(Index)))
    Text1(Index) = FORMATNUM(XVALO(Text1(Index)), "F13." + DECI$)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text2.SetFocus
    End If

End Sub

Private Sub Text7_Change()
   Label18 = XVALO(Text5) + XVALO(Text7)
   Label18 = FORMATNUM$(Label18, "F13." + CNTDC$(CODINT(Label28)))
End Sub

Private Sub Text7_LostFocus()
   Text7 = FORMATNUM(XVALO(Text7), "F13." + CNTDC$(CODINT(Label28)))
End Sub

Private Sub Text8_Change()
   DEPO% = XVALO(Text6)
   Call GENTAUBI(DEPO%)
   Label17 = ECOARCH$("!TAUBICA", Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
End Sub

Sub MUESTRASEL(Index)
   '
   CANSA = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
   LARGX = 0: If ANCHX > 0 Then LARGX = CANSA / ANCHX
    '
End Sub

Sub MUESTRASEL1(Index)
   '
   If Index < 4 Then
      CIXI% = CODINT%(TRIM$(Label39(Index)))
      IDLOT$ = TRIM$(Label40(Index))
      NUMDEP% = XVALO(Text6)
   End If
   '
End Sub

Sub CIERROFAB()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ECORDEF", "ECORDEP")
    J& = 0
    '
    Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT NuOrProd, StatuOrf, CanApro, CanProy, REFEREN FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
    SQLX$ = SQLX$ + " and NuOrProd <> '' "
    SQLX$ = SQLX$ + " ORDER BY NuOrProd "
    '
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "NO Existen Ordenes de Fabricación\Correspondientes a Criterio de Selección.\  \Imposible Continuar la Carga.")
        GoTo 90
      End If
      On Error GoTo 0
      JJ& = 0
      Do While Not .EOF
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, SAFETEXT(!NUORPROD), 1, 12)
        Call REPLA(Y$, SAFETEXT(!REFEREN), 13, 64)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Set ORDEFABTMP = Nothing
    
    '
90  Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
    Call FRESHECHO("ECORDEP")
    Screen.MousePointer = 1
    '
End Sub

Private Sub VScroll2_Change()
   If VScroll2.Value < 360 Then
    Picture5.Top = (-10) * VScroll2.Value
   End If
End Sub

Private Sub VScroll2_Scroll()
   If VScroll2.Value < 360 Then
    Picture5.Top = (-10) * VScroll2.Value
   End If
End Sub

Private Sub VScroll3_Change()
   If VScroll3.Value < 360 Then
    Picture2.Top = (-10) * VScroll3.Value
   End If
End Sub

Private Sub VScroll3_Scroll()
   If VScroll3.Value < 360 Then
    Picture2.Top = (-10) * VScroll3.Value
   End If
End Sub
'
Sub CargarMatAsignados(DESTINO$)
    Dim Reque As ADODB.Recordset
    Set Reque = New ADODB.Recordset
    List1.Clear
    SQLX$ = "SELECT COD_INTE, IDENLOTE, IDENMOVI, SUM(CANTINGRE) AS CAIN, SUM(CANTSALID) AS CANCON FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE DoSecCor = '" & DESTINO$ & "' "
    SQLX$ = SQLX$ + " AND CODIMOVI = 'CF' "
    SQLX$ = SQLX$ + "GROUP BY COD_INTE, IDENLOTE, IDENMOVI "
    Set Reque = FLEXCONN.Execute(SQLX$)
    With Reque
      Do While Not (.EOF)
       LINEA$ = ""
       CI% = XVALO(!Cod_Inte)
       CODIGO$ = CODEXT$(CI%)
       DESC$ = DESCRIT0(CI%)
       UNI$ = UNIMED$(CI%)
       LOTE$ = SAFETEXT(!idenLote)
       CANTASI# = XVALO(!CANCON) - XVALO(!CAIN)
       IDREQUE$ = SAFETEXT$(!IdenMovi)
       '
       Call REPLA(LINEA$, CODIGO$, 1, 16)
       Call REPLA(LINEA$, DESC$, 17, 30)
       Call REPLA(LINEA$, LOTE$, 49, 16)
       Call REPLA(LINEA$, UNI$, 64, 5)
       cx$ = FORMATSTAN$(CANTASI#, "F13", CI%)  'FORMATNUM$(CANTASI#, "F13." & CNTDC$)
       Call REPLA(LINEA$, cx$, 70, 13)
       Call REPLA(LINEA$, IDREQUE$, 90, 16)
       List1.AddItem LINEA$
       .MoveNext
      Loop
    End With
End Sub
