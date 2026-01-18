VERSION 5.00
Begin VB.Form TRAZORF07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Materiales a O/F"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Materiales No Trazables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   210
      TabIndex        =   80
      Top             =   6360
      Width           =   8730
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
         Index           =   7
         Left            =   6405
         TabIndex        =   123
         Top             =   1470
         Width           =   175
      End
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
         Index           =   6
         Left            =   6405
         TabIndex        =   121
         Top             =   1155
         Width           =   175
      End
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
         Index           =   5
         Left            =   6405
         TabIndex        =   119
         Top             =   840
         Width           =   175
      End
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
         Index           =   4
         Left            =   6405
         TabIndex        =   116
         Top             =   525
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
         Index           =   4
         Left            =   2205
         TabIndex        =   88
         Top             =   510
         Width           =   175
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
         Index           =   4
         Left            =   7455
         TabIndex        =   87
         Top             =   510
         Width           =   1065
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
         Left            =   2205
         TabIndex        =   86
         Top             =   825
         Width           =   175
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
         Index           =   5
         Left            =   7455
         TabIndex        =   85
         Top             =   825
         Width           =   1065
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
         Left            =   2205
         TabIndex        =   84
         Top             =   1140
         Width           =   175
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
         Index           =   6
         Left            =   7455
         TabIndex        =   83
         Top             =   1140
         Width           =   1065
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
         Left            =   2205
         TabIndex        =   82
         Top             =   1455
         Width           =   175
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
         Index           =   7
         Left            =   7455
         TabIndex        =   81
         Top             =   1455
         Width           =   1065
      End
      Begin VB.Label Label52 
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
         Index           =   7
         Left            =   6090
         TabIndex        =   124
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1455
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   6
         Left            =   6090
         TabIndex        =   122
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   5
         Left            =   6090
         TabIndex        =   120
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   825
         Width           =   330
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4305
         TabIndex        =   118
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label52 
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
         Index           =   4
         Left            =   6090
         TabIndex        =   117
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   510
         Width           =   330
      End
      Begin VB.Label Label21 
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
         Index           =   4
         Left            =   2625
         TabIndex        =   89
         Top             =   510
         Width           =   3300
      End
      Begin VB.Label Label39 
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
         Index           =   4
         Left            =   525
         TabIndex        =   104
         ToolTipText     =   "Código"
         Top             =   510
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   4
         Left            =   6720
         TabIndex        =   103
         Top             =   510
         Width           =   420
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   6930
         TabIndex        =   102
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -1155
         TabIndex        =   101
         Top             =   285
         Width           =   1965
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
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
         Left            =   6300
         TabIndex        =   100
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label39 
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
         Index           =   5
         Left            =   525
         TabIndex        =   99
         ToolTipText     =   "Código"
         Top             =   825
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   5
         Left            =   6720
         TabIndex        =   98
         Top             =   825
         Width           =   420
      End
      Begin VB.Label Label21 
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
         Index           =   5
         Left            =   2625
         TabIndex        =   97
         Top             =   825
         Width           =   3300
      End
      Begin VB.Label Label39 
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
         Index           =   6
         Left            =   525
         TabIndex        =   96
         ToolTipText     =   "Código"
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   6
         Left            =   6720
         TabIndex        =   95
         Top             =   1140
         Width           =   420
      End
      Begin VB.Label Label21 
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
         Index           =   6
         Left            =   2625
         TabIndex        =   94
         Top             =   1140
         Width           =   3300
      End
      Begin VB.Label Label39 
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
         Index           =   7
         Left            =   525
         TabIndex        =   93
         ToolTipText     =   "Código"
         Top             =   1455
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   7
         Left            =   6720
         TabIndex        =   92
         Top             =   1455
         Width           =   420
      End
      Begin VB.Label Label21 
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
         Index           =   7
         Left            =   2625
         TabIndex        =   91
         Top             =   1455
         Width           =   3300
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Descripción:"
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
         Left            =   2415
         TabIndex        =   90
         Top             =   300
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Asignacion de Papeles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2500
      Left            =   210
      TabIndex        =   32
      Top             =   1785
      Width           =   8730
      Begin VB.PictureBox Picture6 
         Height          =   1020
         Left            =   360
         ScaleHeight     =   960
         ScaleWidth      =   8085
         TabIndex        =   130
         Top             =   1320
         Width           =   8140
         Begin VB.VScrollBar VScroll1 
            Height          =   980
            Left            =   7860
            Max             =   340
            TabIndex        =   252
            Top             =   0
            Width           =   225
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   4440
            Left            =   0
            ScaleHeight     =   4440
            ScaleWidth      =   7860
            TabIndex        =   131
            Top             =   0
            Width           =   7860
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   2
               Left            =   2535
               TabIndex        =   176
               Top             =   660
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   1
               Left            =   2535
               TabIndex        =   175
               Top             =   390
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   0
               Left            =   2535
               TabIndex        =   174
               Top             =   120
               Width           =   175
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   0
               Left            =   5370
               TabIndex        =   173
               Top             =   120
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   1
               Left            =   5370
               TabIndex        =   172
               ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
               Top             =   390
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   2
               Left            =   5370
               TabIndex        =   171
               Top             =   660
               Width           =   1065
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
               Height          =   250
               Index           =   0
               Left            =   1800
               TabIndex        =   170
               Top             =   120
               Width           =   175
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
               Height          =   250
               Index           =   1
               Left            =   1800
               TabIndex        =   169
               Top             =   390
               Width           =   175
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
               Height          =   250
               Index           =   2
               Left            =   1800
               TabIndex        =   168
               Top             =   660
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   5
               Left            =   2535
               TabIndex        =   167
               Top             =   1500
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   4
               Left            =   2535
               TabIndex        =   166
               Top             =   1230
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   3
               Left            =   2535
               TabIndex        =   165
               Top             =   960
               Width           =   175
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   3
               Left            =   5370
               TabIndex        =   164
               Top             =   960
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   4
               Left            =   5370
               TabIndex        =   163
               ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
               Top             =   1230
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   5
               Left            =   5370
               TabIndex        =   162
               Top             =   1500
               Width           =   1065
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
               Height          =   250
               Index           =   3
               Left            =   1800
               TabIndex        =   161
               Top             =   960
               Width           =   175
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
               Height          =   250
               Index           =   4
               Left            =   1800
               TabIndex        =   160
               Top             =   1230
               Width           =   175
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
               Height          =   250
               Index           =   5
               Left            =   1800
               TabIndex        =   159
               Top             =   1500
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   8
               Left            =   2535
               TabIndex        =   158
               Top             =   2340
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   7
               Left            =   2535
               TabIndex        =   157
               Top             =   2070
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   6
               Left            =   2535
               TabIndex        =   156
               Top             =   1800
               Width           =   175
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   6
               Left            =   5370
               TabIndex        =   155
               Top             =   1800
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   7
               Left            =   5370
               TabIndex        =   154
               ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
               Top             =   2070
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   8
               Left            =   5370
               TabIndex        =   153
               Top             =   2340
               Width           =   1065
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
               Height          =   250
               Index           =   6
               Left            =   1800
               TabIndex        =   152
               Top             =   1800
               Width           =   175
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
               Height          =   250
               Index           =   7
               Left            =   1800
               TabIndex        =   151
               Top             =   2070
               Width           =   175
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
               Height          =   250
               Index           =   8
               Left            =   1800
               TabIndex        =   150
               Top             =   2340
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   11
               Left            =   2535
               TabIndex        =   149
               Top             =   3180
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   10
               Left            =   2535
               TabIndex        =   148
               Top             =   2910
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   9
               Left            =   2535
               TabIndex        =   147
               Top             =   2640
               Width           =   175
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   9
               Left            =   5370
               TabIndex        =   146
               Top             =   2640
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   10
               Left            =   5370
               TabIndex        =   145
               ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
               Top             =   2910
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   11
               Left            =   5370
               TabIndex        =   144
               Top             =   3180
               Width           =   1065
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
               Height          =   250
               Index           =   9
               Left            =   1800
               TabIndex        =   143
               Top             =   2640
               Width           =   175
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
               Height          =   250
               Index           =   10
               Left            =   1800
               TabIndex        =   142
               Top             =   2910
               Width           =   175
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
               Height          =   250
               Index           =   11
               Left            =   1800
               TabIndex        =   141
               Top             =   3180
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   14
               Left            =   2535
               TabIndex        =   140
               Top             =   4020
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   13
               Left            =   2535
               TabIndex        =   139
               Top             =   3750
               Width           =   175
            End
            Begin VB.CommandButton Command7 
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
               Height          =   255
               Index           =   12
               Left            =   2535
               TabIndex        =   138
               Top             =   3480
               Width           =   175
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   12
               Left            =   5370
               TabIndex        =   137
               Top             =   3480
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   13
               Left            =   5370
               TabIndex        =   136
               ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
               Top             =   3750
               Width           =   1065
            End
            Begin VB.TextBox Text11 
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
               Height          =   250
               Index           =   14
               Left            =   5370
               TabIndex        =   135
               Top             =   4020
               Width           =   1065
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
               Height          =   250
               Index           =   12
               Left            =   1800
               TabIndex        =   134
               Top             =   3480
               Width           =   175
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
               Height          =   250
               Index           =   13
               Left            =   1800
               TabIndex        =   133
               Top             =   3750
               Width           =   175
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
               Height          =   250
               Index           =   14
               Left            =   1800
               TabIndex        =   132
               Top             =   4020
               Width           =   175
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   2
               Left            =   2220
               TabIndex        =   251
               Top             =   660
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   1
               Left            =   2220
               TabIndex        =   250
               Top             =   390
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   0
               Left            =   2220
               TabIndex        =   249
               Top             =   120
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   2
               Left            =   6735
               TabIndex        =   248
               Top             =   660
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   2
               Left            =   4005
               TabIndex        =   247
               Top             =   660
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   2
               Left            =   2955
               TabIndex        =   246
               Top             =   660
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   1
               Left            =   6735
               TabIndex        =   245
               Top             =   390
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   1
               Left            =   4005
               TabIndex        =   244
               Top             =   390
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   1
               Left            =   2955
               TabIndex        =   243
               Top             =   390
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   0
               Left            =   6735
               TabIndex        =   242
               Top             =   120
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   0
               Left            =   4005
               TabIndex        =   241
               Top             =   120
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   0
               Left            =   2955
               TabIndex        =   240
               Top             =   120
               Width           =   810
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   239
               Top             =   120
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   238
               Top             =   390
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   237
               Top             =   660
               Width           =   1755
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   5
               Left            =   2220
               TabIndex        =   236
               Top             =   1500
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   4
               Left            =   2220
               TabIndex        =   235
               Top             =   1230
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   3
               Left            =   2220
               TabIndex        =   234
               Top             =   960
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   5
               Left            =   6735
               TabIndex        =   233
               Top             =   1500
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   5
               Left            =   4005
               TabIndex        =   232
               Top             =   1500
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   5
               Left            =   2955
               TabIndex        =   231
               Top             =   1500
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   4
               Left            =   6735
               TabIndex        =   230
               Top             =   1230
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   4
               Left            =   4005
               TabIndex        =   229
               Top             =   1230
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   4
               Left            =   2955
               TabIndex        =   228
               Top             =   1230
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   3
               Left            =   6735
               TabIndex        =   227
               Top             =   960
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   3
               Left            =   4005
               TabIndex        =   226
               Top             =   960
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   3
               Left            =   2955
               TabIndex        =   225
               Top             =   960
               Width           =   810
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   224
               Top             =   960
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   223
               Top             =   1230
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   222
               Top             =   1500
               Width           =   1755
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   8
               Left            =   2220
               TabIndex        =   221
               Top             =   2340
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   7
               Left            =   2220
               TabIndex        =   220
               Top             =   2070
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   6
               Left            =   2220
               TabIndex        =   219
               Top             =   1800
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   8
               Left            =   6735
               TabIndex        =   218
               Top             =   2340
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   8
               Left            =   4005
               TabIndex        =   217
               Top             =   2340
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   8
               Left            =   2955
               TabIndex        =   216
               Top             =   2340
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   7
               Left            =   6735
               TabIndex        =   215
               Top             =   2070
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   7
               Left            =   4005
               TabIndex        =   214
               Top             =   2070
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   7
               Left            =   2955
               TabIndex        =   213
               Top             =   2070
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   6
               Left            =   6735
               TabIndex        =   212
               Top             =   1800
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   6
               Left            =   4005
               TabIndex        =   211
               Top             =   1800
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   6
               Left            =   2955
               TabIndex        =   210
               Top             =   1800
               Width           =   810
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   209
               Top             =   1800
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   208
               Top             =   2070
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   207
               Top             =   2340
               Width           =   1755
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   11
               Left            =   2220
               TabIndex        =   206
               Top             =   3180
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   10
               Left            =   2220
               TabIndex        =   205
               Top             =   2910
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   9
               Left            =   2220
               TabIndex        =   204
               Top             =   2640
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   11
               Left            =   6735
               TabIndex        =   203
               Top             =   3180
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   11
               Left            =   4005
               TabIndex        =   202
               Top             =   3180
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   11
               Left            =   2955
               TabIndex        =   201
               Top             =   3180
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   10
               Left            =   6735
               TabIndex        =   200
               Top             =   2910
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   10
               Left            =   4005
               TabIndex        =   199
               Top             =   2910
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   10
               Left            =   2955
               TabIndex        =   198
               Top             =   2910
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   9
               Left            =   6735
               TabIndex        =   197
               Top             =   2640
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   9
               Left            =   4005
               TabIndex        =   196
               Top             =   2640
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   9
               Left            =   2955
               TabIndex        =   195
               Top             =   2640
               Width           =   810
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   194
               Top             =   2640
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   193
               Top             =   2910
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   192
               Top             =   3180
               Width           =   1755
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   14
               Left            =   2220
               TabIndex        =   191
               Top             =   4020
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   13
               Left            =   2220
               TabIndex        =   190
               Top             =   3750
               Width           =   330
            End
            Begin VB.Label Label55 
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
               Height          =   255
               Index           =   12
               Left            =   2220
               TabIndex        =   189
               Top             =   3480
               Width           =   330
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   14
               Left            =   6735
               TabIndex        =   188
               Top             =   4020
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   14
               Left            =   4005
               TabIndex        =   187
               Top             =   4020
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   14
               Left            =   2955
               TabIndex        =   186
               Top             =   4020
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   13
               Left            =   6735
               TabIndex        =   185
               Top             =   3750
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   13
               Left            =   4005
               TabIndex        =   184
               Top             =   3750
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   13
               Left            =   2955
               TabIndex        =   183
               Top             =   3750
               Width           =   810
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   12
               Left            =   6735
               TabIndex        =   182
               Top             =   3480
               Width           =   1020
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   12
               Left            =   4005
               TabIndex        =   181
               Top             =   3480
               Width           =   1020
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Height          =   255
               Index           =   12
               Left            =   2955
               TabIndex        =   180
               Top             =   3480
               Width           =   810
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   179
               Top             =   3480
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   178
               Top             =   3750
               Width           =   1755
            End
            Begin VB.Label Label23 
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
               Height          =   255
               Index           =   14
               Left            =   120
               TabIndex        =   177
               Top             =   4020
               Width           =   1755
            End
         End
      End
      Begin VB.CommandButton Command17 
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
         Left            =   2940
         TabIndex        =   39
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         TabIndex        =   129
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label3 
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
         Height          =   405
         Left            =   4320
         TabIndex        =   128
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label1 
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
         Height          =   405
         Left            =   3240
         TabIndex        =   127
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7065
         TabIndex        =   126
         Top             =   1050
         Width           =   885
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   125
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
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
         Left            =   4725
         TabIndex        =   105
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.Rec.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -840
         TabIndex        =   40
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -525
         TabIndex        =   38
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label27 
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
         Left            =   840
         TabIndex        =   37
         ToolTipText     =   "Doble-Click Para Lectura Código de Barras"
         Top             =   315
         Width           =   2175
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
         Height          =   285
         Left            =   840
         TabIndex        =   36
         ToolTipText     =   "Descripción"
         Top             =   630
         Width           =   7680
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Metros:"
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
         Left            =   2940
         TabIndex        =   35
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Cantidad:"
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
         Left            =   6090
         TabIndex        =   34
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   7455
         TabIndex        =   33
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
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
      Height          =   1485
      Left            =   6810
      TabIndex        =   24
      Top             =   210
      Width           =   2130
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
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
         Left            =   850
         TabIndex        =   30
         Top             =   1050
         Width           =   1200
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
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
         Left            =   850
         TabIndex        =   29
         Top             =   735
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
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
         Left            =   850
         TabIndex        =   28
         Top             =   420
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   -525
         TabIndex        =   27
         Top             =   1095
         Width           =   1360
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyect:"
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
         Left            =   -525
         TabIndex        =   26
         Top             =   465
         Width           =   1360
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
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
         Left            =   -525
         TabIndex        =   25
         Top             =   780
         Width           =   1360
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
      Top             =   315
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
      Left            =   3120
      TabIndex        =   0
      Top             =   315
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8500
      Left            =   9135
      ScaleHeight     =   8445
      ScaleWidth      =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "TRAZORF07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ver y Reimprimir Vale de Materiales"
         Top             =   2845
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Height          =   650
         Left            =   120
         Picture         =   "TRAZORF07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   4695
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   1005
         Left            =   105
         Picture         =   "TRAZORF07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar Asignacion de Materiales a O/F"
         Top             =   6300
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         Picture         =   "TRAZORF07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   4035
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
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
         Picture         =   "TRAZORF07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "TRAZORF07.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anular Vale de Materiales"
         Top             =   2205
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   7455
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
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
         Height          =   690
         Left            =   120
         Picture         =   "TRAZORF07.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "TRAZORF07.frx":14BE
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7770
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tintas e Insumos Trazables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2050
      Left            =   210
      TabIndex        =   11
      Top             =   4320
      Width           =   8730
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
         Index           =   3
         Left            =   5040
         TabIndex        =   114
         Top             =   1365
         Width           =   175
      End
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
         Index           =   2
         Left            =   5040
         TabIndex        =   112
         Top             =   1050
         Width           =   175
      End
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
         Index           =   1
         Left            =   5040
         TabIndex        =   110
         Top             =   735
         Width           =   175
      End
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
         Index           =   0
         Left            =   5040
         TabIndex        =   107
         Top             =   420
         Width           =   175
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
         Index           =   3
         Left            =   7455
         TabIndex        =   71
         Top             =   1365
         Width           =   1065
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
         Left            =   4410
         TabIndex        =   70
         Top             =   1365
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
         Left            =   2520
         TabIndex        =   68
         Top             =   1365
         Width           =   175
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
         Index           =   2
         Left            =   7455
         TabIndex        =   64
         Top             =   1050
         Width           =   1065
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
         Left            =   4410
         TabIndex        =   63
         Top             =   1050
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
         Left            =   2520
         TabIndex        =   61
         Top             =   1050
         Width           =   175
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
         Index           =   1
         Left            =   7455
         TabIndex        =   57
         Top             =   735
         Width           =   1065
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
         Left            =   4410
         TabIndex        =   56
         Top             =   735
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
         Left            =   2520
         TabIndex        =   54
         Top             =   735
         Width           =   175
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
         Index           =   0
         Left            =   7455
         TabIndex        =   45
         Top             =   420
         Width           =   1065
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
         Left            =   4410
         TabIndex        =   44
         Top             =   420
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
         Left            =   2520
         TabIndex        =   42
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label52 
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
         Index           =   3
         Left            =   4725
         TabIndex        =   115
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   2
         Left            =   4725
         TabIndex        =   113
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   1
         Left            =   4725
         TabIndex        =   111
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   735
         Width           =   330
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   109
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label52 
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
         Index           =   0
         Left            =   4725
         TabIndex        =   108
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
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
         Left            =   7455
         TabIndex        =   79
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total: ..................."
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
         Left            =   5145
         TabIndex        =   78
         Top             =   1755
         Width           =   2280
      End
      Begin VB.Label Label21 
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
         Index           =   3
         Left            =   2835
         TabIndex        =   77
         Top             =   1365
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   2
         Left            =   2835
         TabIndex        =   76
         Top             =   1050
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   1
         Left            =   2835
         TabIndex        =   75
         Top             =   735
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   0
         Left            =   2835
         TabIndex        =   74
         Top             =   420
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
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
         Index           =   3
         Left            =   6720
         TabIndex        =   73
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   5355
         TabIndex        =   72
         Top             =   1365
         Width           =   1050
      End
      Begin VB.Label Label40 
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
         Index           =   3
         Left            =   3045
         TabIndex        =   69
         ToolTipText     =   "Código"
         Top             =   1365
         Width           =   1440
      End
      Begin VB.Label Label39 
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
         Index           =   3
         Left            =   840
         TabIndex        =   67
         ToolTipText     =   "Código"
         Top             =   1365
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   2
         Left            =   6720
         TabIndex        =   66
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   5355
         TabIndex        =   65
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label40 
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
         Index           =   2
         Left            =   3045
         TabIndex        =   62
         ToolTipText     =   "Código"
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label Label39 
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
         Index           =   2
         Left            =   840
         TabIndex        =   60
         ToolTipText     =   "Código"
         Top             =   1050
         Width           =   1755
      End
      Begin VB.Label Label42 
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
         Index           =   1
         Left            =   6720
         TabIndex        =   59
         Top             =   735
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   5355
         TabIndex        =   58
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label40 
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
         Index           =   1
         Left            =   3045
         TabIndex        =   55
         ToolTipText     =   "Código"
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label39 
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
         Index           =   1
         Left            =   840
         TabIndex        =   53
         ToolTipText     =   "Código"
         Top             =   735
         Width           =   1755
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5985
         TabIndex        =   52
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3780
         TabIndex        =   51
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1365
         TabIndex        =   50
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label44 
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
         Height          =   405
         Left            =   -1155
         TabIndex        =   49
         Top             =   450
         Width           =   1965
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6090
         TabIndex        =   48
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label42 
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
         Index           =   0
         Left            =   6720
         TabIndex        =   47
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   5355
         TabIndex        =   46
         Top             =   420
         Width           =   1050
      End
      Begin VB.Label Label40 
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
         Index           =   0
         Left            =   3045
         TabIndex        =   43
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label39 
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
         Index           =   0
         Left            =   840
         TabIndex        =   41
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   1755
      End
   End
   Begin VB.Label Label51 
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
      Left            =   5300
      TabIndex        =   106
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label7 
      Height          =   225
      Left            =   210
      TabIndex        =   31
      Top             =   1470
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vale Nro:"
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
      Left            =   3885
      TabIndex        =   22
      Top             =   795
      Width           =   1335
   End
   Begin VB.Label Label32 
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
      Left            =   1365
      TabIndex        =   21
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   105
      TabIndex        =   20
      Top             =   1220
      Width           =   1125
   End
   Begin VB.Label Label30 
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
      Height          =   495
      Left            =   1365
      TabIndex        =   19
      Top             =   1155
      Width           =   5370
   End
   Begin VB.Label Label28 
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
      Left            =   1365
      TabIndex        =   18
      Top             =   735
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   -315
      TabIndex        =   17
      Top             =   800
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
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
      Left            =   4200
      TabIndex        =   16
      Top             =   165
      Width           =   1335
   End
   Begin VB.Label Label24 
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
      Left            =   5565
      TabIndex        =   15
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Fabr. Nº:"
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
      Left            =   105
      TabIndex        =   14
      Top             =   380
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
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
      Left            =   1470
      TabIndex        =   13
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   1470
      TabIndex        =   12
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "TRAZORF07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%

Private Sub Command1_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then
      If COMALTER%("¿ Realmente Cancela la Asignación\Automática de Lotes Propuesta ?\\No, Conservarla\Si, Cancelarla") = 2 Then GoTo 10
      Exit Sub
   End If
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command10_Click()

    Command10.Enabled = False
    FIKARDEX09.Show 1
    Command10.Enabled = True
    
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
    'VALIDAR CANTIDAD
    VALONU = XVALO(Label38)
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    '
    For U& = 0 To Text11.Count
       If XVALO(Text11(U&)) > 0 Then
          If XVALO(Label55(U&)) < 1 Then
             Call MENSERR(24, "Falta Indicar Depósito de Consumo")
             GoTo 99
          End If
          If TRIM$(Label23(U&)) = "" Then
             Call MENSERR(24, "Papel Asignado sin Lote / Partida")
             GoTo 99
          End If
       End If
    Next U&
    '
    For U& = 0 To Text11.Count
       If XVALO(Text11(U&)) > XVALO(Label25(U&)) Then
         Call MENSERR(24, "Metros Asignados Superan Largo de Lote\en Papel según N.C.Rec. " + TRIM$(Label23(U&)))
         GoTo 99
       End If
    Next U&
    '
    For U& = 0 To Text1.Count
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label39(U&)) = "" Then
             Call MENSERR(24, "Tinta / Insumo sin Código de Material")
             GoTo 99
          End If
       End If
    Next U&
    '
    For U& = 0 To 3
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label40(U&)) = "" Then
             Call MENSERR(24, "Tinta sin Lote / Partida")
             GoTo 99
          End If
       End If
    Next U&
    '
    For U& = 0 To 3
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label39(U&)) <> "" Then
             CIXI% = CODINT%(Label39(U&))
             If ESTRAZABLE%(CIXI%) < 1 Then
                Call MENSERR(24, "Material No Trazable '" + CODEXT$(CIXI%) + "' Entre Tintas")
                GoTo 99
             End If
          End If
          If XVALO(Label52(U&)) < 1 Then
             Call MENSERR(24, "Falta Indicar Depósito de Consumo")
             GoTo 99
          End If
          If XVALO(Text1(U&)) > XVALO(Label41(U&)) Then
             Call MENSERR(24, "Exceso de Consumo de Tinta")
             GoTo 99
           End If
       End If
    Next U&
    '
    For U& = 4 To 7
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label39(U&)) <> "" Then
             CIXI% = CODINT%(Label39(U&))
             If ESTRAZABLE%(CIXI%) > 0 Then
                Call MENSERR(24, "Material Trazable '" + CODEXT$(CIXI%) + "' Entre Insumos")
                GoTo 99
             End If
             DEPX% = XVALO(Label52(U&))
             If DEPX% < 1 Then
                Call MENSERR(24, "Falta Indicar Depósito de Consumo")
                GoTo 99
             End If
          End If
          If XVALO(Text1(U&)) > STODEPOS(CIXI%, DEPX%) Then
             Call MENSERR(24, "Stock Insuficiente en Código '" + CODEXT$(CIXI%) + "'")
             GoTo 99
          End If
       End If
    Next U&
    '
    Call BLOQARCH("ECORDEP") ' Agregado para provocar una demora para mas seguridad de bloqueo
    TINI = Int(Timer)
    While Int(Timer) < TINI + 5: Wend
    '
    Call BLOQARCH("ECORDEP") ' VALIDACION DE QUE NO SE HAYA USADO EL NUMERO DE COMPROBANTE
    NVM$ = TRIM$(ProNroComprobante("CF"))
    NVMLABEL$ = TRIM$(Mid$(Label51, 4))
    If XVALO(NVM$) > XVALO(NVMLABEL$) Then
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
       Label51 = "CF-" + NVM$
       Call CIERRARCH("ECORDEP")
       Call COMUNI("FLEXOFT ha Re-Numerado el C/Material\como C/Material '" + Label51 + "'.\  \Verifique y Vuelva a Grabar.")
       GoTo 99
    End If
    Call CIERRARCH("ECORDEP")
    '
    'En este caso se realiza primero la impresion para poder obtener la ubicacion de la MP,
    'Esto es porque en la grabacion, en caso que el consumo sea total, el saldo es 0 (cero)
    'y por lo tanto el lote de la MP no existe mas, por ende no puede tener ubicacion
    '
    Call PRIVALEMA
    Call GRABACONSUM
    '
    VScroll1.Value = 0
98  Screen.MousePointer = 1
    Call Label32_DblClick
99  Screen.MousePointer = 1
    Command13.Enabled = True
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then
       Call CIERRARCH("*.*")
       Hide
    End If
    '
    
End Sub
'
Sub GRABACONSUM()
     '
     FF% = FECHANUM(Label24)
     CMO$ = "CF"    ' CONSUMO DE FABRICACION
     DOPRI$ = Label51
     DOSEC$ = Label32
     REFE$ = "Sol-Mat-Dep. p/" + Label32
     NUORIT% = 0
     '
     For YX& = 0 To 14
        NLOTX$ = TRIM$(Label23(YX&))
        If NLOTX$ <> "" Then
           CIXI% = CODINT%(Label27)
           NUORIT% = NUORIT% + 1
           CANTI = (-1) * XVALO(Label35(YX&))
           LOTE$ = Label23(YX&)
           DepOrig% = XVALO(Label55(YX&))
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                         '\\\OT-06-0244-WAR-20/06/06///
        End If
     Next YX&
     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = Label40(YX&)
           CANTI = (-1) * XVALO(Text1(YX&))
           DepOrig% = XVALO(Label52(YX&))
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = ""
           CANTI = (-1) * XVALO(Text1(YX&))
           DepOrig% = XVALO(Label52(YX&))
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
        End If
     Next YX&
     '
End Sub
'
Sub PRIVALEMA()
     '
     FORIPRE$ = TRIM$(Control$("", "FORCONSU"))
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
     DOCUORIG$ = Label32
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
     For YX& = 0 To 14
        NLOTX$ = TRIM$(Label23(YX&))
        If NLOTX$ <> "" Then
          If XVALO(Label35(YX&)) >= 0.05 Then
            CIXI% = CODINT%(Label27)
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
            TETABU1$(2, CAITAB1%) = Label35(YX&)
            TETABU1$(3, CAITAB1%) = Label27
            TETABU1$(4, CAITAB1%) = Label6
            TETABU1$(5, CAITAB1%) = Label23(YX&)
            TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label23(YX&))
          End If
        End If
     Next YX&
     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
          If XVALO(Text1(YX&)) >= 0.05 Then
            CIXI% = CODINT%(CODEX$)
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
            TETABU1$(2, CAITAB1%) = Text1(YX&)
            TETABU1$(3, CAITAB1%) = CODEX$
            TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
            TETABU1$(5, CAITAB1%) = Label40(YX&)
            TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label40(YX&))
          End If
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
          If XVALO(Text1(YX&)) >= 0.05 Then
            CIXI% = CODINT%(CODEX$)
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
            TETABU1$(2, CAITAB1%) = Text1(YX&)
            TETABU1$(3, CAITAB1%) = CODEX$
            TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
            TETABU1$(5, CAITAB1%) = ""
            TETABU1$(6, CAITAB1%) = UBISTO(CIXI%)
          End If
        End If
     Next YX&
     '
     Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
     Screen.MousePointer = 1
     '
     Call CIERRARCH("*.*")
     '
End Sub

Sub Command14_Click()
    '\\\OT-06-0476-RG-07/12/06///
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    OPOFTRA07.Show 1
    '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Command16_Click()
    '
    Command16.Enabled = False
    ' ARMA LISTA DE SELECCION
    SQLX$ = " SELECT DISTINCT DOPRILAR, ReferCor, FECHMOV FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF'"
    SQLX$ = SQLX$ + " ORDER BY FECHMOV ASC, DOPRILAR ASC "
    Call APERBASDAT("STOCKS")
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With MOVITMP
      JJ& = 0
      Do While (.EOF = False)
        y$ = REGBLAN("ECORDEP")
        Call REPLA(y$, !DOPRILAR, 1, 12)
        FEX = CVINT(!fechmov)
        Call REPLA(y$, FECHATEX$(FEX) + "  " + !ReferCor, 13, 52)
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", y$, JJ&)
      .MoveNext
      Loop
    End With
    Set MOVITMP = Nothing
    Call SETULTREG("!ECORDEP", JJ&)
    Call FRESHECHO("!ECORDEP")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
        Call MENSERR(24, "No Hay Vales de Material para Anular")
        GoTo 99
    End If
    '
    Call SELECHO("!ECORDEP/Indice General de Vales de Material")
    NVALE$ = TRIM$(VALACT1$("!ECORDEP"))
    If NVALE$ = "" Then GoTo 99
    '
    Call SELECHO("SELFECHA")
    FECHVA$ = TRIM$(VALACT1$("SELFECHA"))
    If FECHVA$ = "" Then GoTo 99
    FECHVAL% = FECHANUM(FECHVA$)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE DoPriLar = '" & NVALE$ & "'"
    SQLX$ = SQLX$ + " AND FechMov = '" & FETEXCOR1$(FECHVAL%) & "'"
    Dim MOVIANU As ADODB.Recordset
    Set MOVIANU = New ADODB.Recordset
    MOVIANU.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    With MOVIANU
      Do While (.EOF = False)
        LOTE$ = !IdenLote
        CIXI% = !Cod_Inte
        !Cantsalid = 0
        .Update
        Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
        .MoveNext
      Loop
    End With
    Set MOVIANU = Nothing
     FLEXCONN.Execute (FILTSQL$("Delete from MOVISTO where DopriLar = '" & NVALE$ & "'"))
    Call MENSERR(24, "Se Ha Anulado el Vale de Materiales Numero " + NVALE$)
    '
99  Screen.MousePointer = 1
    Command16.Enabled = True
    '
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
       CICOMPO% = CODINT%(VDEVU$)
       Label27 = VDEVU$
       CACONSU = 0
       Text4 = FORMATNUM$(CACONSU, "f10.1")
    End If
    Command17.Enabled = True
End Sub

Private Sub Command18_Click(Index As Integer)
   If Label32 <> "" Then
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
       CICOMPO% = CODINT%(VDEVU$)
       Label39(Index) = VDEVU$
    End If
   End If
End Sub

Private Sub Command19_Click(Index As Integer)
  If Label32 <> "" Then
     CIXI% = CODINT%(Label39(Index))
     If CIXI% > 0 Then
        If ESTRAZABLE(CIXI%) < 1 Then
           Call MENSERR(24, "Material No Trazable")
           Exit Sub
        End If
        Call ARMALILOT(CIXI%, 0, 0, 1)
        Call SELECHO("!LILOTSE")
        VDEVU$ = VALACT1$("!LILOTSE")
        VDEV2$ = VALACT2$("!LILOTSE")
        If VDEVU$ <> "" Then
           Label40(Index) = VDEVU$
        End If
    End If
  End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Label52(Index) = VALACT1$("TADEPOSI")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
'      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
'      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call CommSel("SELFECHA", Command4, Label24)
End Sub

Private Sub Command5_Click()
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Call CIERRAOF07.ACTECOR_ACHERNAR(2)     ' para seleccionar solo las activas lanzadas
    Call CommSel("ECOACHER/Indice de Ordenes de Fabricación Pendientes", Command5, Label32)
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    '
End Sub

Private Sub Command6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command6_Click(Index)
   End If
End Sub

Private Sub Command7_Click(Index As Integer)
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Label55(Index) = VALACT1$("TADEPOSI")
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   SQLX$ = " SELECT DISTINCT DOPRILAR, ReferCor, FECHMOV FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF'"
   SQLX$ = SQLX$ + " ORDER BY FECHMOV ASC, DOPRILAR ASC "
   Call APERBASDAT("STOCKS")
   Dim MOVITMP As ADODB.Recordset
   Set MOVITMP = New ADODB.Recordset
   Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With MOVITMP
     JJ& = 0
     Do While (.EOF = False)
       y$ = REGBLAN("ECORDEP")
       Call REPLA(y$, !DOPRILAR, 1, 12)
       FEX = CVINT(!fechmov)
       Call REPLA(y$, FECHATEX$(FEX) + "  " + !ReferCor, 13, 52)
       JJ& = JJ& + 1
       Call GRAREG("!ECORDEP", y$, JJ&)
     .MoveNext
     Loop
   End With
   Set MOVITMP = Nothing
   Call SETULTREG("!ECORDEP", JJ&)
   Call FRESHECHO("!ECORDEP")
2  Call SELECHO("!ECORDEP/Indice General de Vales de Material")
   NVALE$ = TRIM$(VALACT1$("!ECORDEP"))
   If NVALE$ = "" Then GoTo 99
   '
   TIDOCUM$ = "Vale de Consumo de Materiales " + NVALE$
   Screen.MousePointer = 11
   '
   For UI& = ULTREG&("PDOCLST") To 1 Step -1
     XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 9
      End If
   Next UI&
   Screen.MousePointer = 1
   Call COMUNI("Documento no Encontrado")
   GoTo 99

9  Call MUESTRADOC(DOCUNU&)
   GoTo 2
   '
99 Command8.Enabled = True
   '
End Sub

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

Private Sub Command6_Click(Index As Integer)
    If TRIM$(Label27) = "" Or CODINT%(Label27) < 0 Then
       Exit Sub
    End If
    '
    For KKU& = 0 To Index - 1
       If Label23(KKU&) = "" Then
          Call MENSERR(24, "No Dejar Campos no Informados")
          Exit Sub
       End If
    Next KKU&
    '
    CIXI% = CODINT%(Label27)
    '
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    If VDEVU$ <> "" Then
      For KKU& = 0 To 14
        If KKU& <> Index Then
          If Label23(KKU&) = VDEVU$ Then
             Call MENSERR(24, "No Válido Repetir Lotes")
             Exit Sub
          End If
        End If
      Next KKU&
      Label23(Index) = VDEVU$
      '
      Label12(Index) = TRIM$(Mid$(VDEV2$, 39, 8))
      Call MUESTRASEL(Index)
    End If
    '
    Text11(Index).SetFocus
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ASEMBLI") > 0 Then
        TRAZORF14.Show 1
        Call FINAL("")
    End If
    '
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       Frame9.Visible = True
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
    VScroll1.Value = 0
    '
End Sub



Private Sub Label23_Change(Index As Integer)
    CIXI% = CODINT%(TRIM$(Label27))
    IDLOT$ = TRIM$(Label23(Index))
    Label55(Index) = FORMATNUM$(DEPOLOT%(CIXI%, IDLOT$), "I3")
    Call MUESTRASEL(Index)
End Sub

Private Sub Label25_Change(Index As Integer)
   Text11(Index) = TRIM$(Label25(Index))
End Sub

Private Sub Label27_Change()
   CIXI% = CODINT%(Label27)
   For KKU& = 0 To 14
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Label12(KKU&) = ""
       Label25(KKU&) = ""
       Label55(KKU&) = ""
   Next KKU&
   '
   For KKI% = 0 To 7: Text1(KKI%) = " ": Text1(KKI%) = "": Next KKI%
   If CIXI% > 0 Then
      Label6 = DESCRIT0$(CIXI%)
      Call ARMALILOT(CIXI%, 0, 0, 1, 1)
      Screen.MousePointer = 1
   End If
End Sub

Private Sub Label27_DblClick()
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 1 Then
       Label27.Caption = TRIM$(Left$(TTXX$, 16))
       LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
       Label23(0).Caption = LOTSE$
       '
       Label12(0) = ""
       For KKL& = 1 To ULTREG("!LILOTSE")
         XX$ = REGLEIDO$("!LILOTSE", KKL&)
         If TRIM$(Left$(XX$, 12)) = LOTSE$ Then
            Label12(0) = TRIM$(Mid$(XX$, 56, 8))
            Call MUESTRASEL(0)
            Exit Sub
         End If
       Next KKL&
    End If
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
    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Label28 = ""
    Label30 = ""
    Text8 = NORFA$
    Text5 = ""
    Text6 = ""
    Text2 = ""
    Text3 = ""
    Label34 = ""
    Label27 = ""
    Label6 = ""
    Label16 = ""
    Text4 = ""
    Label38 = ""
    '
    For KKU& = 0 To 14
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Label12(KKU&) = ""
       Label25(KKU&) = ""
       Label55(KKU&) = ""
    Next KKU&
    '
    For KKU& = 0 To 3
       Label39(KKU&) = ""
       Text1(KKU&) = ""
       Label40(KKU&) = ""
       Label41(KKU&) = ""
       Label42(KKU&) = ""
       Label52(KKU&) = ""
    Next KKU&
    '
    For KKU& = 4 To 7
       Label39(KKU&) = ""
       Text1(KKU&) = ""
       Label21(KKU&) = ""
       Label42(KKU&) = ""
       Label52(KKU&) = ""
    Next KKU&
    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
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
              If !Statuorf >= 3 Then CAPENDE# = 0
           Label34 = FORMATNUM$(!CanProy, "F10.1")
           Label37 = FORMATNUM$(!CanApro, "F10.1")
           Label38 = FORMATNUM$(CAPENDE#, "F10.1")
           '
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

    Label27.Caption = ""
    Label6.Caption = ""
    Text4.TEXT = ""
  
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

Private Sub Label35_Change(Index As Integer)
   SUTRAZA = XVALO(Label35(0)) + XVALO(Label35(1)) + XVALO(Label35(2))
   Label16 = FORMATNUM$(SUTRAZA, "F8.1")
End Sub

Private Sub Label39_Change(Index As Integer)
    CIXI% = CODINT%(Label39(Index).Caption)
    On Error Resume Next
    Label40(Index) = ""
    On Error GoTo 0
    Label21(Index) = DESCRIT0$(CIXI%)
    Label42(Index) = UNIMED(CIXI%)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label39_DblClick(Index As Integer)
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 1 Then
       Label39(Index).Caption = TRIM$(Left$(TTXX$, 16))
       CIXI% = CODINT%(Label39(Index).Caption)
       LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
       Label40(Index).Caption = LOTSE$
       '
       Label12(0) = ""
       For KKL& = 1 To ULTREG("!LILOTSE")
         XX$ = REGLEIDO$("!LILOTSE", KKL&)
         If TRIM$(Left$(XX$, 12)) = LOTSE$ Then
            On Error Resume Next
            Label40(Index) = TRIM$(Left$(XX$, 12))
            Label52(Index) = DEPOLOT%(CIXI%, LOTSE$)
            On Error GoTo 0
            Call MUESTRASEL1(Index)
            Exit Sub
         End If
       Next KKL&
    End If
End Sub

Private Sub Label40_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label41_Change(Index As Integer)
    Text1(Index) = Label41(Index)
End Sub

Private Sub Label52_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label55_Change(Index As Integer)
    Call MUESTRASEL(Index)
End Sub

Private Sub Text1_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 3
       If Label39(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text1(KKI%))
         Else
           Text1(KKI%) = ""
       End If
    Next KKI%
    Label49 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
End Sub

Private Sub text1_lostfocus(Index As Integer)
    Text1(Index) = FORMATNUM(XVALO(Text1(Index)), "F10.1")
End Sub

Private Sub Text11_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 14
       If Label23(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text11(KKI%))
         Else
           Text11(KKI%) = ""
       End If
       Label35(Index) = FORMATNUM$(XVALO(Text11(Index)) * XVALO(Label12(Index)) / 1000, "F10.1")
    Next KKI%
    Label50 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
    '
End Sub

Private Sub Text11_DblClick(Index As Integer)
    Text11(Index) = ""
End Sub

Private Sub Text11_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text11_LostFocus(Index)
    End If
End Sub

Private Sub Text11_LostFocus(Index As Integer)
    Text11(Index) = TRIM$(FORMATNUM$(XVALO(Text11(Index)), "F10.1"))
End Sub

Sub MUESTRASEL(Index)
   '
   CIXI% = CODINT%(TRIM$(Label27))
   IDLOT$ = TRIM$(Label23(Index))
   NUMDEP% = XVALO(Label55(Index))
   CANSA = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
   ANCHX = XVALO(Label12(Index).Caption) / 1000
   LARGX = 0: If ANCHX > 0 Then LARGX = CANSA / ANCHX
   Label25(Index) = FORMATNUM$(LARGX, "F10.1")
      IXX% = Index
   Call Text11_Change(IXX%)
   '
End Sub

Sub MUESTRASEL1(Index)
   '
   If Index < 4 Then
      CIXI% = CODINT%(TRIM$(Label39(Index)))
      IDLOT$ = TRIM$(Label40(Index))
      NUMDEP% = XVALO(Label52(Index))
      CANSA = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
      Label41(Index) = FORMATNUM$(CANSA, "F10.1")
   End If
   '
End Sub

Private Sub VScroll1_Change()
   If VScroll1.Value < 360 Then
    Picture5.Top = (-10) * VScroll1.Value
   End If
End Sub

Private Sub VScroll1_Scroll()
   If VScroll1.Value < 360 Then
    Picture5.Top = (-10) * VScroll1.Value
   End If
End Sub
