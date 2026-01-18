VERSION 5.00
Begin VB.Form CDESA04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuestación - Desarrollos Especiales"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "CDESA04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   9
      Left            =   105
      TabIndex        =   153
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   9
         Left            =   1365
         TabIndex        =   160
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   9
         Left            =   7020
         TabIndex        =   159
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   9
         Left            =   5730
         TabIndex        =   158
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   9
         Left            =   4560
         TabIndex        =   157
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   9
         Left            =   3375
         TabIndex        =   156
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   9
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   155
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   9
         Left            =   1365
         TabIndex        =   154
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   9
         Left            =   210
         TabIndex        =   162
         Top             =   980
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   9
         Left            =   210
         TabIndex        =   161
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   8
      Left            =   105
      TabIndex        =   143
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   8
         Left            =   1365
         TabIndex        =   150
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   8
         Left            =   7020
         TabIndex        =   149
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   8
         Left            =   5730
         TabIndex        =   148
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   8
         Left            =   4560
         TabIndex        =   147
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   8
         Left            =   3375
         TabIndex        =   146
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   8
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   145
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   8
         Left            =   1365
         TabIndex        =   144
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   8
         Left            =   210
         TabIndex        =   152
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   8
         Left            =   210
         TabIndex        =   151
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   7
      Left            =   105
      TabIndex        =   133
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   7
         Left            =   1365
         TabIndex        =   140
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   7
         Left            =   7020
         TabIndex        =   139
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   7
         Left            =   5730
         TabIndex        =   138
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   7
         Left            =   4560
         TabIndex        =   137
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   7
         Left            =   3375
         TabIndex        =   136
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   7
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   135
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   7
         Left            =   1365
         TabIndex        =   134
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   7
         Left            =   210
         TabIndex        =   142
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   7
         Left            =   210
         TabIndex        =   141
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   6
      Left            =   105
      TabIndex        =   123
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   6
         Left            =   1365
         TabIndex        =   130
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   6
         Left            =   7020
         TabIndex        =   129
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   6
         Left            =   5730
         TabIndex        =   128
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   6
         Left            =   4560
         TabIndex        =   127
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   6
         Left            =   3375
         TabIndex        =   126
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   6
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   125
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   6
         Left            =   1365
         TabIndex        =   124
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   6
         Left            =   210
         TabIndex        =   132
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   6
         Left            =   210
         TabIndex        =   131
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   5
      Left            =   105
      TabIndex        =   113
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   5
         Left            =   1365
         TabIndex        =   120
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   5
         Left            =   7020
         TabIndex        =   119
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   5
         Left            =   5730
         TabIndex        =   118
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   5
         Left            =   4560
         TabIndex        =   117
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   5
         Left            =   3375
         TabIndex        =   116
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   5
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   115
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   5
         Left            =   1365
         TabIndex        =   114
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   5
         Left            =   210
         TabIndex        =   122
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   5
         Left            =   210
         TabIndex        =   121
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   4
      Left            =   105
      TabIndex        =   103
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   4
         Left            =   1365
         TabIndex        =   110
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   4
         Left            =   7020
         TabIndex        =   109
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   4
         Left            =   5730
         TabIndex        =   108
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   4
         Left            =   4560
         TabIndex        =   107
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   4
         Left            =   3375
         TabIndex        =   106
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   4
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   4
         Left            =   1365
         TabIndex        =   104
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   4
         Left            =   210
         TabIndex        =   112
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   4
         Left            =   210
         TabIndex        =   111
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   3
      Left            =   105
      TabIndex        =   93
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   3
         Left            =   1365
         TabIndex        =   100
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   3
         Left            =   7020
         TabIndex        =   99
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   3
         Left            =   5730
         TabIndex        =   98
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   3
         Left            =   4560
         TabIndex        =   97
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   3
         Left            =   3375
         TabIndex        =   96
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   3
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   95
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   1365
         TabIndex        =   94
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   3
         Left            =   210
         TabIndex        =   102
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   3
         Left            =   210
         TabIndex        =   101
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   2
      Left            =   105
      TabIndex        =   83
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   2
         Left            =   1365
         TabIndex        =   90
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   2
         Left            =   7020
         TabIndex        =   89
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   2
         Left            =   5730
         TabIndex        =   88
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   2
         Left            =   4560
         TabIndex        =   87
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   2
         Left            =   3375
         TabIndex        =   86
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   2
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   2
         Left            =   1365
         TabIndex        =   84
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   2
         Left            =   210
         TabIndex        =   92
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   2
         Left            =   210
         TabIndex        =   91
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   1
      Left            =   105
      TabIndex        =   73
      Top             =   1680
      Width           =   8520
      Begin VB.TextBox Text2 
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
         Index           =   1
         Left            =   1365
         TabIndex        =   80
         Top             =   420
         Width           =   6945
      End
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
         Height          =   3645
         Index           =   1
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   945
         Width           =   6945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   1
         Left            =   3375
         TabIndex        =   78
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   1
         Left            =   4560
         TabIndex        =   77
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   1
         Left            =   5730
         TabIndex        =   76
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   1
         Left            =   7020
         TabIndex        =   75
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   1
         Left            =   1365
         TabIndex        =   74
         Top             =   4590
         Width           =   1995
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Left            =   210
         TabIndex        =   82
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Left            =   210
         TabIndex        =   81
         Top             =   945
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Importes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   8715
      TabIndex        =   33
      Top             =   1680
      Width           =   2270
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   9
         Left            =   105
         TabIndex        =   62
         Top             =   4560
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   735
         TabIndex        =   61
         Top             =   4515
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   8
         Left            =   105
         TabIndex        =   59
         Top             =   4140
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   735
         TabIndex        =   58
         Top             =   4095
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   7
         Left            =   105
         TabIndex        =   56
         Top             =   3720
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   735
         TabIndex        =   55
         Top             =   3675
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   53
         Top             =   3300
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   735
         TabIndex        =   52
         Top             =   3255
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   50
         Top             =   2880
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   735
         TabIndex        =   49
         Top             =   2835
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   47
         Top             =   2460
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   735
         TabIndex        =   46
         Top             =   2415
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   44
         Top             =   2040
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   735
         TabIndex        =   43
         Top             =   1995
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   41
         Top             =   1620
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   735
         TabIndex        =   40
         Top             =   1575
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   38
         Top             =   1200
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   735
         TabIndex        =   37
         Top             =   1155
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0E0FF&
         Caption         =   "Option1"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   35
         Top             =   780
         Width           =   225
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   735
         TabIndex        =   34
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "9)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   460
         TabIndex        =   63
         Top             =   4530
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "8)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   460
         TabIndex        =   60
         Top             =   4110
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "7)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   460
         TabIndex        =   57
         Top             =   3690
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "6)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   460
         TabIndex        =   54
         Top             =   3270
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "5)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   460
         TabIndex        =   51
         Top             =   2850
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "4)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   460
         TabIndex        =   48
         Top             =   2430
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   460
         TabIndex        =   45
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   460
         TabIndex        =   42
         Top             =   1590
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   460
         TabIndex        =   39
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H00D0E0FF&
         Caption         =   "0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   460
         TabIndex        =   36
         Top             =   750
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
      Caption         =   "Especificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   0
      Left            =   105
      TabIndex        =   32
      Top             =   1680
      Width           =   8520
      Begin VB.CommandButton Command1 
         Height          =   260
         Index           =   0
         Left            =   1365
         TabIndex        =   72
         Top             =   4590
         Width           =   1995
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Index           =   0
         Left            =   7020
         TabIndex        =   71
         Top             =   4590
         Width           =   1290
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Pegar"
         Height          =   260
         Index           =   0
         Left            =   5730
         TabIndex        =   70
         Top             =   4590
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   260
         Index           =   0
         Left            =   4560
         TabIndex        =   69
         Top             =   4590
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cortar"
         Height          =   260
         Index           =   0
         Left            =   3375
         TabIndex        =   68
         Top             =   4590
         Width           =   1170
      End
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
         Height          =   3645
         Index           =   0
         Left            =   1365
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Top             =   945
         Width           =   6945
      End
      Begin VB.TextBox Text2 
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
         Index           =   0
         Left            =   1365
         TabIndex        =   66
         Top             =   420
         Width           =   6945
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Detalle:"
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
         Index           =   0
         Left            =   210
         TabIndex        =   65
         Top             =   945
         Width           =   1065
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0E0FF&
         Caption         =   "Referencia:"
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
         Index           =   0
         Left            =   210
         TabIndex        =   64
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.TextBox Text75 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5115
      TabIndex        =   30
      Top             =   1155
      Width           =   3300
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5115
      TabIndex        =   25
      Top             =   840
      Width           =   870
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5115
      TabIndex        =   24
      Top             =   210
      Width           =   3300
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5115
      TabIndex        =   23
      Top             =   525
      Width           =   3300
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6060
      TabIndex        =   22
      Top             =   840
      Width           =   2340
   End
   Begin VB.TextBox Text38 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9200
      TabIndex        =   20
      Top             =   7035
      Width           =   1650
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
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
      Left            =   8730
      TabIndex        =   18
      Top             =   1155
      Width           =   2220
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9820
      TabIndex        =   10
      Top             =   210
      Width           =   1125
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9975
      TabIndex        =   9
      Top             =   8595
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9820
      TabIndex        =   8
      Top             =   525
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   7845
      Left            =   11025
      ScaleHeight     =   7785
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command15 
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
         Left            =   105
         Picture         =   "CDESA04.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Guardar "
         Top             =   2040
         Width           =   640
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
         Height          =   690
         Left            =   120
         Picture         =   "CDESA04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   5600
         Width           =   650
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
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
         Left            =   105
         Picture         =   "CDESA04.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprime Cotización de Cámara"
         Top             =   2730
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Left            =   105
         Picture         =   "CDESA04.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copia Interactiva de Rubro"
         Top             =   3405
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   6645
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CDESA04.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command33 
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
         Left            =   105
         Picture         =   "CDESA04.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command44 
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
         Left            =   105
         Picture         =   "CDESA04.frx":1796
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CDESA04.frx":1AA0
         Stretch         =   -1  'True
         Top             =   6885
         Width           =   1410
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "At.Sr:"
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
      Left            =   4150
      TabIndex        =   31
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   3960
      TabIndex        =   28
      Top             =   305
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   4170
      TabIndex        =   27
      Top             =   935
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   3990
      TabIndex        =   26
      Top             =   620
      Width           =   1110
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Desarrollos Especiales:"
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
      Left            =   6315
      TabIndex        =   21
      Top             =   7140
      Width           =   2745
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11025
      Y1              =   6825
      Y2              =   6825
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable.:"
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
      Left            =   8505
      TabIndex        =   19
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   1155
      TabIndex        =   17
      Top             =   525
      Width           =   2910
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1155
      TabIndex        =   16
      Top             =   210
      Width           =   2160
   End
   Begin VB.Label Label75 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. de Presupuesto:"
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
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Denom.:"
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
      Left            =   -210
      TabIndex        =   14
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
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
      Height          =   330
      Left            =   9045
      TabIndex        =   13
      Top             =   305
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Validez:"
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
      Left            =   8925
      TabIndex        =   12
      Top             =   620
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   9450
      TabIndex        =   11
      Top             =   8700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "CDESA04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim ARCHIPRE$


Private Sub AyuCli_Click()
  Call Command33_Click
End Sub

Private Sub CmdLiPre_Click()
    '
    If TRIM$(Label31) = "" Then
       On Error Resume Next
       Call MENSERR(24, "Imposible Continuar.\Número de Presupuesto no válido")
       On Error GoTo 0
       Exit Sub
    End If
    '
    ARCHEC$ = "LISTAS": LOCA% = 2
    ARGU$ = CBIN$(TRIM$(TEXT21.TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    VDEVU$ = TRIM$(VALACT1$(ARCHEC$))
    If VDEVU$ <> "" Then
       TEXT21.TEXT = VDEVU$
    End If
    '
End Sub

Private Sub CoInLiMa_Click()
  Call Command22_Click
End Sub

Private Sub Command14_Click()
  Command14.Enabled = False
  OPPRFLE04.Show 1
  Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   Call GRADESPE
   Command15.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
  COMIE = Text3(Index).SelStart
  LOSEL = Text3(Index).SelLength
  Clipboard.SetText Mid$(Text3(Index).TEXT, Text3(Index).SelStart + 1, Text3(Index).SelLength)
  Text3(Index).TEXT = Left$(Text3(Index).TEXT, COMIE) + Mid$(Text3(Index).TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command22_Click()
  Command22.Enabled = False
  Command22.Enabled = True
End Sub

Private Sub Command444_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command33_Click
End Sub

Private Sub EditCli_Click()
   'Call Command9_Click
End Sub

Private Sub ConGenCli_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub Exit_Click()
   Call Command33_Click
End Sub

Private Sub Command3_Click(Index As Integer)
  Clipboard.SetText Mid$(Text3(Index).TEXT, Text3(Index).SelStart + 1, Text3(Index).SelLength)
End Sub

Private Sub Command33_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command4_Click(Index As Integer)
   COMIE = Text3(Index).SelStart
   LOSEL = Text3(Index).SelLength
   Text3(Index).TEXT = Left$(Text3(Index).TEXT, COMIE) + Clipboard.GetText + Mid$(Text3(Index).TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command5_Click(Index As Integer)
   Text3(Index).SelStart = 0
   Text3(Index).SelLength = Len(Text3(Index).TEXT)
   Text3(Index).SetFocus
End Sub

'
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)

End Sub
'
Private Sub Form_Activate()
  WindowState = 0
  OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
  If WindowState = 1 Then
    Call CIERRARCH("*.*")
    Call BAJALDISCO
  End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Command333_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub



Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Label31_Change()
   '
   NUOTR$ = Label31
   '
   With REPRFLEX04
     Label7 = .Text3
     Text9 = .Text9
     Text10 = .Text10
     Text8 = .Text11
     Text18 = .Text18
     Text5 = .Text5
     Text7 = .Text7
     Text4 = .Text8
     Text75 = .Text2
   End With
   '
   For KKI% = 0 To 9
     Text2(KKI%) = ""
     Text3(KKI%) = ""
     Text1(KKI%) = ""
   Next KKI%
   Option1(0).Value = True
   '
   Call ABREPRECAM
   SUDESPE# = 0
   With Despecial
     .FindFirst "  NRO_Presu= '" & NUOTR$ & "' "
10   If .NoMatch = False Then
       IKI% = !Num_Orden
       If IKI% >= 0 Then
         If IKI% <= 9 Then
           On Error Resume Next
           Text2(IKI%) = !Referen
           Text3(IKI%) = !Descrip
           Text1(IKI%) = FORMATNUM$(!IMPORTE, "F11.2")
           On Error GoTo 0
           SUDESPE# = SUDESPE# + !IMPORTE
         End If
       End If
       .FindNext "  NRO_Presu= '" & NUOTR$ & "' "
       GoTo 10
     End If
   End With
   '
End Sub

Private Sub Option1_Click(Index As Integer)
  '
  Call RESALAB(Index) 'DES*RESALTA TODOS - RESALTA ELEGIDO
  Call OCUFRAM(Index) 'OCULTA TODOS - ACTIVA  ELEGIDO
  '
End Sub
'
Sub RESALAB(ELEGI1%)
  '
  For I% = 0 To Label13.Count - 1
      Label13(I%).FontSize = 12
      Label13(I%).ForeColor = &H80000012
      Text1(I%).Enabled = False
  Next I%
  Label13(ELEGI1%).ForeColor = &HC00000
  Label13(ELEGI1%).FontSize = 13
  Text1(ELEGI1%).Enabled = True
  '
End Sub

Sub OCUFRAM(EXECP1%)
  '
  For I% = 0 To Frame1.Count - 1
      Frame1(I%).Visible = False
  Next I%
  Frame1(EXECP1%).Visible = True 'DESOCULTA EL ELEGIDO
  '
End Sub

Private Sub Text1_Change(Index As Integer)
  '
  TDE1 = 0
  For I% = 0 To Text1.Count - 1
    TDE1 = TDE1 + Val(Text1(I%).TEXT)
  Next I%
  Text38.TEXT = FORMATNUM$(TDE1, "F12.2")
  '
End Sub

Private Sub Text1_LostFocus(Index As Integer)
  '
  TOTXT1 = Val(Text1(Index))
  Text1(Index).TEXT = FORMATNUM$(TOTXT1, "F11.2")
  '
End Sub

Private Sub Text75_GotFocus()
  '
  Command333.SetFocus
  '
End Sub


Private Sub Text10_GOTFOCUS()
  '
  Command333.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command333.SetFocus
  '
End Sub

Private Sub Text13_Change()
  REPRFLEX04.Label42 = Text13
  REPRFLEX04.Label43 = FORMATNUM(Val(Text38) - Val(Text13), "F11.2")
End Sub

Private Sub Text13_GOTFOCUS()
  '
  Command333.SetFocus
  '
End Sub
Private Sub Text18_GOTFOCUS()
  '
  Command33.SetFocus
  '
End Sub






Sub BLANFORMU()
   '

   Text24 = ""

  Text4.TEXT = ""

  Text19.TEXT = ""

     Text23 = ""
     Text35 = ""
     Text36 = ""
     Text37 = ""
   
     
     Text18 = ""
     Text5 = ""
     Text7 = ""
     Text8 = ""
     '
     Text9 = ""
     Text10 = ""
     Text25 = ""
     Text24 = ""
     Text11 = ""
     Text13 = ""
     Text12 = ""
     Text17 = ""

   List3.Clear

   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Unload Me
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub



Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub






Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub



Private Sub text38_gotfocus()
  '
  Command33.SetFocus
  '
End Sub
Private Sub text38_change()
  REPRFLEX04.Label44 = FORMATNUM(Val(Text38) - Val(Text13), "F11.2")
End Sub

Private Sub Text4_GOTFOCUS()
  '
  Command33.SetFocus
  '
End Sub

Private Sub Text5_GOTFOCUS()
  '
  Command33.SetFocus
  '
End Sub

Private Sub Text7_GOTFOCUS()
  '
    Command33.SetFocus
  '
End Sub

Private Sub Text8_GOTFOCUS()
  '
  Command33.SetFocus
  '
End Sub

Private Sub Text9_GOTFOCUS()
  '
  Command33.SetFocus
  '
End Sub

Private Sub TICAMBIS_CLICK()
  '
  Command33.SetFocus
  '
End Sub
Sub COEQUIP()
  '
    '\\\OT-05-0393-WAR-10/06/05///
  Command22.Enabled = False
  If DERACCE%("COPYRUB", "Generar Rubros por Copia") < 2 Then
    GoTo 99
  End If
  '
9 Screen.MousePointer = 1
  '
  VDEF$ = Space$(64) + String$(6, 0)
  '
  OBX$ = OBJPLA$("DEACOPRE", VDEF$)
  If OBX$ = "" Then GoTo 99
  VDEF$ = OBX$
  '
  PRESU1$ = TRIM$(Mid$(OBX$, 1, 16))
  PRESU2$ = TRIM$(Mid$(OBX$, 17, 48))
  '//////////////VALIDACIONES////////////
  If TRIM$(PRESU1$) = "" Then
    Call COMUNI("Número de Presupuesto no valido")
    GoTo 9
  End If
  '
  If TRIM$(PRESU2$) = "" Then
    Call COMUNI("Número de Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU1$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU2$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If TRIM$(PRESU1$) = TRIM$(PRESU2$) Then
    Call MENSERR(24, "Copia Imposible.\Coinciden Presupuesto Origen y Destino.")
    GoTo 99
  End If
  'validación de que en el presupuesto
  'no haya detalle con este rubro.
  NRUB% = 4 'EQUIPO
  Call ABREPRECAM
  SQLX$ = " Nro_Presu= '" & PRESU2$ & "' "
  SQLX$ = SQLX$ + " AND Rubro_Det = " & NRUB% & " "
  DetaPresu.FindFirst SQLX$
  If DetaPresu.NoMatch = False Then
      If COMALTER%("El Presupuesto de Cámara " & PRESU2$ & " \ya Contiene Detalle en Rubro.\¿Desea Sobre-Escribirlo ?\\Si, Sobre-escribir.\\No, Cancelar.") = 2 Then
          On Error GoTo 0
          Screen.MousePointer = 1
          Command22.Enabled = True
          Exit Sub
        Else
13        DetaPresu.FindFirst SQLX$
          If DetaPresu.NoMatch = False Then
            DetaPresu.Delete
            GoTo 13
          End If
      End If
    End If
    '
17  SQLX$ = " SELECT * FROM detapresu "
    SQLX$ = SQLX$ + " WHERE Nro_Presu= '" & PRESU1$ & "' "
    SQLX$ = SQLX$ + " AND Rubro_Det = " & NRUB% & ""
    Set DETAPRESUTMP = PresuCam.OpenRecordset(SQLX$, 4)
    '
    With DETAPRESUTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
     Else
       'VALIDA LA MONEDA EN ENCABEZADO DE PRESUP. DESTINO
        With EncaPresu
          .FindFirst " Nro_Presu = '" & PRESU1$ & "' "
          MONEDET% = !Mone_Presu
          .FindFirst " Nro_Presu = '" & PRESU2$ & "' "
          MONDEST% = !Mone_Presu
        End With
          TC1# = 1
          TC2# = 1
        If MONDEST% <> MONEDET% Then
          TC1# = TICAMONE(MONEDET%)
          TC2# = TICAMONE(MONDEST%)
        End If
        '
       Do While .EOF = False
        NURUB% = !rubro_det
        SUBRUB% = !SUBRUB_DET
        CODET% = !CODINTE_DET
        CODEXTET$ = !CODEXTE_DET
        CANMAT# = !Cantid_Det
        PREUMAT# = !PreUni_Det
        MED1 = !MED1_DET
        MED2 = !Med2_Det
        MED3 = !Med3_Det
        IMPODEP# = !Impo_Det
        DESCRIDET$ = !DESCRI_DET
        HORADET = !Hora_Det
        OPERADET% = !Opera_Det
        VALOHORDET# = !ValoHor_Det
        OBSGRU$ = !OBSERVA_GRUP
        With DetaPresu
          .AddNew
            !Nro_PRESU = PRESU2$
            !rubro_det = NURUB%
            !SUBRUB_DET = SUBRUB%
            !CODINTE_DET = CODET%
            !CODEXTE_DET = CODEXTET$
            !Cantid_Det = CANMAT#
            !PreUni_Det = PREUMAT# * TC1# / TC2#
            !MONE_DET = MONEDET%
            !MED1_DET = MED1
            !Med2_Det = MED2
            !Med3_Det = MED3
            !Impo_Det = (PREUMAT# * TC1# / TC2#) * CANMAT#
            !DESCRI_DET = DESCRIDET$
            !Hora_Det = HORADET
            !Opera_Det = OPERADET%
            !ValoHor_Det = VALOHORDET#
            !OBSERVA_GRUP = OBSGRU$
          .Update
        End With
        .MoveNext
       Loop
       'COPIA EXITOSA
       FORPRE04.Text1 = PRESU2$
       Label31 = PRESU2$
       Call Form_Activate
       '
       End If
    End With
    '
99  Screen.MousePointer = 1
    Command22.Enabled = True
  '\\\OT-05-0393-WAR-FIN///
End Sub

Sub GRADESPE()
   '
   If DERACCE%("GRADATPR", "Grabar Datos de Presupuesto") < 2 Then
       Exit Sub
   End If
   '
   'VALIDACIONES
   If TRIM$(Label31) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Label31) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   NUOTR$ = TRIM$(Label31)
   If NUOTR$ = "" Then Exit Sub
   '
   Call ABREPRECAM
   SUDESPE# = 0
   With Despecial
10   .FindFirst "  NRO_Presu= '" & NUOTR$ & "' "
     If .NoMatch = False Then
       .Delete
       GoTo 10
     End If
     '
     For KKI% = 0 To 9
       REFEX3$ = Text2(KKI%)
       DESCR$ = Text3(KKI%)
       IMPOR# = Val(Text1(KKI%))
       '
       If TRIM$(REFEX3$) <> "" Or TRIM$(DESCR$) <> "" Then
         .AddNew
         !Nro_PRESU = NUOTR$
         !Num_Orden = KKI%
         !Referen = REFEX3$
         !Descrip = DESCR$
         !IMPORTE = IMPOR#
         .Update
       End If
       '
     Next KKI%
   End With
   '
End Sub
