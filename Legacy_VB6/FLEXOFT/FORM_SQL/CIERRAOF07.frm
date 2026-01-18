VERSION 5.00
Begin VB.Form CIERRAOF07 
   BackColor       =   &H00FFE0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Fabricacion"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
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
   ScaleHeight     =   8565
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   495
      Left            =   4440
      TabIndex        =   183
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3300
      TabIndex        =   1
      ToolTipText     =   "Incluye O.Fabricación Cerradas"
      Top             =   210
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Solo Devolución"
      Height          =   435
      Left            =   1360
      TabIndex        =   5
      Top             =   900
      Width           =   1320
   End
   Begin VB.CommandButton Command111 
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
      Left            =   1845
      TabIndex        =   64
      Top             =   1400
      Width           =   175
   End
   Begin VB.TextBox Text55 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1365
      TabIndex        =   63
      Top             =   1400
      Width           =   500
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      TabIndex        =   6
      Top             =   1025
      Width           =   2620
   End
   Begin VB.TextBox Text6 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4080
      TabIndex        =   3
      Top             =   580
      Width           =   540
   End
   Begin VB.CommandButton Command9 
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
      Left            =   4605
      TabIndex        =   4
      Top             =   580
      Width           =   175
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Devolución de Materiales No Trazables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   137
      Top             =   6480
      Width           =   8730
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   7
         Left            =   9000
         MaxLength       =   32
         TabIndex        =   181
         Top             =   1545
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   6
         Left            =   9000
         MaxLength       =   32
         TabIndex        =   180
         Top             =   1230
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   5
         Left            =   9000
         MaxLength       =   32
         TabIndex        =   179
         Top             =   915
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   4
         Left            =   9000
         MaxLength       =   32
         TabIndex        =   178
         Top             =   600
         Width           =   800
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
         Index           =   7
         Left            =   6405
         TabIndex        =   53
         Top             =   1590
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
         TabIndex        =   50
         Top             =   1275
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
         TabIndex        =   47
         Top             =   960
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
         TabIndex        =   44
         Top             =   645
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
         TabIndex        =   43
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   7455
         TabIndex        =   45
         Top             =   630
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
         TabIndex        =   46
         Top             =   945
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   7455
         TabIndex        =   48
         Top             =   945
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
         TabIndex        =   49
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   7455
         TabIndex        =   51
         Top             =   1260
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
         TabIndex        =   52
         Top             =   1575
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   7455
         TabIndex        =   54
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   6090
         TabIndex        =   164
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1575
         Width           =   330
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   6090
         TabIndex        =   163
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   6090
         TabIndex        =   162
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   945
         Width           =   330
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         Height          =   195
         Left            =   4305
         TabIndex        =   161
         Top             =   420
         Width           =   1965
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   6090
         TabIndex        =   160
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   2625
         TabIndex        =   138
         Top             =   630
         Width           =   3300
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   525
         TabIndex        =   153
         ToolTipText     =   "Código"
         Top             =   630
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   6720
         TabIndex        =   152
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   300
         Left            =   6930
         TabIndex        =   151
         Top             =   405
         Width           =   1125
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         Height          =   405
         Left            =   -1155
         TabIndex        =   150
         Top             =   405
         Width           =   1965
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         Height          =   300
         Left            =   6300
         TabIndex        =   149
         Top             =   405
         Width           =   705
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   525
         TabIndex        =   148
         ToolTipText     =   "Código"
         Top             =   945
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   6720
         TabIndex        =   147
         Top             =   945
         Width           =   420
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   2625
         TabIndex        =   146
         Top             =   945
         Width           =   3300
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   525
         TabIndex        =   145
         ToolTipText     =   "Código"
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   6720
         TabIndex        =   144
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   2625
         TabIndex        =   143
         Top             =   1260
         Width           =   3300
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   525
         TabIndex        =   142
         ToolTipText     =   "Código"
         Top             =   1575
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   6720
         TabIndex        =   141
         Top             =   1575
         Width           =   420
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   2625
         TabIndex        =   140
         Top             =   1575
         Width           =   3300
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   225
         Left            =   2415
         TabIndex        =   139
         Top             =   420
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Devolucion de Papeles"
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
      Left            =   240
      TabIndex        =   86
      Top             =   2160
      Width           =   8730
      Begin VB.TextBox Text60 
         Height          =   250
         Index           =   2
         Left            =   3360
         TabIndex        =   21
         Top             =   1800
         Width           =   980
      End
      Begin VB.TextBox Text60 
         Height          =   250
         Index           =   1
         Left            =   3360
         TabIndex        =   18
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
         Top             =   1530
         Width           =   980
      End
      Begin VB.TextBox Text60 
         Height          =   250
         Index           =   0
         Left            =   3360
         TabIndex        =   14
         Top             =   1260
         Width           =   980
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
         Index           =   2
         Left            =   3135
         TabIndex        =   20
         Top             =   1800
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
         Left            =   3135
         TabIndex        =   16
         Top             =   1530
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
         Left            =   3135
         TabIndex        =   13
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   250
         Index           =   0
         Left            =   6330
         TabIndex        =   15
         Top             =   1260
         Width           =   1065
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   250
         Index           =   1
         Left            =   6330
         TabIndex        =   8
         ToolTipText     =   "Doble-Click para Cancelar Reserva Trazada"
         Top             =   1530
         Width           =   1065
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Height          =   250
         Index           =   2
         Left            =   6330
         TabIndex        =   22
         Top             =   1800
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
         Left            =   2520
         TabIndex        =   12
         Top             =   1260
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
         Left            =   2520
         TabIndex        =   17
         Top             =   1530
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
         Left            =   2520
         TabIndex        =   19
         Top             =   1800
         Width           =   175
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
         TabIndex        =   11
         Top             =   315
         Width           =   175
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubic.:"
         Height          =   405
         Left            =   3260
         TabIndex        =   177
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         Height          =   195
         Left            =   1260
         TabIndex        =   168
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2820
         TabIndex        =   167
         Top             =   1800
         Width           =   330
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   166
         Top             =   1530
         Width           =   330
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2820
         TabIndex        =   165
         Top             =   1260
         Width           =   330
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4725
         TabIndex        =   154
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   405
         Left            =   6300
         TabIndex        =   109
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   7455
         TabIndex        =   108
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   5260
         TabIndex        =   107
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   4395
         TabIndex        =   106
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   7455
         TabIndex        =   105
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   5260
         TabIndex        =   104
         Top             =   1530
         Width           =   1020
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   4395
         TabIndex        =   103
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   7455
         TabIndex        =   102
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   5260
         TabIndex        =   101
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   4395
         TabIndex        =   100
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   99
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N.C.Rec.:"
         Height          =   405
         Left            =   -840
         TabIndex        =   98
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ancho:"
         Height          =   405
         Left            =   3050
         TabIndex        =   97
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Largo:"
         Height          =   405
         Left            =   3850
         TabIndex        =   96
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
         Height          =   405
         Left            =   5050
         TabIndex        =   95
         Top             =   1050
         Width           =   1965
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   94
         Top             =   1530
         Width           =   1755
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   93
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   300
         Left            =   -525
         TabIndex        =   92
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   91
         ToolTipText     =   "Código"
         Top             =   315
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   840
         TabIndex        =   90
         ToolTipText     =   "Descripción"
         Top             =   630
         Width           =   7680
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Metros:"
         Height          =   300
         Left            =   2940
         TabIndex        =   89
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tot.Cantidad:"
         Height          =   300
         Left            =   6090
         TabIndex        =   88
         Top             =   390
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7455
         TabIndex        =   87
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Cantidad"
      Height          =   1980
      Left            =   6810
      TabIndex        =   79
      Top             =   120
      Width           =   2130
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   850
         TabIndex        =   10
         Top             =   1500
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   850
         TabIndex        =   9
         Top             =   1180
         Width           =   1200
      End
      Begin VB.TextBox Tlabel38 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         TabIndex        =   7
         Top             =   865
         Width           =   1200
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R/Cajas:"
         Height          =   300
         Left            =   -525
         TabIndex        =   173
         Top             =   1575
         Width           =   1365
      End
      Begin VB.Label Label151 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Et/Rollo:"
         Height          =   300
         Left            =   -525
         TabIndex        =   172
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   855
         TabIndex        =   84
         Top             =   540
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   855
         TabIndex        =   83
         Top             =   215
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
         Height          =   300
         Left            =   -525
         TabIndex        =   82
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyect:"
         Height          =   300
         Left            =   -525
         TabIndex        =   81
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acumul:"
         Height          =   300
         Left            =   -525
         TabIndex        =   80
         Top             =   610
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
      TabIndex        =   2
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
      Left            =   3120
      TabIndex        =   0
      Top             =   210
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
      Height          =   8500
      Left            =   9135
      ScaleHeight     =   8445
      ScaleWidth      =   1320
      TabIndex        =   65
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command15 
         Caption         =   "Recup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   182
         ToolTipText     =   "Recuperacion Etiquetas x Cajas a partir de la imagen de la Etiqueta"
         Top             =   1740
         Width           =   645
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
         Height          =   690
         Left            =   120
         Picture         =   "CIERRAOF07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Reimpresión de Etiquetas/Modificación de Causas"
         Top             =   3120
         Width           =   650
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Print"
         Height          =   620
         Left            =   120
         Picture         =   "CIERRAOF07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Reimpresón de Etiquetas"
         Top             =   5400
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         Height          =   650
         Left            =   120
         Picture         =   "CIERRAOF07.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   4730
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   1005
         Left            =   120
         Picture         =   "CIERRAOF07.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Guardar Cierre de O/F y Reingreso de Materiales a O/F"
         Top             =   6360
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
         Height          =   660
         Left            =   120
         Picture         =   "CIERRAOF07.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   4035
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Help"
         Height          =   690
         Left            =   120
         Picture         =   "CIERRAOF07.frx":13CA
         Style           =   1  'Graphical
         TabIndex        =   61
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
         Picture         =   "CIERRAOF07.frx":16D4
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Anular Vale de Materiales"
         Top             =   2205
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
         TabIndex        =   66
         Top             =   7455
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
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   120
         Picture         =   "CIERRAOF07.frx":19DE
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CIERRAOF07.frx":1B28
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7770
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Devolución de Tintas e Insumos Trazables"
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
      TabIndex        =   68
      Top             =   4440
      Width           =   8730
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   0
         Left            =   5040
         MaxLength       =   32
         TabIndex        =   26
         Top             =   420
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   1
         Left            =   5040
         MaxLength       =   32
         TabIndex        =   31
         Top             =   735
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   2
         Left            =   5040
         MaxLength       =   32
         TabIndex        =   36
         Top             =   1050
         Width           =   800
      End
      Begin VB.TextBox Text61 
         Height          =   285
         Index           =   3
         Left            =   5040
         MaxLength       =   32
         TabIndex        =   41
         Top             =   1365
         Width           =   800
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
         Index           =   3
         Left            =   4800
         TabIndex        =   40
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
         Left            =   4800
         TabIndex        =   35
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
         Left            =   4800
         TabIndex        =   30
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
         Left            =   4800
         TabIndex        =   25
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7455
         TabIndex        =   42
         Top             =   1320
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
         Left            =   4290
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   1365
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   7455
         TabIndex        =   37
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
         Left            =   4290
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   1050
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   7455
         TabIndex        =   32
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
         Left            =   4290
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   735
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   7455
         TabIndex        =   27
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
         Left            =   4290
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubic.:"
         Height          =   405
         Left            =   4950
         TabIndex        =   176
         Top             =   195
         Width           =   645
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   4485
         TabIndex        =   159
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   4485
         TabIndex        =   158
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   4485
         TabIndex        =   157
         Top             =   735
         Width           =   330
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         Height          =   195
         Left            =   2940
         TabIndex        =   156
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   4485
         TabIndex        =   155
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7455
         TabIndex        =   136
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total: ..................."
         Height          =   300
         Left            =   5145
         TabIndex        =   135
         Top             =   1755
         Width           =   2280
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   2715
         TabIndex        =   134
         Top             =   1365
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2715
         TabIndex        =   133
         Top             =   1050
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2670
         TabIndex        =   132
         Top             =   735
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2715
         TabIndex        =   131
         Top             =   420
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   6980
         TabIndex        =   130
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   5900
         TabIndex        =   129
         Top             =   1365
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   2925
         TabIndex        =   128
         Top             =   1365
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   127
         ToolTipText     =   "Código"
         Top             =   1365
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   6980
         TabIndex        =   126
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   5900
         TabIndex        =   125
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2925
         TabIndex        =   124
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   123
         ToolTipText     =   "Código"
         Top             =   1050
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   6980
         TabIndex        =   122
         Top             =   735
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   5900
         TabIndex        =   121
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2925
         TabIndex        =   120
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   119
         ToolTipText     =   "Código"
         Top             =   735
         Width           =   1755
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         Height          =   405
         Left            =   6380
         TabIndex        =   118
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   405
         Left            =   3660
         TabIndex        =   117
         Top             =   195
         Width           =   2805
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         Height          =   405
         Left            =   1365
         TabIndex        =   116
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
         TabIndex        =   115
         Top             =   450
         Width           =   1965
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   405
         Left            =   6350
         TabIndex        =   114
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   6980
         TabIndex        =   113
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   5900
         TabIndex        =   112
         Top             =   420
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2925
         TabIndex        =   111
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   110
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   1755
      End
   End
   Begin VB.Label Label59 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Causa:"
      Height          =   300
      Left            =   120
      TabIndex        =   175
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label Label51 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   174
      Top             =   1395
      Width           =   4680
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicación en el Deposito:"
      Height          =   510
      Left            =   2360
      TabIndex        =   171
      Top             =   900
      Width           =   1725
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      Height          =   270
      Left            =   3495
      TabIndex        =   170
      Top             =   385
      Width           =   915
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   169
      Top             =   580
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
      TabIndex        =   85
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   78
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   210
      Width           =   1800
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      Height          =   300
      Left            =   105
      TabIndex        =   77
      Top             =   1825
      Width           =   1125
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   76
      Top             =   1785
      Width           =   5370
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   75
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   300
      Left            =   -315
      TabIndex        =   74
      Top             =   550
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
      Height          =   510
      Left            =   4200
      TabIndex        =   73
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5565
      TabIndex        =   72
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Fabr. Nº:"
      Height          =   300
      Left            =   105
      TabIndex        =   71
      Top             =   270
      Width           =   1125
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
      Height          =   300
      Left            =   1470
      TabIndex        =   70
      Top             =   4830
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1470
      TabIndex        =   69
      Top             =   4830
      Width           =   1335
   End
End
Attribute VB_Name = "CIERRAOF07"
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

Private Sub Command111_Click()
    '
    Command111.Enabled = False
    Call SELECHO("CAUINGRE")
    CSA% = XVALO(VALACT1$("CAUINGRE"))
    If CSA% > 0 Then
       Text55 = CSA%
    End If
    
    Command111.Enabled = True
End Sub

Private Sub Command15_Click()
   Screen.MousePointer = 11
5  ABUS$ = "Etiqueta de Cierre de O.F "   ' OF-008796-1"
   For U& = 1 To ULTREG&("PDOCLST")
      XX$ = REGLEIDO$("PDOCLST", U&)
      PPXX% = InStr(XX$, ABUS$)
      If PPXX% > 0 Then
        NOFA$ = TRIM$(Mid$(XX$, PPXX% + 26, 12))
        DOCUNU& = CVS(Left$(XX$, 4))
        GoSub EXPLORAETIQUETA
      End If
   Next U&
   '
   Screen.MousePointer = 1
   Call COMUNI("Recuperacion Finalizado")
   '
Exit Sub
   '
EXPLORAETIQUETA:
4  AORI$ = "P": If EXISTE%(LUDAT$ + "RDOCSPL.DT0") > 0 Then AORI$ = "R"
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Or LOF(NX0%) < 1 Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   l& = LI& + E&: If l& < 1 Or l& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, l&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = l&: GoTo 6
   If DCC& > DOCUNU& Then LS& = l&: GoTo 6
   GoTo 8
   '
7  ' Call MENSERR(24, "Documento no Registrado")
   GoTo 19
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While l& > 1
     Get #NX0%, l&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     l& = l& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = l& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   PPXX1% = InStr(TPSP$, NOFA$)
   If PPXX1% > 0 Then
      PPXX2% = InStr(1 + PPXX1%, TPSP$, NOFA$)
      If PPXX2% > PPXX1% Then
        PPXX3% = InStr(1 + PPXX2%, TPSP$, NOFA$)
        If PPXX3% > PPXX2% Then
           TTT$ = Mid$(TPSP$, PPXX3% - 16, 34)
           '
           Codigo$ = TRIM$(Mid$(TTT$, 1, 16))
           Ci% = CODINT(Codigo$)
           '
           IDLOT$ = TRIM$(Mid$(TTT$, 17, 12))
           Cantidad# = XVALO(Mid$(TTT$, 29, 8))
           CONDI$ = "IDENLOTE ='" & IDLOT$ & "' AND COD_INTE =" & Ci%
           VALORANTE = XVALO(VALOBADA("LOTINGRE", "ETIQXCAJA", CONDI$, "NOMESS"))
           If Cantidad# > VALORANTE Then
              Call ACTUREGISTRO("LOTINGRE", "EtiqxCaja", CONDI$, Cantidad#, REG&)
           End If
           '
        End If
      End If
   End If
   '
19 Return

End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    '
'10    OPX% = ALTERNA%("Re-impresión de Etiquetas\Modificación de Causa de Ingreso")
'    If OPX% < 1 Or OPX% > 2 Then GoTo 99
'    If OPX% = 1 Then ' REIMPRESION DE ETIQUETAS
'       Call REIMPETI
'       GoTo 10
'
'
'   ElseIf OPX% = 2 Then ' MODIFICACION DE CAUSA DE INGRESO
10  OPX% = ALTERNA%("Modificación de Causa de Ingreso\Re-fraccionado\Otros Accesos Directos")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    If OPX% = 1 Then
40     Call SELECHO("ECORDEF/Indice de Ordenes de Fabricación ")
       NORFA$ = TRIM$(VALACT1$("ECORDEF"))
       If NORFA$ = "" Then GoTo 10
       '
       CONDI$ = "CODIMOVI = 'OF' AND DOPRICOR = '" & NORFA$ & "' ORDER BY FECHMOV ASC"

50     Call CARGALISEL("CAUPOROR", "MOVISTO", "FECHMOV/D8;CANTINGRE/F9.1;CAUSA/F3;REFERCOR/A32", CONDI$)
       Call SELECHO("CAUPOROR/CAUSAS DE CIERRE")
       '
       If TRIM$(VALACT2$("CAUPOROR")) = "" Then GoTo 40
       '
       fech = FETEXCOR1$(CVINT(VALACT1$("CAUPOROR")))
       CANTI_INGRE = FORMATNUM$(Mid$(VALACT2$("CAUPOROR"), 1, 8), "F9.1")
       CONDI$ = "CANTINGRE = " & XVALO(CANTI_INGRE) & " AND " & CONDI$
       CAUS% = XVALO(VALOBADA("MOVISTO", "CAUSA", CONDI))
       'CAUS% = XVALO(Mid$(VALACT2$("CAUPOROR"), 9, 2))
       '
       Call REPLA(VDEF$, NORFA$, 1, 12)
       Call REPLA(VDEF$, MKI$(CAUS%), 13, 2)
       '
60     OBX$ = OBJPLA$("CAUSADIN", VDEF$) 'ABRE EL OBJETO
       '
       If OBX$ = "" Then GoTo 40
       CAUS% = CVI(Mid$(OBX$, 13, 2))
       REFE$ = TRIM$(ECOARCH$("CAUINGRE", MKI$(CAUS%)))
       If REFE$ = "" Then
          Call COMUNI("Causa No Válida")
          VDEF$ = OBX$
          GoTo 60
       End If
       '
       CONDI$ = "CANTINGRE = " & CANTI_INGRE & " AND FECHMOV = '" & fech & "'"
       Call ACTUREGISTRO("MOVISTO", "CAUSA", CONDI$, CAUS%, REGAF&, "NOMESS")
       Call ACTUREGISTRO("MOVISTO", "REFERCOR", CONDI$, REFE$, REGAF&, "NOMESS")
       Call COMUNI("Cambio Realizado")
       GoTo 40
    ElseIf OPX% = 2 Then
       Call Command31_Click
    ElseIf OPX% = 3 Then
       Call ACCDIR("CONSTOCK")
    End If
    
    
99  Command27.Enabled = True
End Sub


Private Sub Command31_Click()
    '
    Command31.Enabled = False
    '
    Call IMPRE_ETIQUETA(0)
    '
    Command31.Enabled = True
    '
End Sub

Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Label28_Change()
   CIXI% = CODINT%(Label28)
   Text6 = TRIM$(Str(DEPOPRE%(CIXI%)))
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text4.SetFocus
   End If

End Sub

Private Sub Text3_LostFocus()
   Text3 = TRIM$(FORMATNUM$(Text3, "F10.0"))
End Sub

Private Sub Text55_Change()
   '
   If XVALO(Text55) > 0 And ECOARCH$("CAUINGRE", MKI$(XVALO(Text55))) <> "" Then
      Label51 = ECOARCH$("CAUINGRE", MKI$(Text55))
   Else
       Label51 = ""
   End If
   '
End Sub

Private Sub Text6_Change()
   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
       CICOMPO% = CODINT%(VDEVU$)
       CACONSU = 0
       TTYY$ = Space$(128)
       Call REPLA(TTYY$, VDEVU$, 1, 16)
       Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
       Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
       On Error Resume Next
       USOUNI = CACONSU / XVALO(Text1)
       On Error GoTo 0
       Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
       List1.AddItem TTYY$
       List1.ListIndex = List1.ListCount - 1
       Text4 = ""
       Text4.SetFocus
    End If
    '
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       Call SELECHO("MASTER")
       VDEVU$ = VALACT1$("MASTER")
       If VDEVU$ <> "" Then
          CICOMPO% = CODINT%(VDEVU$)
          CACONSU = 0
          TTYY$ = Space$(128)
          Call REPLA(TTYY$, VDEVU$, 1, 16)
          Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
          Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13.1"), 57, 13)
          On Error Resume Next
          USOUNI = CACONSU / XVALO(Text1)
          On Error GoTo 0
          Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16.6"), 97, 16)
          List1.AddItem TTYY$, ILISTA&
          List1.ListIndex = ILISTA&
          Text4 = ""
          Text4.SetFocus
       End If
    End If
    Command12.Enabled = True
End Sub

Private Sub Command13_Click()
    '
    Command13.Enabled = False
    '
    If Label32 = "" Then
       Call MENSERR(24, "Nuúmero de O/F no Válido")
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
    ' SI ESTA TILDADA CHECK DE SOLO DEVOLUCION SALTEA ALGUNAS VALIDACIONES
    If Check1.Value = 1 Then GoTo 50
    '
    STAT% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IDSUBOR = '" & Label32 & "'", "NOMESS"))
    If STAT% >= 3 Then
       Call MENSERR(24, "Nuúmero de O/F no Válido (Orden Cerrada o Anulada)")
       GoTo 99
    End If
    '
    'VALIDAR CANTIDAD
    VALONU = XVALO(Tlabel38) ''''
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    '
    If XVALO(Text3) < 1 Then ' valido que la cantidad ingresada de rollos sea > 0
       Call MENSERR(24, "Cantidad de Etiquetas por Rollo No Válida")
       GoTo 99
    End If
    '
    If XVALO(Text3) > XVALO(Tlabel38) Then
       Call MENSERR(24, "La Cantidad de Etiquetas por Rollo No Puede ser Mayor a la Cantidad Realizada")
       GoTo 99
    End If
    '
    If XVALO(Text4) < 1 Then
       Call MENSERR(24, "Cantidad de Rollos por Caja No Válida")
       GoTo 99
    End If
    '
    DEPENTRA% = XVALO(Text6)
    If DEPENTRA% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPENTRA%)) = "" Then
       Call MENSERR(24, "Deposito de Entrada no Válido")
       GoTo 99
    End If
    '
    If TRIM$(Text2) = "" Then
       Call MENSERR(24, "Falta Ubicación en Depósito P.Terminado")
       GoTo 99
    End If
    '
50  If Check1.Value > 0 Then
       For i& = 0 To Text11.Count - 1
         If XVALO(Text11(i&)) > 0 Then GoTo 60
       Next i&
       For U& = 0 To Text1.Count - 1
        If XVALO(Text1(U&)) > 0.005 Then GoTo 60
        If U& = Text1.Count - 1 Then
          Call MENSERR(24, "Imposible Procesar Devolución\ \Falta Cantidad ")
          GoTo 99
        End If
       Next U&
    End If
60  For i = 0 To 2      'Valida que los metros a devolver no superen los metros de la bobina
        CANORIG = XVALO(VALOBADA("LOTINGRE", "CANALTA", "IDENLOTE = '" & TRIM$(Label23(i)) & "'", "NOMESS"))
        If CANORIG < 0.05 Then ' CALCULA LOS METROS CUADRADOS A PARTIR DEL ANCHO Y LARGO
            ANCH = XVALO(VALOBADA("LOTINGRE", "ANCHOMILI", "IDENLOTE = '" & TRIM$(Label23(i)) & "'", "NOMESS"))
            LARG = XVALO(VALOBADA("LOTINGRE", "METROINI", "IDENLOTE = '" & TRIM$(Label23(i)) & "'", "NOMESS"))
            CANORIG = (ANCH * LARG) * 1000
        End If
        If CANORIG > 0 Then LARGORIG = (CANORIG / XVALO(Label12(i)) * 1000)
        If XVALO(Text11(i)) > LARGORIG Then
           Call MENSERR(24, "Imposible Procesar Devolución\ \Cantidad de Metros a Devolver\Supera Metros Originales")
           GoTo 99
        End If
    Next i
    '
    '
    For U& = 0 To Text11.Count
       If XVALO(Text11(U&)) > 0 Then
          If XVALO(Label55(U&)) < 1 Then
             Call MENSERR(24, "Falta Indicar Depósito de Entrada Papel")
             GoTo 99
          End If
          If TRIM$(Label23(U&)) = "" Then
             Call MENSERR(24, "Papel Asignado sin Lote / Partida")
             GoTo 99
          End If
          If TRIM$(Text60(U&)) = "" Then
             Call MENSERR(24, "Falta Ingresar Ubicación \ Devolución de Papeles")
             GoTo 99
          End If

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
             Call MENSERR(24, "Falta Indicar Depósito de Entrada Tinta")
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
                Call MENSERR(24, "Falta Indicar Depósito de Entrada")
                GoTo 99
             End If
             '
             If U& < 4 Then   ' NO EXIGE UBICACION PARA NO TRAZABLES
               If TRIM$(Text61(U&)) = "" Then
                 Call MENSERR(24, "Falta Ingresar Ubicación \ Devolución de Tintas e Inzumos Trazables")
                 GoTo 99
               End If
             End If
          End If
       End If
    Next U&
    '
    If Check1.Value < 1 Then ' SI NO ESTA TILDADO SOLO DEVOLUCION
      Call GRACIERROF
    End If
    Call GRABACONSUM
        CICONJ% = CODINT%(Label28)
        LOTE$ = Label32
    'Call PRINSTIKAPRO(CICONJ%, LOTE$, DEPENTRA%)
    '
    If Check1.Value < 1 Then ' SI NO ESTA TILDADO SOLO DEVOLUCION
        NORFA$ = TRIM$(Label32)
        CANTI = XVALO(Tlabel38)
        ETIROLLO = XVALO(Text3)
        ROLLOCAJA = XVALO(Text4)
        Ci% = CODINT%(Label28)
        '
        Call PRINETICIEROF(Ci%, NORFA$, CANTI, ETIROLLO, ROLLOCAJA)
    '
    End If
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
Sub GRACIERROF()
     '1.- Actualiza saldo pendiente de la Orden de Fabricacion
     NORFA$ = Label32
     CANTI = XVALO(Tlabel38)
     SCRAP = 0
'     CIERRA% = 0: If XVALO(Tlabel38) + XVALO(Label37) >= XVALO(Label34) Then CIERRA% = 1
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM  ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
     Dim ABREORFFFAB As ADODB.Recordset
     Set ABREORFFFAB = New ADODB.Recordset
     ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     '
     On Error Resume Next
     With ABREORFFFAB
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           'REFEOF$ = !REFEREN
          !CanApro = !CanApro + CANTI
          !CanRech = !CanRech + SCRAP
          '
          'Permite cerrar la OF aunque se encuentre parcialmente informada
          CANTAPRO = XVALO(!CanApro)
          CANTPROY = XVALO(!CanProy)
          CANPEN = CANTPROY - CANTAPRO
          CIERRA% = 0
          If CANPEN <= 0 Then
            CIERRA% = 1
          Else
            MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un saldo Pendiente.\\Mantener Abierta\Cerrar")
            If MDXX% = 2 Then CIERRA% = 1
          End If
          '
          If CIERRA% = 1 Then
            !Statuorf = 3         ' MARCA COMO CERRADA
          End If
          !FechCierr = CVDAT(FECHANUM(Label24))
          .Update
        End If
        On Error GoTo 0
     End With
     ABREORFFFAB.Close
     Set ABREORFFFAB = Nothing
     '
     If CANTI >= 0.05 Then
        DEPX% = XVALO(Text6)
        If DEPO% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
        NUORIT% = 0
        FECHA% = FECHANUM(Label24)
        CICONJ% = CODINT%(Label28)
        LOTE$ = Label32
        CMOV$ = "OF"
        DOPRI$ = Label32
        DOSEC$ = Label32
        CONDI$ = "IDSUBOR='" & Label32 & "'"
        REFE$ = VALOBADA("ORDEFABR", "REFEREN", CONDI$) 'TRIM$(Label51)
        VUNI# = 0
        MONEII% = 1
        NC% = 0: UBIK$ = Left$(Text2, 32)
        Call CREALOTIN(FECHA%, FECHA%, 32767, CICONJ%, LOTE$, "ACHERNAR", LOTE$, CDbl(CANTI), 0, 0, DEPX%, UBIK$)
        '
        'GUARDAR DATOS EN NUEVOS CAMPOS
        EXCAJA# = XVALO(Text4) * XVALO(Text3)
        CONDI$ = "COD_INTE= " & CICONJ% & " AND IDENLOTE='" & LOTE$ & "'"
        Call ACTUREGISTRO("LOTINGRE", "Rolxcaja", CONDI$, XVALO(Text4), REG&)
        Call ACTUREGISTRO("LOTINGRE", "EtiqxRollo", CONDI$, XVALO(Text3), REG&)
        Call ACTUREGISTRO("LOTINGRE", "EtiqxCaja", CONDI$, EXCAJA#, REG&)
        '
        Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI)
        ' SERIA RECOMENDABLE INGRESAR COMO PARAMETRO TAMBIEN LA CAUSA EN EL MOVISTOK
        CONDI$ = "DOPRICOR = '" & DOPRI$ & "' AND FECHMOV = '" & FETEXCOR1$(FECHA%) & "'"
        Call ACTUREGISTRO("MOVISTO", "CAUSA", CONDI$, XVALO(Text55), REGAF&, "NOMESS")
        Call ACTUREGISTRO("MOVISTO", "USER_TRAN", CONDI$, USERNAME$, REGAF&, "NOMESS")
        '
        '
     End If
     '
     If CICONJ% > 0 Then
        SQLX$ = "UPDATE  MASTER "
        SQLX$ = SQLX$ + " SET DepoEntra = " & DEPX%
        SQLX$ = SQLX$ + " WHERE CODINT = " & CICONJ% & ""
         FLEXCONN.Execute (FILTSQL$(SQLX$))
     End If
     '
End Sub
'
Sub GRABACONSUM()
     '
     FF% = FECHANUM(Label24)
     CMO$ = "CF"    ' CONSUMO DE FABRICACION
     DOPRI$ = Label32
     DOSEC$ = Label32
     REFE$ = "Sobrante " + Label32
     NUORIT% = 0
     '
     For YX& = 0 To 2
        NLOTX$ = TRIM$(Label23(YX&))
        If NLOTX$ <> "" Then
           CIXI% = CODINT%(Label27)
           NUORIT% = NUORIT% + 1
           CANTI = XVALO(Label35(YX&))
           LOTE$ = Label23(YX&)
           DepOrig% = XVALO(Label55(YX&))
           UBIC$ = TRIM$(Text60(YX&))
           
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                         '\\\OT-06-0244-WAR-20/06/06///
           CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "' "
           Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBIC$, REGAF&, "NOMESS")
        End If
     Next YX&
     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = Label40(YX&)
           CANTI = XVALO(Text1(YX&))
           DepOrig% = XVALO(Label52(YX&))
           UBIC$ = TRIM$(Text61(YX&))
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
           CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "' "
           Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBIC$, REGAF&, "NOMESS")
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = ""
           CANTI = XVALO(Text1(YX&))
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
        NLOTX$ = TRIM$(Label23(YX&))
        If NLOTX$ <> "" Then
           CIXI% = CODINT%(Label27)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(2, CAITAB1%) = Label35(YX&)
           TETABU1$(3, CAITAB1%) = Label27
           TETABU1$(4, CAITAB1%) = Label6
           TETABU1$(5, CAITAB1%) = Label23(YX&)
           TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label23(YX&))
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

Sub Command14_Click()
    '
    Command14.Enabled = False
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
10  OPX% = ALTERNA%("Configuración de Formularios\Tabla de Ingreso de Motivos")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      OPOFTRA07.Show 1
      GoTo 10
    ElseIf OPX% = 2 Then
        Call TRALOCAL("CAUINGRE", "")
        j& = 0
        VTB% = VENTABU%("CAUINGRE/NC/TITUPAN=Tabla de Causas de Ingresos ")
        If VTB% >= 0 Then
            For i& = 1 To ULTREG&("!CAUINGRE")
                x$ = REGLEIDO$("!CAUINGRE", i&)
                If CVI(Left$(x$, 2)) > 0 Then
                    j& = j& + 1
                    Call GRAREG("CAUINGRE", x$, j&)
                End If
            Next i&
            Call SETULTREG("CAUINGRE", j&)
            Call CIERRARCH("CAUINGRE")
       End If
   End If
    
    GoTo 10
99  Command14.Enabled = True
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
        Call REPLA(y$, !DoPriLar, 1, 12)
        FEX = CVINT(!Fechmov)
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
        LOTE$ = !IDENLOTE
        CIXI% = Cod_Inte
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
       Label21(Index) = DESCRIT0$(CICOMPO%)
       Label42(Index) = Unimed(CICOMPO%)
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
'   CUDI$ = UCase$(CurDir)
'   PPXX% = InStr(CUDI$, "\WINDOWS")
'   If PPXX% > 0 Then
'      RUTAX$ = Left$(CUDI$, PPXX%) + "MANUALES\"
'      If EXISTE%(RUTAX$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE RUTAX$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
'      End If
'   End If
'   '
'20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
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
    Call ACTECOR_ACHERNAR(2)     ' para seleccionar solo las activas lanzadas
    ENCABE$ = "Orden             Código y Descripción                                                          Cant. Proy."

    Call CommSel("ECOACHER/" & ENCABE$, Command5, Label32)
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
    '
    
End Sub

Sub TAMARCAS()
    '
    Ci% = CODINT%(Label27)
    If Ci% > 0 Then GoTo 20
    '
10  Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
      Ci% = CODINT(COD1$)
20    If Ci% > 0 Then
         Call FILTERFILE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
         VTB% = VENTABU%("TAMARCAS")
         If VTB% >= 0 Then
            Call FILTERUPDATE("TAMARCAS", "!TAMARCAS", 4, MKI$(Ci%))
            If CODINT%(Label27) < 1 Then GoTo 10
         End If
      End If
    End If
    '
End Sub
'
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

Sub REIMPETI()
   '
   Call SETULTREG("!ECORDEP", JJ&)
   Call FRESHECHO("!ECORDEP")
2  Call SELECHO("!ECORDEP/Indice General de Vales de Material")
   NVALE$ = TRIM$(VALACT1$("!ECORDEP"))
   If NVALE$ = "" Then Exit Sub
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
   Exit Sub

9  Call MUESTRADOC(DOCUNU&)
   GoTo 2
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
    Ci% = !Cod_Inte
    CARPRODOF$ = CODEXT$(Ci%)
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
      For KKU& = 0 To 2
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
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       Frame9.Visible = True
    End If
    '
    If HOY(HO$) >= FECHANUM("30/04/10") Then
       Command15.Visible = False
    End If
    
    Call FRESHALL
    '
    'PROCESO PARA RECUPERAR, POR QUE SE DETECTO QUE AL GRABAR LOS CIERRES SOLO
    'CONSIDERABA QUE LA CANTIDAD APROBADA SEA > A LA PROYECTADA PARA QUE LE PONGO
    'EL STATUS = 3
    Call ABREORFAB
    SQLX$ = "UPDATE OrdeFabr SET STATUORF = 3 WHERE CANAPRO >= CANPROY AND STATUORF < 3 "
    FLEXCONN.Execute (FILTSQL$(SQLX$))
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
    SQLX$ = "SELECT MOVISTO.DOPRICOR AS MIDENLOTE, "
    SQLX$ = SQLX$ + " MOVISTO.REFERCOR AS MREFERCOR, "
    SQLX$ = SQLX$ + " ORDEFABR.REFEREN AS OREFEREN, "
    SQLX$ = SQLX$ + " ORDEFABR.IDSUBOR AS OIDSUBOR "
    SQLX$ = SQLX$ + " FROM MOVISTO INNER JOIN ORDEFABR "
    SQLX$ = SQLX$ + " ON MOVISTO.DOPRICOR = ORDEFABR.IDSUBOR "
    SQLX$ = SQLX$ + " WHERE CODIMOVI ='OF' "
    SQLX$ = SQLX$ + " AND CANTINGRE > 0 "
    SQLX$ = SQLX$ + " AND LEFT(MOVISTO.REFERCOR,10) <> LEFT(ORDEFABR.REFEREN,10)"
    '
    Dim CONSULTA As ADODB.Recordset
    Set CONSULTA = New ADODB.Recordset
25  On Error Resume Next
    CONSULTA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With CONSULTA
      Do While Not (.EOF)
       !MREFERCOR = SAFETEXT(!OREFEREN)
      .Update
      .MoveNext
      Loop
      .Close
    End With
End Sub

Private Sub Label27_Change()
   CIXI% = CODINT%(Label27)
   If CIXI% > 0 Then
      Label6 = DESCRIT0$(CIXI%)
      Call ARMALILOT(CIXI%, 0, 0, 1, 0, "", 1)
      Screen.MousePointer = 1
   End If
End Sub

Private Sub Label32_Change()
    ' CASO 1 - PONE EL CAMPO EN BLANCO
    NORFA$ = Label32
       NVM$ = TRIM$(ProNroComprobante("CF"))
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
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
    Label37 = ""
    Tlabel38 = ""
    Label27 = ""
    Label6 = ""
    Label16 = ""
    Text4 = ""
    Check1.Value = 0
    '
    For KKU& = 0 To 2
       Label23(KKU&) = ""
       Text11(KKU&) = ""
       Label12(KKU&) = ""
       Label25(KKU&) = ""
       Label55(KKU&) = ""
       Text60(KKU&) = ""
    Next KKU&
    '
    For KKU& = 0 To 3
       Label39(KKU&) = ""
       Text1(KKU&) = ""
       Label40(KKU&) = ""
       Label41(KKU&) = ""
       Label42(KKU&) = ""
       Label52(KKU&) = ""
       Text61(KKU&) = ""
    Next KKU&
    '
    For KKU& = 4 To 7
       Label39(KKU&) = ""
       Text1(KKU&) = ""
       Label21(KKU&) = ""
       Label42(KKU&) = ""
       Label52(KKU&) = ""
       Text61(KKU&) = ""
    Next KKU&
    '
    Text55 = ""
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
           Tlabel38 = FORMATNUM$(CAPENDE#, "F10.1")
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
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label40_Change(Index As Integer)
    Call MUESTRASEL1(Index)
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
    If XVALO(Text1(Index)) = 0 Or Text1(Index) = "" Then
       Label21(Index) = ""
       Label39(Index) = ""
       Label52(Index) = ""
       Label42(Index) = ""
       If Index < 4 Then
          Label40(Index) = ""
          Label41(Index) = ""
       End If
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index) = FORMATNUM(XVALO(Text1(Index)), "F10.1")
End Sub

Private Sub Text11_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 2
       If Label23(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text11(KKI%))
         Else
           Text11(KKI%) = ""
       End If
       Label35(Index) = FORMATNUM$(XVALO(Text11(Index)) * XVALO(Label12(Index)) / 1000, "F10.1")
    Next KKI%
    Label50 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
    If XVALO(Text11(Index)) = 0 Or Text11(Index) = "" Then
       Label23(Index) = ""
       Label55(Index) = ""
       Label12(Index) = ""
       Label25(Index) = ""
       Label35(Index) = ""
    End If
    '
End Sub

Private Sub Text11_DblClick(Index As Integer)
    Text11(Index) = ""
End Sub


Private Sub Text17_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index < 5 Then
          Command6(Index + 1).SetFocus
       End If
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       Text2.SetFocus
    End If

End Sub

Private Sub Text11_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Text11_LostFocus(Index)
    End If
End Sub

Private Sub Text11_LostFocus(Index As Integer)
    Text11(Index) = TRIM$(FORMATNUM$(XVALO(Text11(Index)), "F10.1"))
End Sub

Private Sub Text4_LostFocus()
   '
   Text4 = TRIM$(FORMATNUM$(Text4, "F10.0"))
   '
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
       Call Text4_LostFocus
    End If
    
End Sub

Private Sub Text8_GotFocus()
    If UCase$(TRIM$(Label7)) = "AUTO" Then Command13.SetFocus
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If
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

Private Sub Tlabel38_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text3.SetFocus
   End If

End Sub

Private Sub Tlabel38_LostFocus()
   Tlabel38 = TRIM$(FORMATNUM$(Tlabel38, "F10.0"))
End Sub

Private Sub Command8_Click()
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    Call CommSel("ECORDEF/Indice de Ordenes de Fabricación ", Command8, Label32)
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
   '
   '
End Sub

Sub ACTECOR_ACHERNAR(MODOI%)
    '
    ' con MODOI% = 2 selecciona solo activas lanzadas (status =1
    Screen.MousePointer = 11
'    Call COPYSTRU("ECORDEF", "ECOACHER")
    j& = 0
    '
    Call ABREORFAB
    '
    If MODOI% = 2 Then
        CONDI$ = " StatuOrf < 3 "   ' para que no este anulada ni cumplida
        CONDI$ = CONDI$ + " AND CanApro < (CanProy - 0.005) "
      Else
        CONDI$ = " StatuOrf < 3 "
    End If
    '
    If CantRegistros("ORDEFABR", CONDI$, "NOMESS") < 1 Then
        Call MENSERR(24, "NO Existen Ordenes de Fabricación\Correspondientes a Criterio de Selección.\  \Imposible Continuar la Carga.")
        GoTo 90
    End If
    '
    'ORDENO LA CONSULTA
    CONDI$ = CONDI$ + " ORDER BY IdSubOr "
    Call CARGALISEL("ECOACHER", "ORDEFABR", "IDSUBOR/A12;REFEREN/A40;CANPROY/F16", CONDI$, "DISTINCT")
    '
'90  Call SETULTREG("ECOACHER", JJ&)
'    Call CIERRARCH("ECOACHER")
'    Call FRESHECHO("ECOACHER")
90   Screen.MousePointer = 1
    '
End Sub



