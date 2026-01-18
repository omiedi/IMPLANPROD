VERSION 5.00
Begin VB.Form APROLOT09 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspeccion y Aprobacion"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8415
      Picture         =   "APROLOT09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   5055
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   210
      TabIndex        =   63
      Top             =   4050
      Width           =   8100
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   86
         Top             =   1320
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   85
         Top             =   240
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Impresion Fantasma"
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
         Left            =   5955
         TabIndex        =   84
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2040
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rechazado"
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
         Left            =   6675
         TabIndex        =   68
         Top             =   915
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Aprobado"
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
         Left            =   5955
         TabIndex        =   67
         Top             =   645
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.TextBox Text6 
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
         Left            =   1185
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   675
         Width           =   4650
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Analisis"
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
         Left            =   4515
         TabIndex        =   69
         Top             =   255
         Width           =   1245
      End
      Begin VB.Label Label45 
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
         Left            =   1920
         TabIndex        =   66
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario / Inspector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -300
         TabIndex        =   65
         Top             =   285
         Width           =   2145
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   105
         TabIndex        =   64
         Top             =   735
         Width           =   1020
      End
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1470
      TabIndex        =   0
      Top             =   315
      Width           =   2160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Identificacion de Lote / Partida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   195
      TabIndex        =   50
      Top             =   1365
      Width           =   8100
      Begin VB.CommandButton Command12 
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
         Left            =   7665
         TabIndex        =   8
         Top             =   1890
         Width           =   175
      End
      Begin VB.CommandButton Command11 
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
         Left            =   6180
         TabIndex        =   6
         Top             =   1890
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
         Left            =   4695
         TabIndex        =   4
         Top             =   1890
         Width           =   175
      End
      Begin VB.TextBox Text10 
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
         Left            =   6615
         TabIndex        =   7
         Top             =   1890
         Width           =   1050
      End
      Begin VB.TextBox Text9 
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
         Left            =   5145
         TabIndex        =   5
         Top             =   1890
         Width           =   1050
      End
      Begin VB.TextBox Text8 
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
         Left            =   3660
         TabIndex        =   3
         Top             =   1890
         Width           =   1050
      End
      Begin VB.CommandButton Command8 
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
         Left            =   2745
         TabIndex        =   2
         Top             =   435
         Width           =   175
      End
      Begin VB.Label Label57 
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
         Left            =   6865
         TabIndex        =   91
         Top             =   435
         Width           =   1065
      End
      Begin VB.Label Label52 
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
         Height          =   210
         Left            =   5940
         TabIndex        =   90
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   83
         Top             =   885
         Width           =   1545
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
         Height          =   285
         Left            =   6360
         TabIndex        =   82
         Top             =   840
         Width           =   1575
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
         Left            =   1155
         TabIndex        =   75
         Top             =   1860
         Width           =   2235
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
         Left            =   2910
         TabIndex        =   74
         Top             =   1260
         Width           =   5040
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   73
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "F.Fabricación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3645
         TabIndex        =   72
         Top             =   1665
         Width           =   1425
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Retest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4760
         TabIndex        =   71
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F.Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6210
         TabIndex        =   70
         Top             =   1680
         Width           =   1620
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   315
         TabIndex        =   62
         Top             =   1305
         Width           =   795
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
         Left            =   1170
         TabIndex        =   61
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label38 
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
         Left            =   3000
         TabIndex        =   60
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
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
         Left            =   3615
         TabIndex        =   59
         Top             =   435
         Width           =   960
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Left            =   2055
         TabIndex        =   58
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label36 
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
         Left            =   1140
         TabIndex        =   57
         Top             =   435
         Width           =   1650
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id.Bulto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -210
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1460
         TabIndex        =   55
         Top             =   885
         Width           =   1545
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
         Left            =   1155
         TabIndex        =   54
         Top             =   825
         Width           =   1200
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -210
         TabIndex        =   53
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Orig:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4640
         TabIndex        =   52
         Top             =   495
         Width           =   585
      End
      Begin VB.Label Label11 
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
         Left            =   5250
         TabIndex        =   51
         Top             =   435
         Width           =   1065
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   105
      TabIndex        =   21
      Top             =   6015
      Width           =   8205
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
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         Top             =   1725
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Top             =   1725
         Width           =   2055
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
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   22
         Top             =   1260
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   2205
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8205
         Y1              =   440
         Y2              =   440
      End
      Begin VB.Line Line3 
         X1              =   8190
         X2              =   8190
         Y1              =   120
         Y2              =   2205
      End
      Begin VB.Line Line4 
         X1              =   15
         X2              =   15
         Y1              =   120
         Y2              =   2310
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   8200
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   2310
         TabIndex        =   38
         Top             =   1000
         Width           =   810
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Id.lote:"
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
         Left            =   435
         TabIndex        =   37
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         Left            =   735
         TabIndex        =   36
         Top             =   1000
         Width           =   930
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1770
         Width           =   930
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
         Left            =   1080
         TabIndex        =   34
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   6405
         TabIndex        =   33
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Id.lote:"
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
         Left            =   4515
         TabIndex        =   32
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Kilos:"
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
         Left            =   4650
         TabIndex        =   31
         Top             =   1000
         Width           =   810
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicación:"
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
         Left            =   4200
         TabIndex        =   30
         Top             =   1770
         Width           =   930
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
         Left            =   5160
         TabIndex        =   29
         Top             =   600
         Width           =   2235
      End
      Begin VB.Label Label32 
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
         Left            =   5160
         TabIndex        =   28
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Label Label33 
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
         Left            =   6720
         TabIndex        =   27
         Top             =   1260
         Width           =   1200
      End
      Begin VB.Label Label34 
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
         Left            =   5160
         TabIndex        =   26
         Top             =   1725
         Width           =   2115
      End
      Begin VB.Label Label20 
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
         Left            =   2625
         TabIndex        =   25
         Top             =   1260
         Width           =   1200
      End
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
      Left            =   7980
      TabIndex        =   18
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
      Left            =   3600
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5055
      Left            =   8400
      ScaleHeight     =   4995
      ScaleWidth      =   1320
      TabIndex        =   14
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Ficha Kardex"
         Top             =   3900
         Width           =   640
      End
      Begin VB.CommandButton Command38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Re-Imprimir Certificado de Calidad"
         Top             =   2640
         Width           =   640
      End
      Begin VB.CommandButton Command37 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   1380
         Width           =   640
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Modificación de Datos"
         Top             =   3270
         Width           =   640
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2010
         Width           =   640
      End
      Begin VB.CommandButton Command9 
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
         Left            =   3255
         TabIndex        =   49
         Top             =   0
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   46
         Top             =   1050
         Width           =   6675
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7035
         TabIndex        =   45
         Top             =   525
         Width           =   1320
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
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   750
         Width           =   640
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
         Height          =   620
         Left            =   105
         Picture         =   "APROLOT09.frx":181E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   640
      End
      Begin VB.Label Label10 
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
         Left            =   1680
         TabIndex        =   48
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label9 
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
         Height          =   210
         Left            =   210
         TabIndex        =   47
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   270
         Left            =   0
         TabIndex        =   44
         Top             =   1095
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5355
         TabIndex        =   43
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label18 
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
         Left            =   1680
         TabIndex        =   42
         Top             =   525
         Width           =   1650
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   41
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Height          =   210
         Left            =   2730
         TabIndex        =   40
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label2 
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
         Left            =   4515
         TabIndex        =   39
         Top             =   525
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "APROLOT09.frx":1968
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4600
         Width           =   1515
      End
   End
   Begin VB.Label Label41 
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
      Height          =   210
      Left            =   7320
      TabIndex        =   89
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label49 
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
      Left            =   7830
      TabIndex        =   88
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label54 
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
      Left            =   4875
      TabIndex        =   79
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Recep.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4050
      TabIndex        =   78
      Top             =   180
      Width           =   780
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   5355
      TabIndex        =   20
      Top             =   375
      Width           =   1545
   End
   Begin VB.Label Label13 
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
      Left            =   7035
      TabIndex        =   19
      Top             =   315
      Width           =   1020
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
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   210
      TabIndex        =   17
      Top             =   795
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
      Height          =   285
      Left            =   1470
      TabIndex        =   16
      Top             =   735
      Width           =   5730
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   -210
      TabIndex        =   15
      Top             =   375
      Width           =   1545
   End
End
Attribute VB_Name = "APROLOT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub


Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text8 = VDEV$
   End If
End Sub

Private Sub Command11_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text9 = VDEV$
   End If
End Sub

Private Sub Command12_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text10 = VDEV$
   End If
End Sub


Private Sub Command14_Click()
   '
   Command4.Enabled = False
   If DERACCE%("MODLOPRO", "Modificacion Lote del Proveedor") > 1 Then
  
10    LOTE$ = LOTSELEC$(CIXI%)
      If CIXI% < 1 Or TRIM(LOTE$) = "" Then
         GoTo 99
      End If
      '
      ' PARA ACTUALIZAR NUMERO DE CERTIFICADO EN BALAREQMO
      CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE ='" & LOTE$ & "'"
      NUCERT& = XVALO(VALOBADA("LOTINGRE", "NuCert_Apro", CONDI$))
      Call ACTUREGISTRO("BALAREQMO", "NuCert_Apro", CONDI$, NUCERT&, REG&)
      '
      PARAMETROS$ = SAFETEXT(CIXI%) + ";" + LOTE$
      Call STDFORM("MODIRECE", PARAMETROS$)
      GoTo 10
   End If
99 Command4.Enabled = True
   '
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    CIXXI% = CODINT(Text7)
    FFXX% = FECHANUM(Label13.Caption)
    '
    If FFXX% = 0 Then
       Call MENSERR%(24, "Fecha Invalida")
       GoTo 99
    End If
    If CIXXI% = 0 Then
       Call MENSERR%(24, "Codigo de Articulo Invalido")
       GoTo 99
    End If
    '
    If ESTRAZABLE(CIXXI%) > 0 Then
        If XVALO(CONTROL("BOLRECEP", "IFERTEST")) > 0 Then
           If FECHANUM(Text9) < 1 Then
              COMUNI ("Falta Ingresar Fecha de Retest")
              Text9.SetFocus
              GoTo 99
            End If
            '
            If FECHANUM(Text8) > FECHANUM(Text9) Then
               Call COMUNI("Fecha de Retest es Menor a Fecha de Fabricación")
               Text9.SetFocus
               GoTo 99
            End If
            '
        End If
        '
        If XVALO(CONTROL("BOLRECEP", "IFEVEN")) > 0 Then
           If FECHANUM(Text10) < 1 Then
              Call COMUNI("Falta Ingresar Fecha de Vencimiento")
              Text10.SetFocus
              GoTo 99
            End If
            '
            If XVALO(CONTROL("BOLRECEP", "IFERTEST")) > 0 Then
               If FECHANUM(Text9) > FECHANUM(Text10) Then
                  Call COMUNI("Fecha de Vencimiento es Menor a Fecha de Retest")
                  Text10.SetFocus
                  GoTo 99
               End If
            End If
            '
        End If
        '
        If FECHANUM(Text9) < 1 Then Text6 = Text6 + " ATENCION: Falta Fecha de Re-Test."
        If FECHANUM(Text10) < 1 Then Text6 = Text6 + " ATENCION: Falta Fecha de Vencimiento del Lote."
        '
    End If
    '
    '
    COD$ = TRIM$(Text7)                     'Codigo ext
    CI% = CODINT(COD$)                      'Cod Interno
    '
    CONDI$ = "codint =" & CI% & "and Trazable = 1"
    Traza% = XVALO(VALOBADA("Master", "Trazable", CONDI$))
    If Traza% > 0 Then
'       If Text11 = "" Then
'          Call COMUNI("Ingresar Lote")
'          Text11.SetFocus
'          GoTo 99
'       End If
    End If
    '
    'SE ANULA LA VALIDACION DE LA REFERENCIA
'    If Text6 = "" Then
'       Call MENSERR%(24, "Debe Ingresar Referencia de Aprobación")
'       Text6.SetFocus
'       GoTo 99
'    End If
    'Nuevos campos a guardar 05/10/2009
    NUCERT& = XVALO(Label54)
    CONDI2$ = "NuCert_Apro = " & NUCERT&
    If CantRegistros("LOTINGRE", CONDI2$) > 0 Then
       NUCERT& = XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0")) + 1
       Label54 = NUCERT&
       Call COMUNI("Numero de Certificacion de Aprobacion fue Cambiado por " & NUCERT&)
       GoTo 99
    End If
    '
    Dim LOTAPRO As ADODB.Recordset
    LOTPRO$ = TRIM$(Label51)
    If LOTPRO$ <> "" Then
       CONDI$ = "COD_INTE = " & CIXXI%
       CONDI$ = CONDI$ & " AND LoteCol_Prov ='" & LOTPRO$ & "'"
       CONDI$ = CONDI$ & " AND LIBERA ='CUAR'"
       QTR% = CantRegistros("LOTINGRE", CONDI$)
       If QTR% > 1 Then
         TTXX$ = "Hay " & QTR% & " Partidas del Mismo Lote Proveedor.\   \¿ Desea Realizar una Aprobación / Rechazo Masivos\de TODAS esta Partidas ?\\No, Procesar Cada Una\Solo Para Igual I-R\Si, Procesar TODAS"
         OPPX% = COMALTER%(TTXX$)
         If OPPX% = 2 Then
            Screen.MousePointer = 11
            NIRX$ = TRIM$(Label36)
            PPXX% = InStr(NIRX$, "-"): If PPXX% < 2 Then PPXX% = 1 + Len(NIRX$)
            If PPXX% < 2 Then
               Call MENSERR(24, "Peticion Imposible")
               GoTo 99
            End If
            NIR$ = Left$(NIRX$, PPXX% - 1)
            CONDI$ = CONDI$ + " AND IdenLote LIKE '" & NIR$ & "%' "
         End If
         If OPPX% = 2 Or OPPX% = 3 Then
           Screen.MousePointer = 11
           FIN& = QTR%: UU& = 0
           Set LOTAPRO = New ADODB.Recordset
             SQLX$ = "SELECT IDENLOTE FROM LOTINGRE WITH(NOLOCK) "
             SQLX$ = SQLX$ + " WHERE " & CONDI$
           Set LOTAPRO = FLEXCONN.Execute(FILTSQL$(SQLX$))
           With LOTAPRO
             Do While Not .EOF
               UU& = UU& + 1
               Call TRACE(20, UU&, FIN&)
               IDLOT$ = !idenLote
               Call ApruebaImprime(IDLOT$)
             .MoveNext
             Loop
           End With
           LOTAPRO.Close
           Set LOTAPRO = Nothing
           GoTo 90
         End If
       End If
    End If
    '
    ' PARA APROBACION INDIVIDUAL
    IDLOT$ = TRIM$(Label36)
    Call ApruebaImprime(IDLOT$)
    '
    '
    Screen.MousePointer = 11
    '
    IDLOT$ = TRIM$(Label36)
    '
90  If Option1.Value = True Then
        Call COMUNI("Aprobación Procesada")
        'ENVIO DE MAIL PARA RECEPCIONES APROBADAS
        Call ENVIMAILAPROB(CI%, IDLOT$)
      ElseIf Option2.Value = True Then
        Call ENVIMAILRECH(CI%, IDLOT$)
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE='" & IDLOT$ & "'"
        Call BORRAREGISTROS("BALAREQMO", CONDI$, REG&)
        Call COMUNI("Rechazo Procesado")
    End If
    '
    CIXXI% = CODINT%(Text7)
    If CIXXI% > 0 Then
      CONDI$ = "COD_INTE = " & CIXXI% & " AND LIBERA = 'CUAR'"
      REGG& = CantRegistros&("LOTINGRE", CONDI$)
      If REGG < 1 Then
           Text7 = ""
         Else
           Call Text7_Change
      End If
    End If
    Call LIMPCONMAT
    Call Lista_PendApro
    '
99  Command2.Enabled = True
    Screen.MousePointer = 1
    '
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    FIKARDEX09.Text1 = Text7
    FIKARDEX09.Show 1
    Command20.Enabled = True
End Sub

Private Sub Command37_Click()
   
   CFGRECEP.Show 1
   
End Sub

Private Sub Command38_Click()
    Command38.Enabled = False
    LOTE$ = LOTSELEC$(CIXI%)
    If CIXI% < 1 Or TRIM(LOTE$) = "" Then
       GoTo 99
    End If
    '
    ' PARA ACTUALIZAR NUMERO DE CERTIFICADO EN BALAREQMO
    CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE ='" & LOTE$ & "'"
    NUCERT& = XVALO(VALOBADA("LOTINGRE", "NuCert_Apro", CONDI$))
    Call ACTUREGISTRO("BALAREQMO", "NuCert_Apro", CONDI$, NUCERT&, REG&)
    '
    'Presentar Formulario en pantalla
    Screen.MousePointer = 11
    TTTT$ = LOTE$
    FORMU$ = CONTROL$("BOLRECEP", "FORAPRO")
    If TRIM(FORMU$) <> "" Then
       If EXISTE%(LUDAT$ + FORMU$ + ".SQR") > 0 Then
          Call STDFORM(FORMU$, TTTT$, "PRINT")
       End If
    End If
    '
99  Screen.MousePointer = 1
    Command38.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    'Arma lista seleccion pendiente de aprobacion
     Call SELECHO("PENDAPRO")
     If VALACT1$("PENDAPRO") <> "" Then
        Text7 = VALACT1$("PENDAPRO")
     End If

99   Command8.SetFocus
End Sub

Private Sub Command6_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label13 = VDEV$
   End If
End Sub

Private Sub Command8_Click()
    '
    'Presenta lista de seleccion con el numero de remito de ingreso
    Call SELECHO("!LILOTRE/Lista de Seleccion de Lotes Pendientes de Aprobacion")
    IDLO$ = VALACT1$("!LILOTRE")
    If IDLO$ <> "" Then
       CIXI% = CODINT%(Text7.TEXT)
       Call ACSALOTE(CIXI%, IDLO$)
       Label36.Caption = IDLO$
    End If
    '
End Sub

Private Sub Form_Load()
   Label13 = FECHATEX(HOY(HO$))
   Label45 = USERNAME
   Visible = True
   Call Lista_PendApro
   NUCERT& = XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0")) + 1
   Label54 = NUCERT&
   Hide
End Sub

Private Sub Label36_Change()
    '
    If Label36 = "" Then
       Exit Sub
    End If
    CIXXI% = XVALO(CODINT(Text7))
   
    'Label11 = FORMATNUM(SALNOCONS(Label36, CIXXI%), "F9.1")
    '
    CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE ='" & TRIM$(Label36) & "'"
    Label11 = TRIM$(FORMATNUM(VALOBADA("LOTINGRE", "CanAlta", CONDI$), "F13." & CNTDC$(CIXXI%))) ' CANTIDAD
    CanConsu = FORMATNUM(VALOBADA("LOTINGRE", "CanConsu", CONDI$), "F13." & CNTDC$(CIXXI%)) ' CANTIDAD
    'SE AGREGA PARA PRESENTAR EL SALDO DEL LOTE
    Label57 = TRIM$(FORMATNUM(XVALO(Label11) - XVALO(CanConsu), "F13." & CNTDC$(CIXXI%)))
      FALTA = CVINT(VALOBADA("LOTINGRE", "FECHALTA", CONDI$)) 'FECHA ALTA
    Label37 = FECHATEX$(FALTA)
    Label42 = VALOBADA("LOTINGRE", "COD_DESTI", CONDI$) 'DESTINO
    Label38 = VALOBADA("Lotingre", "ORIGEN", CONDI$)
    Label24 = VALOBADA("LOTINGRE", "UBICFIS", CONDI$)
    'Label47 = Label24 'REUBICACION
    Label51 = VALOBADA("LOTINGRE", "LoteCol_Prov", CONDI$)
    'SE COMENTA PORQUE EL LABEL52 NO EXISTE. LA UBICACION ESTA EN EL LABEL24 Y SE CARGA MAS ARRIBA
    'Label52 = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
    Text8 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$)))
    Text9 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "Fechco_Rep", CONDI$)))
    Text10 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$)))
    'Label49 = VALOBADA("MASTER", "Umedida", "codint= " & CODINT(TRIM(Text7)))
    Label55 = VALOBADA("LOTINGRE", "DOCUINGRE", "Cod_Inte= " & CODINT(TRIM(Text7)) & " AND IDENLOTE ='" & TRIM$(Label36) & "'")
    '
    'Valida si ya hay un Bulto ingresado con mismo Nro Lote Prov
    'y mismo Codigo de articulo.
    If TRIM$(Label51) <> "" Then
       'Extrae Fechas de LOTINGRE
       CONDI$ = "Cod_Inte=" & CIXXI% & " and LoteCol_prov = '" & TRIM$(Label51) & "' and Libera ='APRO'"
       If CantRegistros("LOTINGRE", CONDI$) > 0 Then
          'Extrae Fechas de LOTINGRE
          PORLOTAN% = 0
          If TRIM$(Text8) = "" Then
             Text8 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$)))
             PORLOTAN% = 1
          End If
          If TRIM$(Text9) = "" Then
             Text9 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "Fechco_Rep", CONDI$)))
             PORLOTAN% = 1
          End If
          If TRIM$(Text10) = "" Then
             Text10 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$)))
             PORLOTAN% = 1
          End If
          LotePropApro$ = VALOBADA("LOTINGRE", "IdenLote", CONDI$)
          '
          'Extrae Referencia de MOVISTO
          CONDI$ = "COD_INTE =" & CIXXI% & " AND IDENLOTE ='" & LotePropApro$ & "' "
          CONDI$ = CONDI$ & "AND CODIMOVI = 'CC' AND CANTINGRE > 0"
          Text6 = VALOBADA("MOVISTO", "REFERCOR", CONDI$)
          '
          If PORLOTAN% > 0 Then
             Call COMUNI("Fechas y Lotes Aplicados por Ingresos Anteriores.")
          End If
       End If
    End If
End Sub

Private Sub Label42_Change()
  Label40 = ECOARCH$("ECORDEF", Label42)
End Sub

Private Sub Text10_DblClick()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text10 = VDEV$
   End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text10_LostFocus()
   FEC = FECHANUM(TRIM(Text10))
   Text10 = FECHATEX(FEC)
End Sub

Private Sub Text7_Change()
   '
   CIXXI% = CODINT%(Text7)
   Label30 = DESCRIT0$(CIXXI%)
   Call LIMPCONMAT
      
   If ESTRAZABLE%(CIXXI%) < 1 Then
      Text8 = "": Text8.Enabled = False: Command10.Enabled = False
      Text9 = "": Text9.Enabled = False: Command11.Enabled = False
      Text10 = "": Text10.Enabled = False: Command12.Enabled = False
   Else
      Text8.Enabled = True: Command10.Enabled = True
      Text9.Enabled = True: Command11.Enabled = True
      Text10.Enabled = True: Command12.Enabled = True
   End If
   '
   'ARMA LISTA DE SELECCION DE LOTE EN CUARENTENA EN BASE AL CODIGO
   'NO SE UTILIZA EL ARMALILOT PORQUE SE SOLICITO QUE PRESENTE REMITO DE INGRRESO
   If CIXXI% > 0 Then
       CONDI$ = "COD_INTE = " & CIXXI% & " AND (LIBERA = 'CUAR' or LIBERA = '')"
       REGG& = CantRegistros&("LOTINGRE", CONDI$)
       If REGG& < 1 Then
          Call MENSERR(24, "No hay Partidas En Cuarentena\para el Código Elegido.")
          Text7.SetFocus
          Exit Sub
       End If
       CONDI$ = CONDI$ + " ORDER BY FechAlta"
       Call CARGALISEL("!LILOTRE", "LOTINGRE", "IDENLOTE/A16;FechAlta/D8;CanAlta/F13." & CNTDC$(CIXXI%) & ";(CanAlta-CanConsu)/F13." & CNTDC$(CIXXI%) & ";FechVenc/D8;DOCUINGRE/A18;LOTECOL_PROV/A16", CONDI$, "NOMESS")
       Call COPYSTRU("LILOTSE", "LILOTRE")
       Label49 = VALOBADA("MASTER", "Umedida", "codint= " & CODINT(TRIM(Text7)))
   End If
End Sub

Private Sub Text7_DblClick()
   Call Command5_Click
End Sub

Sub LIMPCONMAT()
    '
    Label36 = ""
    Label24 = ""
    Label11 = ""
    'Label52 = ""
    Text6 = ""
    Label40 = ""
    Label37 = ""
    Label38 = ""
    Label42 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Label51 = ""
    Label49 = ""
    NUCERT& = XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0")) + 1
    Label54 = NUCERT&
    Text5 = ""
    Label55 = ""
    '
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command8.SetFocus
   End If
End Sub

Private Sub Text8_DblClick()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text8 = VDEV$
   End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Text9.SetFocus
   End If
End Sub

Private Sub Text8_LostFocus()
   FEC = FECHANUM(TRIM(Text8))
   Text8 = FECHATEX(FEC)
End Sub

Private Sub Text9_DblClick()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text9 = VDEV$
   End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text10.SetFocus
   End If
End Sub

Private Sub Text9_LostFocus()
   FEC = FECHANUM(TRIM(Text9))
   Text9 = FECHATEX(FEC)
End Sub

Sub Lista_PendApro()
    '
    Screen.MousePointer = 11
    CONDI$ = "(LIBERA ='CUAR' or LIBERA = '')"
    Call CARGALISEL("PENDAPRO", "LOTINGRE+MASTER/COD_INTE=CODINT", "MASTER.CODIGO/A16;MASTER.DESCRIPCION/A24", CONDI$, "DISTINCT")
    Screen.MousePointer = 1
    '
99 End Sub

Sub ApruebaImprime(IDLOT$)
   '
   'Este procedimiento solo funciona para este formulario ya que toma variables del formulario
   
   Destino$ = TRIM$(Label42)          'Destino
   If Option1 = False Then
         Operacion$ = "RECH"
       Else
         Operacion$ = "APRO"
   End If
   '
   FFXX% = FECHANUM(Label13.Caption)
   '
   COMP$ = ProNroComprobante("CC")
   Do While Len(COMP$) < 6
     COMP$ = "0" & COMP$
   Loop
   '
   COMPROB1$ = "CC-" & AJUSTD(COMP$, 6) & "-1"
   COMPROB2$ = "CC-" & AJUSTD(COMP$, 6) & "-2"
   '
   CIXXI% = XVALO(CODINT(Text7))
   DPL% = DEPOLOTE%(CIXXI%, IDLOT$)
   '
   Call MOVISTOK(FFXX%, CIXXI%, IDLOT$, "CC", COMPROB1$, COMPROB1$, COMPROB1$, COMPROB1$, Operacion$ & ". Lote " & IDLOT$, 1, 0, 1, 0, DPL%, 0)
      '
      'Fechas
       FECFAB% = FECHANUM(TRIM$(Text8)) 'Fabricacion
       FECRET% = FECHANUM(TRIM$(Text9)) 'Retest
       FECVEN% = FECHANUM(TRIM$(Text10)) 'Vencimiento
       FECHAX = FETEXCOR1$(FFXX%)
    CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE ='" & IDLOT$ & "'"
    '
    Call ACTUREGISTRO("LOTINGRE", "LIBERA", CONDI$, Operacion$, REG&)
    Call ACTUREGISTRO("LOTINGRE", "FechFab", CONDI$, FETEXCOR1$(FECFAB%), REG&)
    Call ACTUREGISTRO("LOTINGRE", "Fechco_Rep", CONDI$, FETEXCOR1$(FECRET%), REG&)
    Call ACTUREGISTRO("LOTINGRE", "FechVenc", CONDI$, FETEXCOR1$(FECVEN%), REG&)
    Call ACTUREGISTRO("LOTINGRE", "LoteCol_Prov", CONDI$, TRIM$(Label51), REG&)
    Call ACTUREGISTRO("LOTINGRE", "Ref_Apro", CONDI$, TRIM$(Text6), REG&)
       NUCERT& = 1 + XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0"))
    Call ACTUREGISTRO("LOTINGRE", "NuCert_Apro", CONDI$, NUCERT&, REG&)
    Call ACTUREGISTRO("LOTINGRE", "Usu_Apro", CONDI$, USERNAME, REG&)
    Call ACTUREGISTRO("LOTINGRE", "Fech_Apro", CONDI$, FECHAX, REG&)
    Call ACTUREGISTRO("LOTINGRE", "Nro_Analisis", CONDI$, TRIM$(Text5), REG&)
    '
    Call ACTUREGISTRO("BALAREQMO", "NuCert_Apro", CONDI$, NUCERT&, REG&)
    '
    Call ACSALOTE(CIXXI%, IDLOT$, "NOMESS")
    '
    'Presentar Formulario en pantalla
    
    If Check1.Value = 0 Then
      TTTT$ = IDLOT$
      FORMU$ = CONTROL$("BOLRECEP", "FORAPRO")
      If TRIM(FORMU$) <> "" Then
        If EXISTE%(LUDAT$ + FORMU$ + ".SQR") > 0 Then
           Call STDFORM(FORMU$, TTTT$, "PRINT")
        End If
      End If
    End If
    '
End Sub

Sub ENVIMAILRECH(CI%, IDLOT$)
   '
   mailProgra$ = CONTROL("BOLRECEP", "PROGMAIL")
   '
   mailCompras$ = CONTROL("BOLRECEP", "COMPMAIL")
   '
   ASUNTO$ = "Rechazo de Aprobacion"
   '
   TEXTOMAI$ = "Producto Rechazado: " + CODEXT$(CI%) & " - " & DESCRIT$(CI%) & CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   NUPED$ = VALOBADA("LOTINGRE", "DOCUORIG", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Nº de Pedido: " & NUPED$ + CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   SOLFA$ = VALOBADA("BALAREQMO", "IDENREQUE", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Destino: " & SOLFA$ & CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   RERECHA$ = VALOBADA("LOTINGRE", "Ref_Apro", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Ha sido rechazado por el siguiente motivo: " & RERECHA$ & CRLF$
   '
   MAI_DESTI$ = TRIM$(mailProgra$ & "; " & mailCompras$)
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   '
   Call GENERAMAIL
   '
End Sub
Sub ENVIMAILAPROB(CI%, IDLOT$)
   'EMAIL PARA APROBACION DE RECEPCIONES
   '
   mailProgra$ = CONTROL("BOLRECEP", "PROGMAIL")
   '
   ASUNTO$ = "Aprobacion"
   '
   TEXTOMAI$ = "Producto Aprobado: " + CODEXT$(CI%) & " - " & DESCRIT$(CI%) & CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   NUPED$ = VALOBADA("LOTINGRE", "DOCUORIG", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Nº de Pedido: " & NUPED$ + CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   SOLFA$ = VALOBADA("BALAREQMO", "IDENREQUE", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Destino: " & SOLFA$ & CRLF$
   CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
   APROBA$ = VALOBADA("LOTINGRE", "Ref_Apro", CONDI$, "NOMESS")
   TEXTOMAI$ = TEXTOMAI$ & "Ha sido aprobado: " & APROBA$ & CRLF$
   '
   MAI_DESTI$ = TRIM$(mailProgra$)
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   '
   Call GENERAMAIL
   '
End Sub

