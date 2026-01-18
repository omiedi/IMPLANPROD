VERSION 5.00
Begin VB.Form CIERRAOF09 
   BackColor       =   &H00FFE0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Fabricacion"
   ClientHeight    =   5700
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
   ScaleHeight     =   5700
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFE0E0&
      Height          =   375
      Left            =   4650
      TabIndex        =   139
      Top             =   480
      Width           =   2055
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFE0E0&
         Caption         =   "Solo Devolucion"
         Height          =   195
         Left            =   120
         TabIndex        =   140
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Devolución de Materiales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3390
      Left            =   210
      TabIndex        =   31
      Top             =   1995
      Width           =   8775
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
         TabIndex        =   76
         Top             =   420
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
         Index           =   0
         Left            =   4410
         TabIndex        =   75
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   7320
         TabIndex        =   74
         Top             =   420
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
         Index           =   1
         Left            =   2520
         TabIndex        =   73
         Top             =   735
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
         Index           =   1
         Left            =   4410
         TabIndex        =   72
         Top             =   735
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   71
         Top             =   735
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
         Index           =   2
         Left            =   2520
         TabIndex        =   70
         Top             =   1050
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
         Index           =   2
         Left            =   4410
         TabIndex        =   69
         Top             =   1050
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   7320
         TabIndex        =   68
         Top             =   1050
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
         Index           =   3
         Left            =   2520
         TabIndex        =   67
         Top             =   1365
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
         Index           =   3
         Left            =   4410
         TabIndex        =   66
         Top             =   1365
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   7320
         TabIndex        =   65
         Top             =   1365
         Width           =   1305
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
         TabIndex        =   64
         Top             =   420
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
         TabIndex        =   63
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
         Index           =   2
         Left            =   5040
         TabIndex        =   62
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
         Index           =   3
         Left            =   5040
         TabIndex        =   61
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
         Index           =   6
         Left            =   5040
         TabIndex        =   60
         Top             =   2310
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
         Left            =   5040
         TabIndex        =   59
         Top             =   1995
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
         Left            =   5040
         TabIndex        =   58
         Top             =   1680
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   7320
         TabIndex        =   57
         Top             =   2310
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
         Index           =   6
         Left            =   4410
         TabIndex        =   56
         Top             =   2310
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
         Index           =   6
         Left            =   2520
         TabIndex        =   55
         Top             =   2310
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   7320
         TabIndex        =   54
         Top             =   1995
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
         Index           =   5
         Left            =   4410
         TabIndex        =   53
         Top             =   1995
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
         Index           =   5
         Left            =   2520
         TabIndex        =   52
         Top             =   1995
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   7320
         TabIndex        =   51
         Top             =   1680
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
         Index           =   4
         Left            =   4410
         TabIndex        =   50
         Top             =   1680
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
         Left            =   2520
         TabIndex        =   49
         Top             =   1680
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
         Index           =   7
         Left            =   5040
         TabIndex        =   48
         Top             =   2625
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   7320
         TabIndex        =   47
         Top             =   2625
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
         Index           =   7
         Left            =   4410
         TabIndex        =   46
         Top             =   2625
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
         Index           =   7
         Left            =   2520
         TabIndex        =   45
         Top             =   2625
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
         Index           =   8
         Left            =   5040
         TabIndex        =   44
         Top             =   2940
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   7320
         TabIndex        =   43
         Top             =   2940
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
         Index           =   8
         Left            =   4410
         TabIndex        =   42
         Top             =   2940
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
         Left            =   2505
         TabIndex        =   41
         Top             =   2940
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
         Left            =   6390
         TabIndex        =   40
         Top             =   420
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
         Left            =   6390
         TabIndex        =   39
         Top             =   735
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
         Left            =   6405
         TabIndex        =   38
         Top             =   1050
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
         Left            =   6405
         TabIndex        =   37
         Top             =   2940
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
         Left            =   6405
         TabIndex        =   36
         Top             =   2625
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
         Left            =   6405
         TabIndex        =   35
         Top             =   2310
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
         Left            =   6405
         TabIndex        =   34
         Top             =   1980
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
         Left            =   6405
         TabIndex        =   33
         Top             =   1680
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
         Index           =   3
         Left            =   6405
         TabIndex        =   32
         Top             =   1365
         Width           =   175
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   135
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   1755
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   3045
         TabIndex        =   134
         ToolTipText     =   "Código"
         Top             =   420
         Width           =   1440
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   5355
         TabIndex        =   133
         Top             =   420
         Width           =   1050
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   6720
         TabIndex        =   132
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   405
         Left            =   6090
         TabIndex        =   131
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
         TabIndex        =   130
         Top             =   450
         Width           =   1965
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         Height          =   405
         Left            =   1365
         TabIndex        =   129
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         Height          =   405
         Left            =   5055
         TabIndex        =   128
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U-M:"
         Height          =   285
         Left            =   5985
         TabIndex        =   127
         Top             =   195
         Width           =   1020
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   126
         ToolTipText     =   "Código"
         Top             =   735
         Width           =   1755
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   3045
         TabIndex        =   125
         ToolTipText     =   "Código"
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   5355
         TabIndex        =   124
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   123
         Top             =   735
         Width           =   420
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   122
         ToolTipText     =   "Código"
         Top             =   1050
         Width           =   1755
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   3045
         TabIndex        =   121
         ToolTipText     =   "Código"
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   5355
         TabIndex        =   120
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   6720
         TabIndex        =   119
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   118
         ToolTipText     =   "Código"
         Top             =   1365
         Width           =   1755
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   3045
         TabIndex        =   117
         ToolTipText     =   "Código"
         Top             =   1365
         Width           =   1440
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   5355
         TabIndex        =   116
         Top             =   1365
         Width           =   1050
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   115
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2835
         TabIndex        =   114
         Top             =   420
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   2835
         TabIndex        =   113
         Top             =   735
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   2835
         TabIndex        =   112
         Top             =   1050
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   2835
         TabIndex        =   4
         Top             =   1365
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   4725
         TabIndex        =   111
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         Height          =   195
         Left            =   2940
         TabIndex        =   110
         Top             =   195
         Width           =   1965
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   4725
         TabIndex        =   109
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   735
         Width           =   330
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   4725
         TabIndex        =   108
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1050
         Width           =   330
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   4725
         TabIndex        =   107
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1365
         Width           =   330
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   4725
         TabIndex        =   106
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   2310
         Width           =   330
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   4725
         TabIndex        =   105
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1995
         Width           =   330
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   4725
         TabIndex        =   104
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   2835
         TabIndex        =   103
         Top             =   2310
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   2835
         TabIndex        =   102
         Top             =   1995
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   2835
         TabIndex        =   101
         Top             =   1680
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   6720
         TabIndex        =   100
         Top             =   2310
         Width           =   420
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   5355
         TabIndex        =   99
         Top             =   2310
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   3045
         TabIndex        =   98
         ToolTipText     =   "Código"
         Top             =   2310
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   97
         ToolTipText     =   "Código"
         Top             =   2310
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   6720
         TabIndex        =   96
         Top             =   1995
         Width           =   420
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   5355
         TabIndex        =   95
         Top             =   1995
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   3045
         TabIndex        =   94
         ToolTipText     =   "Código"
         Top             =   1995
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   5
         Left            =   840
         TabIndex        =   93
         ToolTipText     =   "Código"
         Top             =   1995
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   6720
         TabIndex        =   92
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   5355
         TabIndex        =   91
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   3045
         TabIndex        =   90
         ToolTipText     =   "Código"
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   89
         ToolTipText     =   "Código"
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   4725
         TabIndex        =   88
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   2625
         Width           =   330
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   2835
         TabIndex        =   87
         Top             =   2625
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   6720
         TabIndex        =   86
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   5355
         TabIndex        =   85
         Top             =   2625
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   3045
         TabIndex        =   84
         ToolTipText     =   "Código"
         Top             =   2625
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   7
         Left            =   840
         TabIndex        =   83
         ToolTipText     =   "Código"
         Top             =   2625
         Width           =   1755
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   4725
         TabIndex        =   82
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   2940
         Width           =   330
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   2835
         TabIndex        =   81
         Top             =   2940
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   6720
         TabIndex        =   80
         Top             =   2940
         Width           =   420
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   5355
         TabIndex        =   79
         Top             =   2940
         Width           =   1050
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   3045
         TabIndex        =   78
         ToolTipText     =   "Código"
         Top             =   2940
         Width           =   1440
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   8
         Left            =   840
         TabIndex        =   77
         ToolTipText     =   "Código"
         Top             =   2940
         Width           =   1755
      End
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
      Left            =   6570
      TabIndex        =   30
      Top             =   1380
      Width           =   175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5505
      TabIndex        =   28
      Top             =   1380
      Width           =   1080
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1365
      TabIndex        =   25
      Top             =   1365
      Width           =   540
   End
   Begin VB.CommandButton Command30 
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
      Left            =   1905
      TabIndex        =   24
      Top             =   1365
      Width           =   175
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFE0E0&
      Caption         =   "Cantidad"
      Height          =   1590
      Left            =   6810
      TabIndex        =   16
      Top             =   105
      Width           =   2175
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   850
         TabIndex        =   137
         Top             =   1215
         Width           =   1200
      End
      Begin VB.TextBox Tlabel38 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   850
         TabIndex        =   23
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "U/Bultos:"
         Height          =   300
         Left            =   -525
         TabIndex        =   138
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   855
         TabIndex        =   21
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   855
         TabIndex        =   20
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realiz:"
         Height          =   300
         Left            =   -525
         TabIndex        =   19
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyect:"
         Height          =   300
         Left            =   -525
         TabIndex        =   18
         Top             =   345
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acumul:"
         Height          =   300
         Left            =   -525
         TabIndex        =   17
         Top             =   660
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
      Height          =   5745
      Left            =   9135
      ScaleHeight     =   5685
      ScaleWidth      =   1320
      TabIndex        =   5
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command15 
         Height          =   600
         Left            =   120
         Picture         =   "CIERRAOF09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1050
         Width           =   645
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   1005
         Left            =   120
         Picture         =   "CIERRAOF09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Guardar Cierre de O/F y Reingreso de Materiales a O/F"
         Top             =   3750
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
         TabIndex        =   6
         Top             =   4845
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
            TabIndex        =   7
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   120
         Picture         =   "CIERRAOF09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -285
         Picture         =   "CIERRAOF09.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5055
         Width           =   1515
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion:"
      Height          =   225
      Left            =   4320
      TabIndex        =   29
      Top             =   1455
      Width           =   1080
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      Height          =   270
      Left            =   300
      TabIndex        =   27
      Top             =   1410
      Width           =   915
   End
   Begin VB.Label Label38 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2145
      TabIndex        =   26
      Top             =   1380
      Width           =   2205
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
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   1050
      Width           =   1125
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   13
      Top             =   975
      Width           =   5370
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1365
      TabIndex        =   12
      Top             =   585
      Width           =   1965
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   300
      Left            =   -315
      TabIndex        =   11
      Top             =   660
      Width           =   1545
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
      Height          =   510
      Left            =   4335
      TabIndex        =   10
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5565
      TabIndex        =   9
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Fabr. Nº:"
      Height          =   300
      Left            =   105
      TabIndex        =   8
      Top             =   270
      Width           =   1125
   End
End
Attribute VB_Name = "CIERRAOF09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%
Dim NUELOTESS$(8)

Private Sub command1_Click()
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



Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Label19_Click()

End Sub

Private Sub Command10_Click(Index As Integer)
    DEPO% = XVALO(Label52(Index))
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

Private Sub Command15_Click()
   Command15.Enabled = False
   FIKARDEX07.Show 1
   Command15.Enabled = True
End Sub

Private Sub Command2_Click(Index As Integer)
 If Label32 <> "" Then
    Call SELECHO("TADEPOSI")
    NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
    Label52(Index) = NUMDEP%
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

Private Sub Label28_Change()
   CIXI% = CODINT%(Label28)
   'Text6 = TRIM$(Str(DEPOPRE%(CIXI%)))  'EL DEPOSITO LO DEBE SELECCIONAR EL USUARIO
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("TAUBICA")
    
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text2 = UBISELE$
    End If
End Sub

Private Sub Text6_Change()
   Label38 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text6)))
   Text2 = ""
End Sub

Private Sub Text6_DblClick()
   Call Command9_Click
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call SELECHO("MASTER")
    vdevu$ = VALACT1$("MASTER")
    If vdevu$ <> "" Then
       CICOMPO% = CODINT%(vdevu$)
       CACONSU = 0
       TTYY$ = Space$(128)
       Call REPLA(TTYY$, vdevu$, 1, 16)
       Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
       Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13." + CNTDC$), 57, 13)
       On Error Resume Next
       USOUNI = CACONSU / XVALO(Text1)
       On Error GoTo 0
       Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F16." + CNTDC$), 97, 16)
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
       vdevu$ = VALACT1$("MASTER")
       If vdevu$ <> "" Then
          CICOMPO% = CODINT%(vdevu$)
          CACONSU = 0
          TTYY$ = Space$(128)
          Call REPLA(TTYY$, vdevu$, 1, 16)
          Call REPLA(TTYY$, DESCRIT$(CICOMPO%), 17, 40)
          Call REPLA(TTYY$, FORMATNUM$(CACONSU, "F13." + CNTDC$), 57, 13)
          On Error Resume Next
          USOUNI = CACONSU / XVALO(Text1)
          On Error GoTo 0
          Call REPLA(TTYY$, FORMATNUM$(USOUNI, "F13." + CNTDC$), 97, 16)
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
    If Check1.Value = 1 Then GoTo 50
    '
    If Label32 = "" Then
       Call MENSERR(24, "Número de O/F no Válido")
       GoTo 99
    End If
    ' VALIDAR FECHA
    HOII% = HOY(HO$)
    FECHA% = FECHANUM(Label24)
    If FECHA% < 1 Or FECHA% > HOII% Then
       Call MENSERR(24, "Fecha de Lanzamiento No Válida")
       GoTo 99
    End If
    '
    'VALIDAR CANTIDAD
    VALONU = XVALO(Tlabel38)
    If VALONU < 1 Then
       Call MENSERR(24, "Cantidad No Válida")
       GoTo 99
    End If
    'Validar Ubicación
    POSI$ = TRIM$(Text2)
    CI% = CODINT(Label28)
    If PERMITEUBI%(CI%, "", POSI$) <> 1 Then
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
    If XVALO(Text3) <= 0 Then
       Call MENSERR(24, "Falta Definir Unidades por Bulto")
       GoTo 99
    End If
    '
    'Validacion de Items Sobrantes
    '
    'Valida Codigo <>"" en registro con cantidad > 0
50  For U& = 0 To Text1.Count
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label39(U&)) = "" Then
             Call MENSERR(24, "Componente sin Código de Material")
             GoTo 99
          End If
       End If
    Next U&
    '
    For U& = 0 To 8
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label40(U&)) = "" Then
             Call MENSERR(24, "Componente sin Lote / Partida")
             GoTo 99
          End If
       End If
    Next U&
    '
    For U& = 0 To 8
       If XVALO(Text1(U&)) > 0 Then
          If XVALO(Label52(U&)) < 1 Then
             Call MENSERR(24, "Falta Indicar Depósito de Entrada Devolución")
             GoTo 99
          End If
       End If
    Next U&
    '
    'Valida cantidad devuelta sea mayor a cero
    For U& = 0 To 8
      CIXI% = CODINT%(TRIM$(Label39(U&)))
      If CIXI% > 1 Then
        If XVALO(Text1(U&)) < FRACMIN(CIXI%) Then
             Call MENSERR(24, "Falta Indicar Cantidad a Devolver para el Lote " & TRIM$(Label40(U&)) & ".")
             GoTo 99
        End If
      End If
    Next U&
    'Valida si saldo de lote es => a cantidad
    For U& = 0 To 8
       If XVALO(Text1(U&)) > 0 Then
          CIXI% = CODINT(Label39(U&))
          IDLOT$ = TRIM$(Label40(U&))
          NUMDEP% = VALOBADA("lotingre", "DEPOUBIC", "idenlote='" & IDLOT$ & "'")
          CANSA = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
          If CANSA < XVALO(Text1(U&)) Then
             COMUNI ("Cantidad Devuelta Supera a Cantidad en Lote, en Fila Nro " & U& + 1)
             GoTo 99
          End If
       End If
    Next U&
    
    'Valida Ubicacion
    For U& = 0 To 8
       If XVALO(Text1(U&)) > 0 Then
          If TRIM$(Label41(U&)) = "" Then
             COMUNI ("Falta Ingresar Ubicacion en Fila Nro " & U& + 1)
             GoTo 99
          End If
          POSI$ = TRIM$(Label41(U&))
          CIXI% = CODINT(Label39(U&))
          If PERMITEUBI%(CIXI%, "", POSI$) <> 1 Then
             GoTo 99
          End If
        End If
    Next U&
    '
    If Check1.Value = 0 Then
        Call GRACIERROF
    End If
    '
    Call GRABACONSUM
    If Check1.Value = 1 Then GoTo 60
    '
    REALIZAD = XVALO(Tlabel38)
    CANXCAJ = XVALO(Text3)
    CANETIQU = Int(REALIZAD / CANXCAJ)
    RESTO2# = REALIZAD - (CANETIQU * CANXCAJ)
    If RESTO2 > 0 Then
     CANETIQU = CANETIQU + 1
    End If
    '
    If CANETIQU < 1 Then CANETIQU = 1
    '
    CICONJ% = CODINT%(Label28)
    LOTE$ = "OP." & Label32
    '
    If CANETIQU > 1 Then
        For I& = 1 To CANETIQU - 1 ' IMPRIME TODAS MENOS LA ULTIMA
         Call PRINSTIKAPRO1(CICONJ%, LOTE$, DEPENTRA%, CANXCAJ)
        Next I&
        If RESTO2 < 0.5 Then RESTO2 = CANXCAJ
        Call PRINSTIKAPRO1(CICONJ%, LOTE$, DEPENTRA%, RESTO2)    ' IMPRIME LA ULTIMA  ETIQUETA
    Else
        Call PRINSTIKAPRO1(CICONJ%, LOTE$, DEPENTRA%, CANXCAJ)
    End If
    '
60  For I& = 0 To 8
      If CODINT%(Label39(I&)) > 0 Then
        CIXI% = CODINT%(Label39(I&))
        DEP% = XVALO(Label52(YX&))
        LOTEESS$ = NUELOTESS$(I&)
        Call PRINSTIKAPRO1(CIXI%, LOTEESS$, DEP%)      ' IMPRIME ETIQUETAS DE DEVOLUCIONES
      End If
    Next I&

    Check1.Value = 0
        
    
'    'Actualiza el estado de la Ubicacion
'    For U& = 0 To 8
'     UBIC$ = TRIM(Label41(U&))
'     LOTE$ = Label40(U&)
'     'CIXI% = CODINT(Label39(U&))
'     If UBIC$ <> "" Then
'        CONDI$ = "IDENLOTE ='" & LOTE$ & "' and Cod_Inte=" & CIXI%
'        Call ACTUREGISTRO("UBIPRE", "Status", "codiUbic = '" & UBIC$ & "'", 1, nroREG&)
'        Call ACTUREGISTRO("LOTINGRE", "UbicFis", CONDI$, UBIC$, REG&)
'
'        CODI% = CODINT(Label39(U&))
'        CONDI2$ = "IDENLOTE ='" & LOTE$ & "' AND COD_INTE= " & CODI%
'
'        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI2$, "", REG&)
'     End If
'    Next U&
    '
    'Esto se comenta
'    CICONJ% = CODINT%(Label28)
'    LOTE$ = Label32
'    Call PRINSTIKAPRO(CICONJ%, LOTE$, DEPENTRA%)
    '
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
     CIERRA% = 0: If XVALO(Tlabel38) + XVALO(Label37) > XVALO(Label34) Then CIERRA% = 1
     If XVALO(Tlabel38) + XVALO(Label37) < XVALO(Label34) Then
          OPX% = COMALTER%("ATENCION: La OP." & NORFA$ & " \queda con Saldo Pendiente.\\Mantener Abierta\Cerrar")
          If OPX% = 2 Then CIERRA% = 1
     End If
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM  ORDEFABR "
     'SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
     SQLX$ = SQLX$ + " WHERE NuOrProd = '" & NORFA$ & "'"
     Dim ABREORFFFAB As ADODB.Recordset
     Set ABREORFFFAB = New ADODB.Recordset
     ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
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
     If CANTI >= 0.05 Then
        DEPX% = XVALO(Text6)
        If DEPX% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
        NUORIT% = 0
        FECHA% = FECHANUM(Label24)
        CICONJ% = CODINT%(Label28)
        LOTE$ = "OP." & Label32
        CMOV$ = "OP"
        DOPRI$ = "OP." & NORFA$
        CONDI$ = "NUORPROD = '" & NORFA$ & "'"
        DOSEC$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")  ' OBTIENE LA S/F ASOCIADA A LA ORDEN DE PRODUCCION
        REFE$ = REFEOF$
        VUNI# = 0
        MONEII% = 1
        NC% = 0: UBIK$ = Left$(Text2, 32)
        CONDI$ = "IdenLote = '" & LOTE$ & "'"
        REGA& = CantRegistros&("LOTINGRE", CONDI$)
        If REGA& < 1 Then
           LOTEPRO$ = SAFETEXT(VALOBADA("ORDEFABR", "IDENLOTE", "IDSUBOR='" & DOSEC$ & "'"))
           Call CREALOTIN(FECHA%, FECHA%, FECHA%, CICONJ%, LOTE$, "", LOTEPRO$, CDbl(CANTI), 0, 0, DEPX%, UBIK$, "", "", "", "CUAR")
        End If
        Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI)
     
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
     DOPRI$ = "OP." & Label32
     DOSEC$ = "OP." & Label32
     REFE$ = "Sobrante " + Label32
     NUORIT% = 0
     '
     XXPO% = 65
     '
     For YX& = 0 To 8
        If TRIM$(Label39(YX&)) <> "" Then
            CIXI% = CODINT%(Label39(YX&))
            DEP% = XVALO(Label52(YX&))
            UBIC$ = TRIM$(Label41(YX&))
            LOTE$ = Label40(YX&)
            CMO$ = "CF"    ' CONSUMO DE FABRICACION
            CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
            FEFABRI% = CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$, "NOMESS"))
            FEVENCI% = CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$, "NOMESS"))
            LOTEPROV$ = VALOBADA("LOTINGRE", "LoteCol_Prov", CONDI$, "NOMESS")
            PROVEE$ = VALOBADA("LOTINGRE", "ORIGEN", CONDI$, "NOMESS")
            DOCINGRE$ = VALOBADA("LOTINGRE", "DOCUINGRE", CONDI$, "NOMESS")
            '
            CANTIDD# = XVALO(Text1(YX&))
            PPXX% = InStr(LOTE$, "/")
            If PPXX% > 0 Then LOTE$ = Left$(LOTE$, PPXX% - 1)
            If PPXX% = 0 Then XXPO% = 65
            For KKII% = XXPO% To 90
               NUELOTE$ = LOTE$ + "/" + Chr$(KKII%)
               CONDI$ = "IDENLOTE = '" & NUELOTE$ & "'"
               CANTA& = CantRegistros&("LOTINGRE", CONDI$)
               If CANTA < 1 Then
                  LOTTEE$ = NUELOTE$
                  GoTo 30
               End If
               XXPO% = KKII% + 1
            Next KKII%
            '
30          Call MOVISTOK(FF%, CIXI%, LOTTEE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, 0, DEP%, CANTIDD#)
            Call CREALOTIN(FF%, FEFABRI%, FEVENCI%, CIXI%, LOTTEE$, "", LOTEPROV$, CANTIDD#, 0, 0, DEP%, UBIC$, PROVEE$, DOCINGRE$, "", "CUAR")
            Call ACSALOTE(CIXI%, LOTTEE$, "NOMESS")
            NUELOTESS$(YX&) = LOTTEE$
         End If
     Next YX&
     
'     For YX& = 0 To 2
'        If NLOTX$ <> "" Then
'           NUORIT% = NUORIT% + 1
'           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DEPORIG%, CANTI)                         '\\\OT-06-0244-WAR-20/06/06///
'        End If
'     Next YX&
'     '
'     For YX& = 0 To 3
'        CODEX$ = TRIM$(Label39(YX&))
'        If CODEX$ <> "" Then
'           CIXI% = CODINT%(CODEX$)
'           NUORIT% = NUORIT% + 1
'           LOTE$ = Label40(YX&)
'           CANTI = XVALO(Text1(YX&))
'           DEPORIG% = XVALO(Label52(YX&))
'           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DEPORIG%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
'
'        End If
'     Next YX&
     '
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
        Y$ = REGBLAN("ECORDEP")
        Call REPLA(Y$, !DOPRILAR, 1, 12)
        FEX = CVINT(!FechMov)
        Call REPLA(Y$, FECHATEX$(FEX) + "  " + !ReferCor, 13, 52)
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", Y$, JJ&)
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
        LOTE$ = !idenLote
        CIXI% = Cod_Inte
        !cantsalid = 0
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
    vdevu$ = VALACT1$("MASTER")
    If vdevu$ <> "" Then
       CICOMPO% = CODINT%(vdevu$)
       CACONSU = 0
       Text4 = FORMATNUM$(CACONSU, "F10." + CNTDC$)
    End If
    Command17.Enabled = True
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
        vdevu$ = VALACT1$("!MASREQ")
'        Call SELECHO("MASTER")
'        VDEVU$ = VALACT1$("MASTER")
        If vdevu$ <> "" Then
          CICOMPO% = CODINT%(vdevu$)
          Label39(Index) = vdevu$
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

Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Text6 = VALACT1$("TADEPOSI")
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
         FORHLPDF.Pdf1.setShowToolbar (True)
         On Error Resume Next
         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
         If Err Then GoTo 20
         On Error GoTo 0
         FORHLPDF.Show 1
         GoTo 99
      End If
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
    'Call ACTECOR(2)     ' para seleccionar solo las activas lanzadas
    Call CIERROFAB
    Call CommSel("ECORDEP/Indice de Ordenes de Fabricación Pendientes", Command5, Label32)
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
Private Sub Command6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command6_Click(Index)
   End If
End Sub

Private Sub Command7_Click(Index As Integer)
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
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
     xx$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(xx$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(xx$, 4))
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
    '
    '
    '
    Call SELECHO("!LILOTSE")
    vdevu$ = VALACT1$("!LILOTSE")
    VDEV2$ = VALACT2$("!LILOTSE")
    '

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
    Label37 = ""
    Tlabel38 = ""
    Label6 = ""
    Label16 = ""
    Text4 = ""
    '
   
    For KKU& = 0 To 8
       Label39(KKU&) = ""
       Text1(KKU&) = ""
       Label40(KKU&) = ""
       Label41(KKU&) = ""
       Label42(KKU&) = ""
       Label52(KKU&) = ""
    Next KKU&
    '
  
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
           Label34 = FORMATNUM$(!CanProy, "F10." + CNTDC$)
           Label37 = FORMATNUM$(!CanApro, "F10." + CNTDC$)
           Tlabel38 = FORMATNUM$(CAPENDE#, "F10." + CNTDC$)
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


Private Sub Label39_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label40_Change(Index As Integer)
    Call MUESTRASEL1(Index)
End Sub

Private Sub Label52_Change(Index As Integer)
    Call MUESTRASEL1(Index)
    Label41(Index) = ""
End Sub


Private Sub Text1_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 8
       If Label39(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text1(KKI%))
         Else
           Text1(KKI%) = ""
       End If
    Next KKI%
    Label49 = TRIM$(FORMATNUM$(SUMATRA, "F10." + CNTDC$))
    If XVALO(Text1(Index)) < 0 Or TRIM$(Text1(Index)) = "" Then
       Label21(Index) = ""
       Label39(Index) = ""
       Label52(Index) = ""
       Label42(Index) = ""
       If Index < 9 Then
          Label40(Index) = ""
          Label41(Index) = ""
       End If
    End If
End Sub

Private Sub text1_lostfocus(Index As Integer)
    Text1(Index) = FORMATNUM(XVALO(Text1(Index)), "F10." + CNTDC$)
End Sub





Private Sub Text4_KeyPress(KeyAscii As Integer)

    If KeyAscii% = 13 Then
       Call Text4_LostFocus
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




Private Sub Text4_LostFocus()
   '
   Text4 = TRIM$(FORMATNUM$(Text4, "F12." + CNTDC$))
   '
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
   CANSA = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
   LARGX = 0: If ANCHX > 0 Then LARGX = CANSA / ANCHX
    '
End Sub

Sub MUESTRASEL1(Index)
   '
   If Index < 4 Then
      CIXI% = CODINT%(TRIM$(Label39(Index)))
      IDLOT$ = TRIM$(Label40(Index))
      NUMDEP% = XVALO(Label52(Index))
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


