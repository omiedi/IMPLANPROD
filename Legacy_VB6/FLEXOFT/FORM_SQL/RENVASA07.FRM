VERSION 5.00
Begin VB.Form RENVASA07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Envasado"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   8400
      Picture         =   "RENVASA07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Graba - Aprueba la Transacción"
      Top             =   5460
      Width           =   830
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   3570
      TabIndex        =   54
      Top             =   2835
      Width           =   2325
      Begin VB.Label Label36 
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
         Left            =   630
         TabIndex        =   60
         Top             =   1260
         Width           =   1275
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
         Left            =   630
         TabIndex        =   59
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label38 
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
         Left            =   630
         TabIndex        =   58
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
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
         Left            =   105
         TabIndex        =   57
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
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
         Left            =   105
         TabIndex        =   56
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso:"
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
         Left            =   105
         TabIndex        =   55
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   5985
      TabIndex        =   47
      Top             =   2835
      Width           =   2325
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
         Left            =   210
         TabIndex        =   53
         Top             =   1890
         Width           =   1905
      End
      Begin VB.Label Label40 
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
         Left            =   630
         TabIndex        =   52
         Top             =   630
         Width           =   1275
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
         Left            =   630
         TabIndex        =   51
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label10 
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
         Height          =   300
         Left            =   105
         TabIndex        =   50
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label12 
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
         Height          =   300
         Left            =   105
         TabIndex        =   49
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
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
         Left            =   0
         TabIndex        =   48
         Top             =   1575
         Width           =   1035
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   105
      TabIndex        =   32
      Top             =   5355
      Width           =   8205
      Begin VB.TextBox Text4 
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
         Left            =   3780
         MaxLength       =   12
         TabIndex        =   70
         Top             =   2310
         Width           =   1260
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
         Left            =   1155
         TabIndex        =   66
         Top             =   1890
         Width           =   2160
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
         Height          =   285
         Left            =   4860
         TabIndex        =   61
         Top             =   630
         Width           =   175
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
         Height          =   300
         Left            =   3120
         TabIndex        =   68
         Top             =   2310
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
         Left            =   1155
         TabIndex        =   67
         Top             =   2310
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
         Left            =   3810
         MaxLength       =   12
         TabIndex        =   69
         Top             =   1680
         Width           =   1260
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8205
         Y1              =   440
         Y2              =   440
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Envase:"
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
         Left            =   105
         TabIndex        =   72
         Top             =   1935
         Width           =   930
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
         Left            =   420
         TabIndex        =   65
         Top             =   1000
         Width           =   4630
      End
      Begin VB.Label Label44 
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
         TabIndex        =   64
         Top             =   750
         Width           =   1125
      End
      Begin VB.Line Line1 
         X1              =   5250
         X2              =   5250
         Y1              =   120
         Y2              =   2820
      End
      Begin VB.Label Label43 
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
         Left            =   1080
         TabIndex        =   63
         Top             =   660
         Width           =   1545
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
         Left            =   2730
         TabIndex        =   62
         Top             =   630
         Width           =   2220
      End
      Begin VB.Line Line3 
         X1              =   8190
         X2              =   8190
         Y1              =   120
         Y2              =   2820
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nuevo Lote                                                                    Lote Remanente"
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
         Left            =   0
         TabIndex        =   46
         Top             =   120
         Width           =   8265
      End
      Begin VB.Line Line4 
         X1              =   15
         X2              =   15
         Y1              =   120
         Y2              =   2820
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   8200
         Y1              =   2820
         Y2              =   2820
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
         Left            =   3570
         TabIndex        =   45
         Top             =   2050
         Width           =   810
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Left            =   315
         TabIndex        =   44
         Top             =   1470
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
         Left            =   3570
         TabIndex        =   43
         Top             =   1420
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
         TabIndex        =   42
         Top             =   2360
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
         Left            =   1155
         TabIndex        =   41
         Top             =   1470
         Width           =   2130
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   5985
         TabIndex        =   40
         Top             =   1680
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
         Left            =   5460
         TabIndex        =   39
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
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
         Left            =   5985
         TabIndex        =   38
         Top             =   1260
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
         Left            =   5355
         TabIndex        =   37
         Top             =   2050
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
         Left            =   5775
         TabIndex        =   36
         Top             =   840
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
         Left            =   6825
         TabIndex        =   35
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
         Left            =   6825
         TabIndex        =   34
         Top             =   1680
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
         Left            =   5880
         TabIndex        =   33
         Top             =   2310
         Width           =   2140
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   105
      TabIndex        =   24
      Top             =   2835
      Width           =   3375
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
         Left            =   1995
         TabIndex        =   31
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   840
         TabIndex        =   30
         Top             =   1080
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
         Left            =   735
         TabIndex        =   29
         Top             =   1890
         Width           =   2325
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Envase:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1575
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote / Partida:"
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
         Left            =   420
         TabIndex        =   27
         Top             =   315
         Width           =   1335
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
         Left            =   1470
         TabIndex        =   26
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label35 
         Height          =   330
         Left            =   210
         TabIndex        =   25
         Top             =   1050
         Visible         =   0   'False
         Width           =   855
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
      TabIndex        =   20
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
      Left            =   3600
      TabIndex        =   14
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   105
      TabIndex        =   4
      Top             =   1050
      Width           =   8205
      Begin VB.Frame Frame1 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   3930
         TabIndex        =   19
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1220
         Left            =   2880
         TabIndex        =   18
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   6900
         TabIndex        =   17
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame3"
         Height          =   1170
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   1220
         Left            =   5760
         TabIndex        =   8
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         ItemData        =   "RENVASA07.frx":030A
         Left            =   630
         List            =   "RENVASA07.frx":030C
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   6945
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote                  F.Venc.   Ubicacion             kilos"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   630
         TabIndex        =   7
         Top             =   315
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5490
      Left            =   8400
      ScaleHeight     =   5430
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   -120
      Width           =   1380
      Begin VB.CommandButton Command8 
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
         Height          =   690
         Left            =   105
         Picture         =   "RENVASA07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1785
         Width           =   690
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
         Left            =   105
         Picture         =   "RENVASA07.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1050
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   2
         Top             =   4725
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
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
         Left            =   105
         Picture         =   "RENVASA07.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Salir"
         Top             =   300
         Width           =   690
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RENVASA07.frx":0A6C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5040
         Width           =   1515
      End
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
      TabIndex        =   22
      Top             =   270
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
      TabIndex        =   21
      Top             =   210
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
      TabIndex        =   12
      Top             =   690
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
      TabIndex        =   11
      Top             =   630
      Width           =   6690
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
      Left            =   1470
      TabIndex        =   10
      Top             =   210
      Width           =   2220
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
      TabIndex        =   9
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "RENVASA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()

   Command2.Enabled = False
   '
   'VALIDACIONES
   FECHA% = FECHANUM(Label13)
   FEFA% = FECHANUM(Label36)
   FEVE% = FECHANUM(Label38)
   CIXI% = CODINT%(Label28)
   IDLOTX$ = TRIM$(Label27)
   DEPO0% = XVALO(Text6)
   '
   If TRIM$(Label39) = "" Then
      Call MENSERR(24, "Falta Ubicación del Lote Original.") ' Valida Ubicacion
      GoTo 99
   End If
   '
   If CIXI% < 1 Then
      Call MENSERR(24, "Codigo no Válido.") ' Valida Codigo interno
      GoTo 99
   End If
   '
   If IDLOTX$ = "" Then
      Call MENSERR(24, "Falta Seleccionar Lote a Fraccionar.") ' Valida Lote seleccionado
      GoTo 99
   End If
   '
   If XVALO(Text1.TEXT) <= 0 Then ' Verifica que la cantidad de kilos sea valida
      Call MENSERR(24, "Kilos no Valido Nuevo Lote.")
      GoTo 99
   End If
   '
   If XVALO(Label32) < 0 Then ' Verifica que la cantidad de kilos sea valida
      Call MENSERR(24, "Kilos no Valido en Lote Remanente.")
      GoTo 99
   End If
   '
   UBICAIN$ = Mid$(TRIM$(UCase(Text3)), 1, 16) ' Valida Ubicacion Envase 1
   If UBICAIN$ <> "" Then
      If TRIM$(VALOBADA("UBIPRE", "Codiubic", "Codiubic = '" & UBICAIN$ & "'")) = "" Then
        Call MENSERR(24, "Ubicacion " + TRIM$(UBICAIN$) + " Inexistente")
        GoTo 99
      End If
      '
      CODIPROI% = XVALO(VALOBADA("LOTINGRE", "Cod_Inte", "UbicFis = '" & UBICAIN$ & "' AND CANALTA > (CANCONSU + 0.05) ORDER BY FECHALTA ASC "))
      If CODIPROI% > 0 Then
        CODEPRO$ = CODEXT$(CODIPROI%)
        OPX% = COMALTER%("La Ubicacion " + UBICAIN$ + " se Encuentra Ocupada\por Código '" + CODEPRO$ + "'.\  \¿Confirma la Posicion Elegida?\\No, Cambiar\Si, Continuar")
        '
        If OPX% = 1 Then
             GoTo 99
        End If
      End If
   Else
      Call MENSERR(24, "Falta Ingreso de Ubicación en Envase 1")
      GoTo 99
   End If
   '
   CANTINUE = XVALO(Text4)
   If CANTINUE < 0.05 Then
       Call MENSERR(24, "Cantidad No Válida en Nuevo Lote.")
       GoTo 99
   End If
   '
   If Text2.TEXT = "" Then
      Call MENSERR(24, "Envase Nuevo Lote no Valido.")
      GoTo 99
   End If
   '
   CODINU$ = TRIM$(Label42)
   CIXJ% = CODINT%(CODINU$)
   If CODINU$ = "" Or CIXJ% < 1 Then
      Call MENSERR(24, "Código no Válido en Nuevo Lote.")
      GoTo 99
   End If
   '
   DPRI$ = TRIM$(Str$(1 + XVALO(Mid$(VALOBADA("MOVISTO", "MAX(DOPRILAR)", "CodiMovi = 'RE'"), 4, 6))))
   While Len(DPRI$) < 6: DPRI$ = "0" + DPRI$: Wend
   DPRI$ = "RE-" + DPRI$
   HOII% = HOY(HOS$)
   DSEC$ = "RE-" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   For KKI% = 0 To 99
      KKX$ = TRIM$(Str$(KKI%))
      While Len(KKX$) < 2: KKX$ = "0" + KKX$: Wend
      KKY$ = VALOBADA("MOVISTO", "DOSECLAR", "DOSECLAR = '" + DSEC$ + "-" + KKX$ + "'")
      If KKY$ = "" Then
         DSEC$ = DSEC$ + "-" + KKX$
         GoTo 25
      End If
   Next KKI%
   Call MENSERR(24, "Imposible Procesar en Fecha Elegida - Cambie Fecha Movimiento.")
   GoTo 99
   '
25 SQLX$ = "Select * from LOTINGRE WITH(NOLOCK) where IdenLote = '" & IDLOTX$ & "'" 'Selecciona el registro correspondiente al LOTE
   Dim LOTINGTMP As ADODB.Recordset
   Set LOTINGTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   With LOTINGTMP
      If Not (.BOF And .EOF) Then
         CCIIN% = !Cod_Inte
         REMIPROX$ = SAFETEXT$(!DocuIngre)
         FECHALTT% = CVINT(!FechAlta)
         ENVASX$ = SAFETEXT$(!CodEnvase)
         FECHFABR% = CVINT(!FechFab)
         FECHAVEN% = CVINT(!FechVenc)
         MARCAX$ = SAFETEXT$(!Marca)
         ORIGENIN$ = SAFETEXT$(!Origen)
         LOTEINGR$ = SAFETEXT$(!LoteCol_Prov)
         STALIB$ = SAFETEXT$(!Libera)
         GoTo 35
      End If
      Call MENSERR(24, "Datos No Procesables - Consulte.")
      GoTo 99
   End With
35 LOTINGTMP.Close
   '
   Call MOVISTOK(FECHA%, CIXI%, TRIM$(Label27), "RE", DPRI$, DPRI$, DSEC$, DSEC$, "A " + Label42, 0, 0, 0, 0, DEPX%, (-1) * XVALO(Text1))
   Call MOVISTOK(FECHA%, CIXJ%, TRIM$(Label6), "RE", DPRI$, DPRI$, DSEC$, DSEC$, "De " + Label28, 0, 0, 0, 0, DEPX%, XVALO(Text1))
   '
   Call CREALOTIN(FECHA%, FECHFABR%, FECHAVEN%, CIXJ%, Label6, _
           MARCAX$, LOTEINGR$, XVALO(Text1), , , DEPX%, Text3, ORIGENIN$, _
           REMIPROX$, TRIM$(Text2), STALIB$, XVALO(Text4))
           '
   Call ACSALOTE(CIXI%, TRIM$(Label27), "NOMESS")
   Call ACSALOTE(CIXJ%, TRIM$(Label6), "NOMESS")
   '
   Call LEARTRA07.GENTAUBI
   Call FRESHECHO("TAUBICA")
   '
   Call COMUNI("Procesado")
   '
   Call LIMPLAB
   Call LIMPTEXT
   Call Label28_Change
 
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   '
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
    Call SELECHO("TAUBICA")
    UBISELE$ = VALACT1$("TAUBICA")
    If UBISELE$ <> "" Then
       Text3 = UBISELE$
    End If
End Sub

Private Sub Command6_Click()

   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text12.TEXT = VDEV$
   End If
   
End Sub

Private Sub Command5_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label28 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command7_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Label42 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   FIKARDEX07.Show 1
   Command8.Enabled = True
End Sub
'
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

Private Sub Command9_Click()
    Call CommSel("TADEPOSI", Command9, Text6)
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    HOII% = HOY(HOS$)
    Label13 = HOS$
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       OPENFORMS = DoEvents
    End If
    '
End Sub

Sub VERIPERPRUE()
    '
    If CONTROL$("TRAZABI", "FINICIO") = "" Then
       HOII% = HOY(HOS$)
       Call GRACONTROL("TRAZABI", "FINICIO", HOS$)
    End If
    '
    FINICIO = FECHANUM(CONTROL$("TRAZABI", "FINICIO"))
    FECHACT = CVINT(FileDateTime(LUDAT$ + "MONIEXEC.DAT"))
    '
    Dim BB As String * 128
    NX1% = FILEOPEN(LUDAT$ + "MONIEXEC.DAT", 0, 128)
    FINI& = LOF(NX1%) / 128
    COMI& = FINI& - 128: If COMI& < 1 Then COMI& = 1
    '
    For U& = COMI& To FINI&
       Get #NX1%, U&, BB$
       FEJECU = FECHANUM(Mid$(BB$, 69, 8))
       If FEJECU > FECHACT Then
          FECHACT = FEJECU
       End If
    Next U&
    '
    If DIENTRE%(FINICIO, FECHACT) > 7 Then
       Call MENSERR(24, "Plazo de Verificación Vencido.")
       Call FINAL("")
    End If
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Hide
End Sub



Private Sub Label24_Click()

    Label17.Caption = Label24.Caption
    
End Sub



Private Sub Label28_Change()
    CIXI% = CODINT%(Label28)
    Label30 = DESCRIT0$(CIXI%)
    Call LIMPLAB
    Call LIMPTEXT
    Call ARMALILOTES(CIXI%)
End Sub

Sub ARMALILOTES(CIXI%)
    '
     FLEXCONN.EXECUTE (FILTSQL$ ("SET DATEFORMAT dmy"))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " INNER JOIN BOLRECEP"
'    SQLX$ = SQLX$ + " ON LOTINGRE.IDENLOTE = BOLRECEP.IDENLOTE"
    SQLX$ = SQLX$ + " WHERE LOTINGRE.Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND (CanAlta - CanConsu) > 0"    ' no trae los consumidos
    SQLX$ = SQLX$ + " ORDER BY FechAlta ASC, LOTINGRE.IdenLote ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    '
    REBLA$ = Space$(256)
    JJ& = 0: List1.Clear
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Figuran Lotes Disponibles\para el Material Elegido.")
         GoTo 99
       End If
       '
       On Error GoTo 0
       Do While Not .EOF
         IDLOT$ = TRIM$(!IdenLote)
         '
         TTXX$ = REBLA$
         Call REPLA(TTXX$, IDLOT$, 1, 20) 'lote
         FEX = CVINT(!FechVenc)
         Call REPLA(TTXX$, FECHATEX$(FEX), 23, 8)                          ' Fecha vencimiento
         Call REPLA(TTXX$, SAFETEXT(!UbicFis), 33, 16)                      ' Ubicacion
         Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta - !CanConsu), "F12.1"), 48, 12)  ' Kilos
         Call REPLA(TTXX$, XVALO(!DepoUbic), 70, 2)                         ' Deposito
         Call REPLA(TTXX$, SAFETEXT(!CodEnvase), 72, 25)                    ' Envase
         FEXALTA = CVINT(!FechAlta)
         Call REPLA(TTXX$, FECHATEX$(FEXALTA), 97, 8)                       ' fecha Ingreso
         FEXFAB = CVINT(!FechFab)
         Call REPLA(TTXX$, FECHATEX$(FEXFAB), 105, 8)                       ' Fecha fabricación
            COEFI = 1: If !CanAlta > 0 Then COEFI = !Cantidad / !CanAlta
         Call REPLA(TTXX$, FORMATNUM(XVALO(COEFI * (!CanAlta - !CanConsu)), "F12.1"), 113, 12) ' Cantidad
         '
         JJ& = JJ& + 1
         List1.AddItem TTXX$
       .MoveNext
       Loop
    End With
    '
99  LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
End Sub

Private Sub Label28_DblClick()
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 16 Then
       FRABOBI07.Label28 = TRIM$(Left$(TTXX$, 16))
       For KKI% = 1 To List1.ListCount
          LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
          If TRIM$(Left$(List1.List(KKI% - 1), 12)) = LOTSE$ Then
             List1.Selected(KKI% - 1) = True
             Call List1_Click
             Exit Sub
          End If
       Next KKI%
       Call MENSERR(24, "Lectura no Valida")
       GoTo 10
    End If
    '
End Sub

Private Sub Label42_Change()
    CIXJ% = CODINT%(Label42)
    Label45 = DESCRIT0$(CIXJ%)
End Sub

Private Sub List1_Click()
    
    Call LIMPTEXT
    '
    TTXX$ = List1.TEXT
    '
    NCR0$ = Mid(TTXX$, 1, 20)
    COD$ = Label28
    CI% = CODINT%(COD$)
    Label42 = COD$                      ' EN PRINCIPIO ES EL MISMO CODIGO
    Label18.Caption = NCR0$
    Label24.Caption = TRIM$(Mid$(TTXX$, 72, 25)) ' Envase
    Label37.Caption = TRIM$(Mid$(TTXX$, 97, 8))   ' F. Ingreso
    Label36.Caption = TRIM$(Mid$(TTXX$, 105, 8))  ' F. Fabricación
    Label38.Caption = TRIM$(Mid$(TTXX$, 23, 8))   ' F. Vencimiento
    Label40.Caption = TRIM$(Mid$(TTXX$, 113, 12))  ' Cantidad
    Label33 = Label40
    Label41.Caption = TRIM$(Mid$(TTXX$, 49, 12))  ' Kilos
    Label39.Caption = TRIM$(Mid$(TTXX$, 33, 16))  ' Ubicación
    Label34 = Label39.Caption
    Label2 = TRIM$(VALOBADA("LOTINGRE", "Libera", "IdenLote = '" & NCR0$ & "'"))
       COEFI = 1: If XVALO(Label41) > 0 Then COEFI = XVALO(Label40) / XVALO(Label41)
    Label35 = FORMATNUM$(COEFI, "F15.5")
    '
    NCR$ = TRIM$(NCR0$)
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1) ' VERIFICA LA EXISTENCIA DE LA BARRA IDENTIFICADORA DE SUBLOTE
    End If
    '
    Call APERBASDAT("STOCKS")
    For KKII% = 65 To 90
       NCR1$ = NCR$ + "/" + Chr$(KKII%) ' ARMA EL NUEVO SUBLOTE
      SQLX$ = "SELECT * FROM LOTINGRE"  ' VERIFICA LA EXISTENCIA DEL SUBLOTE
      SQLX$ = SQLX$ + " WHERE IdenLote = '" & NCR1$ & "'"
      Set BRTEMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX))
      If BRTEMP.EOF And BRTEMP.BOF Then GoTo 20 ' SI NO EXISTE LO ASGNA A NCR1$
    Next KKII%
    Call MENSERR(24, "Imposible Fraccionar Este Lote.")
   Exit Sub
   '
20 Label6 = NCR1$                   'CARGA EL NUEVO SUBLOTE
   Label27 = NCR0$                  'CARGA EL LOTE
   Label32 = Label41.Caption        'CARGA LOS KILOS DEL LOTE
   Text2.SetFocus
   '
End Sub

Sub LIMPLAB()
    Label2.Caption = ""
    Label24.Caption = ""
    Label36.Caption = ""
    Label37.Caption = ""
    Label38.Caption = ""
    Label39.Caption = ""
    Label40.Caption = ""
    Label41.Caption = ""
    Label6.Caption = ""
    Label27.Caption = ""
    Label32.Caption = ""
    Label18.Caption = ""
    Label33.Caption = ""
    Label42.Caption = ""
    Label45.Caption = ""
    Label34.Caption = ""
End Sub

Sub LIMPTEXT()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text11 = ""
    Text16 = ""
    Text12 = ""
    Text17 = ""
    Call BLANARCH("!CARFRABO")
End Sub

Private Sub Text1_Change()
   '
   Label32.Caption = FORMATNUM$(XVALO(Label41.Caption) - XVALO(Text1.TEXT), "F12.1")  'RESTA AL LOTE ORIGINALLOS KILOS DEL NUEVO LOTE
   COEFI = XVALO(Label35)
   Text4 = FORMATNUM$(XVALO(Text1) * COEFI, "F12.1")
   Label33 = FORMATNUM$(XVALO(Label32) * COEFI, "F12.1")
   If XVALO(Label32.Caption) < 0 Then
       Call MENSERR(24, "Cantidad de Kilos Supera la del Lote Original !!!")
   End If
   '
End Sub

Private Sub Text1_LostFocus()
    '
    Text1.TEXT = FORMATNUM$(Text1.TEXT, "F12.1") ' FORMATEA EL TEXT1 "KILOS"
    '
End Sub
'
Private Sub Text4_LostFocus()
    '
    Text4.TEXT = FORMATNUM$(Text4.TEXT, "F12.1") ' FORMATEA EL TEXT2 "CANTIDAD"
    '
End Sub

Private Sub Text4_DblClick()
    Call Command7_Click
End Sub

Private Sub Text3_DblClick()
    Call Command4_Click
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM$(Text7.TEXT, "F12.1")
End Sub



