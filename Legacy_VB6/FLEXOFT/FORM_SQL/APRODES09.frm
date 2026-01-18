VERSION 5.00
Begin VB.Form APRODES09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobacion por Destino"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   975
      Left            =   12300
      ScaleHeight     =   915
      ScaleWidth      =   1080
      TabIndex        =   43
      Top             =   0
      Width           =   1140
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
         Left            =   3255
         TabIndex        =   47
         Top             =   0
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   7035
         TabIndex        =   45
         Top             =   525
         Width           =   1320
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
         Picture         =   "APRODES09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label16 
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
         TabIndex        =   53
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label12 
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
         TabIndex        =   52
         Top             =   1095
         Width           =   1545
      End
      Begin VB.Label Label5 
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
         TabIndex        =   51
         Top             =   570
         Width           =   1545
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
         Left            =   1680
         TabIndex        =   50
         Top             =   525
         Width           =   1650
      End
      Begin VB.Label Label9 
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
         TabIndex        =   49
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label11 
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
         TabIndex        =   48
         Top             =   525
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   1575
      Left            =   12300
      ScaleHeight     =   1515
      ScaleWidth      =   1080
      TabIndex        =   27
      Top             =   5260
      Width           =   1140
      Begin VB.CommandButton Command3 
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
         Left            =   130
         Picture         =   "APRODES09.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Invierte la Seleccion"
         Top             =   60
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   130
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   31
         Top             =   800
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   12000
         End
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
         TabIndex        =   30
         Top             =   525
         Width           =   1320
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
         TabIndex        =   29
         Top             =   1050
         Width           =   6675
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
         TabIndex        =   28
         Top             =   0
         Width           =   175
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -270
         Picture         =   "APRODES09.frx":0294
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   960
         Width           =   1515
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
         TabIndex        =   40
         Top             =   525
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
         TabIndex        =   39
         Top             =   570
         Width           =   1545
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
         TabIndex        =   38
         Top             =   570
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
         Left            =   1680
         TabIndex        =   37
         Top             =   525
         Width           =   1650
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
         TabIndex        =   36
         Top             =   570
         Width           =   1545
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
         TabIndex        =   35
         Top             =   1095
         Width           =   1545
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
         TabIndex        =   34
         Top             =   0
         Width           =   1650
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
      Left            =   11895
      TabIndex        =   22
      Top             =   525
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1280
      Left            =   105
      TabIndex        =   12
      Top             =   5250
      Width           =   11115
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   6360
         TabIndex        =   55
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2040
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
         Height          =   500
         Left            =   1185
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   675
         Width           =   7350
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   8865
         TabIndex        =   16
         Top             =   225
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   9585
         TabIndex        =   15
         Top             =   495
         Width           =   1380
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Imprime Comprobante"
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
         Left            =   8895
         TabIndex        =   14
         Top             =   840
         Width           =   2040
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
         Left            =   6240
         TabIndex        =   13
         Top             =   240
         Width           =   2280
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
         TabIndex        =   21
         Top             =   735
         Width           =   1020
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
         TabIndex        =   20
         Top             =   285
         Width           =   2145
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
         TabIndex        =   19
         Top             =   240
         Width           =   2775
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
         Left            =   5115
         TabIndex        =   18
         Top             =   255
         Width           =   1245
      End
   End
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
      Height          =   1305
      Left            =   11325
      Picture         =   "APRODES09.frx":21B6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Graba - Procesa la Aprobacion"
      Top             =   5260
      Width           =   825
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
      Left            =   2715
      TabIndex        =   7
      Top             =   195
      Width           =   175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Pendientes de Aprobacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4260
      Left            =   120
      TabIndex        =   0
      Top             =   945
      Width           =   13005
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   10750
         TabIndex        =   54
         Top             =   300
         Width           =   30
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   12580
         TabIndex        =   42
         Top             =   300
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   12170
         TabIndex        =   41
         Top             =   300
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   3800
         TabIndex        =   4
         Top             =   300
         Width           =   30
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   9700
         TabIndex        =   3
         Top             =   300
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   8200
         TabIndex        =   2
         Top             =   300
         Width           =   30
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3840
         Left            =   2050
         TabIndex        =   1
         Top             =   300
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
         Height          =   3630
         ItemData        =   "APRODES09.frx":24C0
         Left            =   160
         List            =   "APRODES09.frx":24C2
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   530
         Width           =   12700
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"APRODES09.frx":24C4
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
         Left            =   160
         TabIndex        =   6
         Top             =   285
         Width           =   12800
      End
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
      Left            =   10950
      TabIndex        =   26
      Top             =   525
      Width           =   1020
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
      Left            =   9330
      TabIndex        =   25
      Top             =   555
      Width           =   1545
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Recep:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9225
      TabIndex        =   24
      Top             =   210
      Width           =   1620
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
      Left            =   10950
      TabIndex        =   23
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label Label4 
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
      TabIndex        =   10
      Top             =   525
      Width           =   8400
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
      Left            =   1155
      TabIndex        =   9
      Top             =   195
      Width           =   1650
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "S/F Nro:"
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
      Left            =   195
      TabIndex        =   8
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "APRODES09"
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
    '
    Command2.Enabled = False
    '
    FFXX% = FECHANUM(Label13.Caption)
    If FFXX% = 0 Then
       Call MENSERR%(24, "Fecha Invalida")
       GoTo 99
    End If
    '
    'SE ANULA VALIDACION DE LA REFERENCIA
'    If Text6 = "" Then
'       Call MENSERR%(24, "Debe Ingresar Referencia de Aprobación")
'       Text6.SetFocus
'       GoTo 99
'    End If
    '
    NUCERT& = XVALO(Label54)
    CONDI2$ = "NuCert_Apro = " & NUCERT&
    If CantRegistros("LOTINGRE", CONDI2$) > 0 Then
       NUCERT& = XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0")) + 1
       Label54 = NUCERT&
       Call COMUNI("Numero de Certificacion de Aprobacion fue Cambiado por " & NUCERT&)
       GoTo 99
    End If
    '
    For u& = 1 To List1.ListCount
      If List1.Selected(u& - 1) = True Then
        XX$ = List1.List(u& - 1)
        CODXX$ = Mid$(XX$, 17, 16)
        CIXXI% = CODINT(CODXX$)
        If CIXXI% = 0 Then
           Call MENSERR%(24, "Codigo de Articulo Invalido")
           GoTo 89
        End If
        '
        IDLOT$ = TRIM$(Left$(XX$, 16))
        Call ApruebaImprime(CIXXI%, IDLOT$)
        If Option2.Value = True Then
           Call ENVIMAILRECH(CIXXI%, IDLOT$)
           CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE='" & IDLOT$ & "'"
           Call BORRAREGISTROS("BALAREQMO", CONDI$, REG&)
        End If
      End If
89  Next u&
    '
    If Option1.Value = True Then
          'ENVIO DE MAIL PARA RECEPCIONES APROBADAS
          Call ENVIMAILAPRO
          Call COMUNI("Aprobación Procesada")
        ElseIf Option2.Value = True Then
          Call COMUNI("Rechazo Procesado")
    End If
    ' Se ha dejado el check1 no visible porque se hace referencia a el
    ' en la rutina ApruebaImprime. Si se quita arroja error 424
    If Check2.Value = 1 Then
        PRINTFORM
    End If
    '
    'limpio los label y los text
    Label36 = "": Label4 = ""
    Text5 = "": Text6 = ""
    List1.Clear
    
    'actualizo
    Call Form_Load
    
99  Command2.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   For u& = 1 To List1.ListCount
      List1.Selected(u& - 1) = Not (List1.Selected(u& - 1))
   Next u&
   Command3.Enabled = True
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
  CONDI$ = "(LIBERA ='CUAR' OR LIBERA = '' OR LIBERA IS NULL)"
  CONDI$ = CONDI$ & " AND COD_DESTI <> '' "
  CONDI$ = CONDI$ & " ORDER BY COD_DESTI ASC"
  '
  Call CARGALISEL("!SOFACUAR", "LOTINGRE+ORDEFABR/COD_DESTI = IDSUBOR", "COD_DESTI/A14;REFEREN/A64", CONDI$, "DISTINCT")
  Call SELECHO("!SOFACUAR/Material en Cuarentena por Destino")
  VASEL$ = VALACT1("!SOFACUAR")
  If VASEL$ <> "" Then
     Label36 = VASEL$
  End If
  '
End Sub

Private Sub Form_Load()
   Label45 = USERNAME
   NUCERT& = XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0")) + 1
   Label54 = NUCERT&
   HOII% = HOY(HOS$)
   Label13 = HOS$
End Sub

Private Sub Label36_Change()

   If TRIM$(Label36) = "" Then Exit Sub
   
   Label4 = VALOBADA("ORDEFABR", "REFEREN", "IDSUBOR = '" & Label36 & "'")
   SQLX$ = "select * from lotingre where "
   SQLX$ = SQLX$ + "cod_desti = '" & Label36 & "' "
   SQLX$ = SQLX$ + "AND (LIBERA ='CUAR' OR LIBERA = '' OR LIBERA IS NULL) "
   Dim pendi As ADODB.Recordset
   Set pendi = FLEXCONN.Execute(FILTSQL(SQLX$))
   List1.Clear
   With pendi
     Do While Not .EOF
       CI% = XVALO(!Cod_Inte)
       CODX$ = CODEXT$(CI%)
       DESC$ = DESCRIT0$(CI%)
       LOTE$ = SAFETEXT(!idenLote)
       LOTEPRO$ = SAFETEXT$(!LOTECOL_PROV)
       UMED$ = UNIMED$(CI%)
       FECHVEN$ = FECHATEX(CVINT(!FECHVENC))
       
       CAN# = XVALO(!CANALTA)
       CANTX$ = FORMATSTAN(CAN#, "F13", CI%)
       'CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE ='" & LOTE$ & "'"
       'LOTEPRO$ = VALOBADA("LOTINGRE", "LOTECOL_PROV", CONDI$)
       '
       LINEA$ = ""
       Call REPLA(LINEA$, LOTE$, 1, 16)
       Call REPLA(LINEA$, CODX$, 17, 16)
       Call REPLA(LINEA$, DESC$, 34, 40)
       Call REPLA(LINEA$, LOTEPRO$, 76, 16)
       'Call REPLA(LINEA$, CANTX$, 81, 13)
       Call REPLA(LINEA$, FECHVEN$, 91, 8)
       Call REPLA(LINEA$, CANTX$, 100, 13)
       Call REPLA(LINEA$, UMED$, 114, 4)
       List1.AddItem LINEA$
       '
       .MoveNext
     Loop
     
   .Close
   End With
   
   
End Sub

Sub ApruebaImprime(CIXI%, IDLOT$)
   '
   'Este procedimiento solo funciona para este formulario ya que toma variables del formulario
   
   Destino$ = TRIM$(Label35)          'Destino
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
   CIXXI% = CIXI%
   Call MOVISTOK(FFXX%, CIXXI%, IDLOT$, "CC", COMPROB1$, COMPROB1$, COMPROB1$, COMPROB1$, Operacion$ & ". Lote " & IDLOT$, 1, 0, 1, 0, 1, 0)
   '
   CONDI$ = "COD_INTE = " & CIXXI% & " AND IDENLOTE ='" & IDLOT$ & "'"
       FECHAX = FETEXCOR1$(FECHANUM(Label13))
   Call ACTUREGISTRO("LOTINGRE", "LIBERA", CONDI$, Operacion$, REG&)
   Call ACTUREGISTRO("LOTINGRE", "Ref_Apro", CONDI$, TRIM$(Text6), REG&)
      NUCERT& = 1 + XVALO(VALOBADA("LOTINGRE", "MAX(NUCERT_APRO)", "Nucert_Apro <> 0"))
   Call ACTUREGISTRO("LOTINGRE", "NuCert_Apro", CONDI$, NUCERT&, REG&)
   Call ACTUREGISTRO("LOTINGRE", "Usu_Apro", CONDI$, USERNAME, REG&)
   Call ACTUREGISTRO("LOTINGRE", "Fech_Apro", CONDI$, FECHAX, REG&)
   Call ACTUREGISTRO("LOTINGRE", "Nro_Analisis", CONDI$, TRIM$(Text5), REG&)
   '
   Call ACTUREGISTRO("BALAREQMO", "NuCert_Apro", CONDI$, NUCERT&, REG&)
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

Sub ENVIMAILAPRO()
   '
   mailProgra$ = CONTROL("BOLRECEP", "PROGMAIL")
   '
   '
   ASUNTO$ = "Aprobacion"
   '
   TEXTOMAI$ = "Producto(s) Aprobado(s): " & CRLF$
   For u& = 1 To List1.ListCount
     If List1.Selected(u& - 1) = True Then
        XX$ = List1.List(u& - 1)
        CODXX$ = Mid$(XX$, 17, 16)
        CIXXI% = CODINT(CODXX$)
        IDLOT$ = TRIM$(Left$(XX$, 16))
        '
        TEXTOMAI$ = TEXTOMAI$ + CODEXT$(CI%) & " - " & DESCRIT$(CI%) & CRLF$
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
        NUPED$ = VALOBADA("LOTINGRE", "DOCUORIG", CONDI$, "NOMESS")
        TEXTOMAI$ = TEXTOMAI$ & "    Nº de Pedido: " & NUPED$ + CRLF$
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
        SOLFA$ = VALOBADA("BALAREQMO", "IDENREQUE", CONDI$, "NOMESS")
        TEXTOMAI$ = TEXTOMAI$ & "    Destino: " & SOLFA$ & CRLF$
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & IDLOT$ & "'"
        RERECHA$ = VALOBADA("LOTINGRE", "Ref_Apro", CONDI$, "NOMESS")
     End If
   Next u&
   TEXTOMAI$ = TEXTOMAI$ & "Aprobados con Referencia: " & RERECHA$ & CRLF$
   '
   MAI_DESTI$ = TRIM$(mailProgra$)
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   '
   Call GENERAMAIL
   '
End Sub
