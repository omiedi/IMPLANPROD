VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PADIAMOD07 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parte Diario  M.O.D."
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10104
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5844
   ScaleWidth      =   10104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   384
      ScaleWidth      =   1128
      TabIndex        =   61
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CMDG 
      Left            =   210
      Top             =   5250
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6405
      TabIndex        =   59
      Text            =   " "
      Top             =   5355
      Width           =   1275
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6405
      TabIndex        =   58
      Text            =   " "
      Top             =   4410
      Width           =   1275
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8085
      TabIndex        =   16
      Top             =   3780
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3885
      TabIndex        =   14
      Text            =   " "
      Top             =   3360
      Width           =   960
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "At.Múltiple"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7900
      TabIndex        =   18
      Top             =   4095
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Células"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1365
      TabIndex        =   4
      ToolTipText     =   "Informe por Célula"
      Top             =   1200
      Width           =   910
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Re-Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7900
      TabIndex        =   19
      Top             =   4410
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   1680
      Width           =   175
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   5
      Top             =   1680
      Width           =   2010
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   1365
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "PADIAMOD07.frx":0000
      Top             =   4200
      Width           =   4740
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8085
      TabIndex        =   15
      Text            =   " "
      Top             =   3360
      Width           =   960
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   13
      Text            =   " "
      Top             =   3780
      Width           =   960
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   12
      Text            =   " "
      Top             =   3360
      Width           =   960
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7770
      TabIndex        =   10
      Top             =   2520
      Width           =   1050
   End
   Begin VB.CommandButton Command7 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8820
      TabIndex        =   11
      Top             =   2520
      Width           =   175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1995
      TabIndex        =   9
      Top             =   2520
      Width           =   175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   8
      Top             =   2520
      Width           =   630
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8820
      TabIndex        =   7
      Top             =   1680
      Width           =   175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   2
      Top             =   840
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2100
      TabIndex        =   3
      Top             =   840
      Width           =   175
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   7875
      Picture         =   "PADIAMOD07.frx":0002
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "O.K."
      Top             =   4935
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8850
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7785
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   1065
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   8460
      Left            =   9225
      ScaleHeight     =   8412
      ScaleWidth      =   1020
      TabIndex        =   27
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Importación de operaciones"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":0616
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Listado de Informe de Produccion por Fecha"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":0920
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   3675
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":0D62
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Corrección Partes de Fabricación"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "A-B-M Padrón de Operarios"
         Top             =   3030
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   84
         ScaleWidth      =   600
         TabIndex        =   28
         Top             =   5145
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":11B6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "PADIAMOD07.frx":1300
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   312
         Left            =   -312
         Picture         =   "PADIAMOD07.frx":160A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5352
         Width           =   1212
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "PADIAMOD07.frx":352C
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PADIAMOD07.frx":3760
      TabIndex        =   62
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6090
      TabIndex        =   57
      Top             =   5150
      Width           =   1125
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro.Inf.NC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6060
      TabIndex        =   56
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No Conf:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   55
      Top             =   3855
      Width           =   1335
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3885
      TabIndex        =   54
      Top             =   3780
      Width           =   960
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hs. Operación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   53
      Top             =   3855
      Width           =   1335
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hs.Descanso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2625
      TabIndex        =   52
      Top             =   3405
      Width           =   1230
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Ini:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4830
      TabIndex        =   51
      Top             =   3405
      Width           =   1125
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cadencia STD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2415
      TabIndex        =   50
      Top             =   2580
      Width           =   1755
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   49
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5985
      TabIndex        =   48
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5985
      TabIndex        =   47
      Top             =   3780
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4620
      TabIndex        =   46
      Top             =   3855
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/F Nro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5775
      TabIndex        =   45
      Top             =   1755
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Parte Diario M.O.D."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   1365
      TabIndex        =   44
      Top             =   210
      Width           =   4215
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observacion.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -45
      TabIndex        =   43
      Top             =   4220
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.OK:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6720
      TabIndex        =   42
      Top             =   3405
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hora Fin:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -105
      TabIndex        =   41
      Top             =   3850
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comienzo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -105
      TabIndex        =   40
      Top             =   3400
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5340
      TabIndex        =   39
      Top             =   2895
      Width           =   3690
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5355
      TabIndex        =   38
      Top             =   2580
      Width           =   2325
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -1050
      TabIndex        =   37
      Top             =   2580
      Width           =   2325
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   36
      Top             =   2895
      Width           =   3795
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   35
      Top             =   1755
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7140
      TabIndex        =   34
      Top             =   1680
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1365
      TabIndex        =   33
      Top             =   2100
      Width           =   7650
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Legajo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   315
      TabIndex        =   32
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   2415
      TabIndex        =   31
      ToolTipText     =   "Haga Doble-Click para Informe por Célula"
      Top             =   840
      Width           =   6600
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6405
      TabIndex        =   30
      Top             =   375
      Width           =   1335
   End
End
Attribute VB_Name = "PADIAMOD07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ALTA_EMP(NLGJO%, NOAP$)
    
  VDEF$ = Space(32)
  Call REPLA(VDEF$, MKI$(NLGJO%), 1, 2)
  Call REPLA(VDEF$, NOAP$, 3, 30)
  
9 obj$ = OBJPLA$("ALTEMP", VDEF$)
  If obj$ = "" Then Exit Sub
  VDEF$ = obj$
  '
  NULEG% = CVI(Mid$(obj$, 1, 2))
  DESCRI$ = TRIM$(Mid$(obj$, 3, 30))
  NUGRU% = CVI(Mid$(obj$, 33, 2))
  '//////////////VALIDACIONES////////////
  If Val(NULEG%) <= 0 Then
    Call COMUNI("Número de Legajo no válido.")
    GoTo 9
  End If
  '
  If TRIM(DESCRI$) = "" Then
    Call COMUNI("Apellido y Nombre no válido.")
    GoTo 9
  End If
  '
  For J& = 1 To ULTREG("PADREMP")
    XX$ = REGLEIDO$("PADREMP", J&)
    LEG% = CVI(Mid$(XX$, 1, 2))
    NOAP$ = TRIM(Mid$(XX$, 3, 30))
    If NULEG% = LEG% Then
      Call COMUNI("Número de Legajo no válido o Pre-Existente.")
      GoTo 9
    End If
    If NOAP$ = DESCRI$ Then
      Call COMUNI("Empleado Pre-Existente.")
      GoTo 9
    End If
  Next J&
  '
  If NUGRU% <= 0 Then
    Call COMUNI("Grupo de Empleado no Válido.")
    GoTo 9
  End If
  'LIMPIO LOS TEXT
  Text1 = ""
  'GRABO EL REGISTRO
  X$ = REGBLAN("PADREMP")
  ULTRE& = ULTREG&("PADREMP")
  Call REPLA(X$, MKI$(NULEG%), 1, 2)
  Call REPLA(X$, DESCRI$, 3, 30)
  Call REPLA(X$, Chr$(NUGRU%), 49, 1)
  Call REPLA(X$, Chr$(1), 48, 1) 'ESTADO ACTIVO
  Call GRAREG("PADREMP", X$, ULTRE& + 1)
  Call SETULTREG("PADREMP", ULTRE& + 1)
  '\\\OT-05-0577-WAR-25/07/06///
  YZ$ = REGBLAN("DATEMDGI")
  ULTRE1& = ULTREG&("DATEMDGI")
  Call REPLA(YZ$, MKI$(NULEG%), 1, 2)
  Call GRAREG("DATEMDGI", YZ$, ULTRE1& + 1)
  Call SETULTREG("DATEMDGI", ULTRE1& + 1)
  '\\\OT-05-0577-WAR-FIN///
  'GRABO EN ARCHIVOS AUXILIARES
  PADA1$ = Space(64)
  Call REPLA(PADA1$, X$, 1, 64)
  Call GRAREG("PADRON", PADA1$, ULTRE& + 1)
  PADA2$ = Mid$(X$, 65, 96)
  Call GRAREG("PADAUX", PADA2$, ULTRE& + 1)
  PADA3$ = Mid$(X$, 161, 48)
  Call GRAREG("PADAUX2", PADA3$, ULTRE& + 1)
  '
  RENUE$ = String$(128, 0)
  Call GRAREG("ACTIVA", RENUE$, ULTRE& + 1) 'RENGLON SEGUN EMPLEADO
  '

End Sub

Function ALTA_OPERACION%(CI%, DESCRIOP$)
    ALTA_OPERACION% = 0
    If CI% < 1 Then Exit Function
    'VALIDO QUE LA CONSULTA DEVUELVA ALGUNA/S OPERACIONES SI NO VA DIRECTAMENTE PARA QUE GENERE NUEVA
    CONDI$ = "NumeOper > 0 AND CODICONJ = " & CI% & " AND STATOPER > 0"
    If CantRegistros&("SECOPER", CONDI$, "NOMESS") < 1 Then GoTo 30
10  Call CARGALISEL("OPERCION", "SECOPER", "NumeOper/I5;DESCOPER/A32", CONDI$, "NOMESS")
    Call SELECHO("OPERCION/CORRESPONDE A OPERACION EXISTENTE")
    NOPE% = XVALO(VALACT1$("OPERCION"))
    
    If NOPE% < 1 Then GoTo 30 'SI NO ELIGE DA DE ALTA DIRECTAMENTE
    
    CONDI$ = "CODICONJ=" & CI% & " And NUMEOPER = " & NOPE%
    Call ACTUREGISTRO("SECOPER", "DESCOPER", CONDI$, DESCRIOP$, REGAF&, "NOMESS")
    ALTA_OPERACION% = NOPE%
    Exit Function
    
    'DAR DE ALTA LA OPERACION
    '
30   NOPE% = XVALO(InputBox$("Ingrese el Nº de Operación", "Alta de Nueva Operacion"))
40   DENO$ = InputBox$("Ingrese la Denominación", "Alta de Nueva Operacion", DESCRIOP$)
     If NOPE% < 1 Then
         Call COMUNI("Imposible Continuar. Número de Operación No Valido") ' VER SI ESTE VALOR DE MAXIMO ES CORRECTO
         GoTo 30
     End If
     If TRIM$(DENO$) = "" Then
         Call COMUNI("Imposible Continuar. Falta Ingresar Denominación de la Operación")
         GoTo 40
     End If
     
     'VALIDO SI EXISTE YA ESA OPERACION PARA QUE NO DUPLIQUE
     CONDI1$ = "NumeOper = " & NOPE% & "  AND CODICONJ = " & CI% & " AND STATOPER > 0"
     If CantRegistros&("SECOPER", CONDI1$, "NOMESS") > 0 Then
        Call COMUNI("Imposible Registrar. Número de Operación Existente")
        GoTo 30
     End If

50  ABUS$ = RANDSTRING$(10)
    CONDI$ = "DESCOPER = '" & ABUS$ & "'"
    If CantRegistros("SECOPER", CONDI$) > 0 Then GoTo 50
    Call CreaRegistro("SECOPER", "DESCOPER", ABUS$, "NOMESS")
    Call ACTUREGISTRO("SECOPER", "NUMEOPER", CONDI$, FORMATNUM$(NOPE%, "F5.0"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("SECOPER", "CODICONJ", CONDI$, FORMATNUM$(CI%, "F5.0"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("SECOPER", "CODXCONJ", CONDI$, CODEXT$(CI%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("SECOPER", "StatOper", CONDI$, "1", REGAF&, "NOMESS")
    Call ACTUREGISTRO("SECOPER", "DESCOPER", CONDI$, DENO$, REGAF&, "NOMESS")
    ALTA_OPERACION% = NOPE%
    '
    
End Function

Private Sub Check1_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      Command2.SetFocus
   End If
   '
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      Check1.SetFocus
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
    '
    Command10.Enabled = False
    '
    'COMALTER CON DOS OPCIONES
8   OPX% = COMALTER%("Opciones de Trabajo\\ Importar Informes de Producción\ Reasignación de Informes Con O.Fab.\Reasignación de Informes Sin O.Fab.")
    If OPX% < 1 Or OPX% > 3 Then GoTo 99
    NORFA$ = ""
    'REASIGNACION DE INFORME CON ORDEN DE FABRICACION
    If OPX% = 2 Then
9      Call SELECHO("ECORDEF")
       NORFA$ = TRIM$(VALACT1$("ECORDEF/ORDENES DE FABRICACIÓN"))
       If NORFA$ = "" Then GoTo 8
       Call REASIGNA_OF(NORFA$)
       GoTo 9
    End If
    'REASIGNACION DE INFORME SIN ORDEN DE FABRICACION
    If OPX% = 3 Then Call REASIGNA_OF(NORFA$): GoTo 8
    
    'IMPORTACION DE INFORMES DE PRODUCCION
    '*************************************
    'RESUMEN
    'SELECCION DE RANGO DE FECHAS PARA FILTRAR LA INFORMACION
    'LEER EL ARCHIVO DE TEXTO Y DAR DE ALTA CODIGOS, OPERACIONES Y LEGAJOS EN LOS CASOS QUE NO EXISTAN
    'VALIDAR SI EXISTE EL INFORME EN CASO CONTRARIO DARLO DE ALTA EN INFOFAB
    '/////////////////////////////////////////////////////////////////////////////////////////
    
    'SELECCION DE RANGO DE FECHAS
    Call DESHASFECHA(DES%, Has%, PERIO$)
    VDEVU$ = PERIO$
    If VDEVU$ = "" Then
      GoTo 99
    End If
    '
    'SELECCION DE ARCHIVO TXT
    CMDG.Filter = "(*.CSV)|*.CSV|"
    CMDG.ShowOpen
    Ruta$ = CMDG.FileName
    If EXISTE%(Ruta$) < 1 Then GoTo 99
    
    '
    'LECTURA DE ARCHIVO DE TEXTO
    NARCHI% = FILEOPEN%(Ruta$, 1, 256)
    Do While Not EOF(NARCHI%)
       Line Input #NARCHI%, XX$
       CODIEX$ = VALOTXT$(XX$, 5)
       DESCRIPCIOX$ = VALOTXT$(XX$, 6)
       FECHAF = VALOTXT$(XX$, 11)
       fech% = FOR_DDMMYY%(FECHAF)
       '
       'VALIDO FILTRO POR FECHA Y EXISTENCIA DE CODIGO
       If fech% >= DES% And fech% <= Has% Then
10        CI% = CODINT(CODIEX$)
          If CI% < 1 Then
             OPX% = COMALTER%("Código '" & CODIEX$ & "' Inexistente\\ Dar de Alta\ Continuar Sin Dar de Alta")
             If OPX% <> 1 Then GoTo 20 'LOOP
             '
             'ALTA DE CODIGO
             TX$ = String$(128, 0)
             Call REPLA(TX$, MKI$(1 + NUMAS%), 1, 2)
             Call REPLA(TX$, DESCRIPCIOX$, 5, 64)
             Call REPLA(TX$, CODIEX$, 77, 24)
             Call REPLA(TX$, "U.", 109, 2)
             Call REPLA(TX$, Chr$(1), 101, 1)
             Call ALTACODIGO(TX$, OK%)
             If OK% > 0 Then GoTo 10
             
          End If
          '
          'LEO EL LEGAJO
          LEGA% = XVALO(VALOTXT$(XX$, 2))
          NOAP$ = TRIM$(VALOTXT$(XX$, 3))
          NOAP$ = NOAP$ + " " + TRIM$(VALOTXT$(XX$, 4))
          
          'VALIDO SI EXISTE
15        XX1$ = RECFILT$("PADREMP", 1, MKI$(LEGA%))
          REG& = CVS(XX1$)
          TXTX$ = REGLEIDO$("PADREMP", REG&)
          If REG& < 1 Then
             OPX% = COMALTER%("Legajo '" & LEGA% & "' Inexistente\\ Dar de Alta\ Continuar Sin Dar de Alta")
             If OPX% = 1 Then
                 Call ALTA_EMP(LEGA%, NOAP$)
                 GoTo 15
             End If
             '
          Else
            'ACTUALIZO EL ARCHIVO DAT
            Call REPLA(TXTX$, MKI$(LEGA%), 1, 2)
            Call REPLA(TXTX$, NOAP$, 3, 30)
            Call GRAREG("PADREMP", TXTX$, REG&)
            
            'ACTUALIZO EL RESTO DE LAS TABLAS
            PADA1$ = Mid$(TXTX$, 1, 64)
            Call GRAREG("PADRON", PADA1$, REG&)
            '
            PADA2$ = Mid$(TXTX$, 65, 96)
            Call GRAREG("PADAUX", PADA2$, REG&)
            '
            PADA3$ = Mid$(TXTX$, 161, 48)
            Call GRAREG("PADAUX2", PADA3$, REG&)

          End If
          '
          'VERIFICAR SI EXISTE LA OPERACION
          If CI% < 1 Then GoTo 20 ' SI NO TIENE CODIGO VALIDO NO PUEDO DAR DE ALTA LA OPERACION
18        DESCOPERA$ = TRIM$(VALOTXT$(XX$, 6))
          CREG& = CantRegistros&("SECOPER", "DESCOPER = '" & DESCOPERA$ & "'", "NOMESS")
          If CREG& < 1 Then
             OPX% = COMALTER%("Operación (" & DESCOPERA$ & ") Inexistente para Articulo (" & CODEXT$(CI%) & ")\\ Dar de Alta\ Continuar Sin Dar de Alta")
             If OPX% = 1 Then
                 NUMEOPE% = ALTA_OPERACION%(CI%, DESCOPERA$)
                 GoTo 18
             End If
          Else 'OBTENGO EL NUMERO DE OPERACION
             NUMEOPE% = XVALO(VALOBADA("SECOPER", "NUMEOPER", "DESCOPER = '" & DESCOPERA$ & "'", "NOMESS"))
          End If
          '
          'BUSCAR COMBINACION DE CODIGO , LEGAJO,OPERACION
          HINI = Int(HORANUM(VALOTXT$(XX$, 9)))
          HFIN = Int(HORANUM(VALOTXT$(XX$, 10)))
          MAQUI$ = VALOTXT$(XX$, 1)

          CONDI$ = " COD_INTE = " & CI%
          CONDI$ = CONDI$ + " AND MAQUINA = '" & MAQUI$ & "'"
          CONDI$ = CONDI$ + " AND NUMEOPER = " & NUMEOPE%
          CONDI$ = CONDI$ + " AND FEHORLIB = '" & FETEXCOR1$(fech%) & "'"
          CONDI$ = CONDI$ + " AND HORA_DES = " & HINI
          CONDI$ = CONDI$ + " AND HORA_HAS = " & HFIN
          CNTREG& = CantRegistros&("INFOFAB", CONDI$, "NOMESS")
          If CNTREG& > 0 Then GoTo 20
          NROOF$ = ""
          
          'OBTENGO CANTIDAD DEL INFORME
          PXG = XVALO(VALOTXT$(XX$, 7)): If PXG = 0 Then PXG = 1 'PIEZAS POR GOLPE
          CANTGOLPES = XVALO(VALOTXT$(XX$, 12)) 'CANATIDAD DE GOLPES
          CantInfo# = PXG * CANTGOLPES 'CALCULO DE PIEZAS POR GOLPE * CANTIDAD DE GOLPES

          'BUSCAR OF LIBRE MEDIANTE CRITERIO FIFO
          Dim RST As ADODB.Recordset
          SQLX$ = "SELECT * FROM OPERFAB WITH(NOLOCK) WHERE COD_INTE=" & CI%
          SQLX$ = SQLX$ & " AND NUMEOPER=" & NUMEOPE%
          SQLX$ = SQLX$ & " AND CANTPROG > CANTREAL ORDER BY IDSUBOR"
          Set RST = READSET(SQLX$)
          
          With RST
            Do While Not .EOF
                CANP# = XVALO(!CANTPROG) - XVALO(!CANTREAL)
                If CANP# >= CantInfo# Then
                    NROOF$ = SAFETEXT$(!IdSubOr)
                    Exit Do
                End If
                .MoveNext
            Loop
          End With
          RST.Close
          Set RST = Nothing
          '
          '
          CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NUMEOPE%
          CAINFANT% = 1 + XVALO(VALOBADA("INFOFAB", "MAX(NuOrInf)", CONDI$, "NOMESS"))
          CANTPROGRA = XVALO(VALOBADA("OPERFAB", "CANTPROG", CONDI$, "NOMESS"))
          SALDO# = XVALO(CAPENOP$(NROOF$, NUMEOPE%, CANPROGRA))
          '
          'GUARDO INFORME EN TABLE INFOFAB
160       RNDX$ = RANDSTRING$(-10)
          If CantRegistros("InfoFab", "IDINFOFAB = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 160

          SQLX$ = " SELECT * FROM InfoFab "
          Dim INFOFFFAB As ADODB.Recordset
          Set INFOFFFAB = New ADODB.Recordset
          INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
          'With InfoFab
          With INFOFFFAB
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !IDINFOFAB = RNDX$
               !IdSubOr = NROOF$
               !DescriOF = Left$(ECOARCH$("ECORDEF", NROOF$), 64)
               !NuOrInf = CAINFANT%
               !Cod_Inte = CI%
               !Cod_Exte = CODEXT$(CI%)
               !Descrip = Left$(DESCRIT$(CI%), 64)
               !NUMEOPER = NUMEOPE%
               !AltOper = 0
               !DESCOPER = Left$(DESCOPERA$, 64) ' VALOBADA DEL SECOPÈR
               !RETRABA = 0
               !AtMulti = 0
               !CanOrPro = CANTPROGRA
               !CANTREAL = CantInfo#
               !CaNoConf = 0
               !MAQUINA = MAQUI$
               !DescMaq = Left$(ECOARCH$("PADMAQ", MAQUI$), 64)
               !LegaOps = MKI$(LEGA%)
                  THAS1% = HFIN: If THAS1% > 2359 Then THAS1% = 0
                  FEHOLI = CDate(FECHATEX$(fech%) + " " + HORATEXNOR$(Int(fech%)))
               !FeHorLib = FEHOLI
               !Hora_Des = HINI
               !Hora_Has = HFIN
               !Hora_Fri = 0
                  MALIB$ = "S"
                If CantInfo# < SALDO# - 0.05 Then
                  MALIB$ = " "
                End If
               !MaqLiber = MALIB$
                  TIEINF = TIEINF + 60 * (DIFHORA(HINI, HFIN))
               !TiemEjec = TIEINF
                  MODO% = CAINFANT%
               !TiemStan = TIESTAN(CI%, NUMEOPE%, CantInfo#, MODO%, ALTEX%)   ' 1 o mas es para que no considere preparacion
                    'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
                    ' A LOS CAMPOS DEL REGISTRO. "!"
                    'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
                  RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
                  If RENDI > 999.9 Then RENDI = 999.9
               !RENDIM = RENDI
               !Observa = ""
               !Referen = ""
               !NroInfNC = TRIM$(Left$(Text12, 32))
               !RefeInf = ""
               !UserControl = USERNAME$
               LEGANAME$ = ECOARCH$("PADRON", MKI$(LEGA%))
               !OPERARIOS = LEGANAME$
               .Update
          End With
          Set INFOFFFAB = Nothing
          CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NUMEOPE%
          CANTPROG = XVALO(VALOBADA("OPERFAB", "CANTPROG", CONDI$, "NOMESS"))
          CANTREAL = XVALO(VALOBADA("OPERFAB", "CANTREAL", CONDI$, "NOMESS"))
          CANTREAL = CANTREAL + CantInfo#
          Call ACTUREGISTRO("OPERFAB", "CANTREAL", CONDI$, CANTREAL, REGAF&, "NOMESS")
          Call ACTUREGISTRO("OPERFAB", "CANTPEND", CONDI$, (CANTPROG - CANTREAL), REGAF&, "NOMESS")
       End If
20  Loop
    '
    Close #NARCHI%

    Call COMUNI("Importación Completada")
    
99  Command10.Enabled = True

End Sub
Sub REASIGNA_OF(NORFA$)
    '
    'RESUMEN
    'SE ELIGE LA OPERACION (SEGUN ORDEN O NO SEGUN EL CASO)
    'SE MUESTRA EL OBJETO CON LA INFORMACIN SELECCIONADA
    'SE SELECCIONA CUAL ES  LA ORDEN A ASIGNAR A ESTE INFORME
    'SE GRABA LA ORDEN SELECCIONADA EN INFOFAB,
    'SE ACTUALIZA LA CANTIDAD REALIZADA EN EL OPERFAB PARA LAS DOS ORDENES
    '---------------------------------------------------------------------
    'PRESENTO LISTA DE OPERACIONES
    CONDI$ = "NUMEOPER > 0 "
    If NORFA$ <> "" Then
       CONDI$ = CONDI$ + " AND IdSubOr = '" & NORFA$ & "' "
    Else
       CONDI$ = CONDI$ + " AND IdSubOr = '' "
    End If
    
    CONDI$ = CONDI$ + " ORDER BY NumeOper  "
'    If CantRegistros&("INFOFAB", CONDI$, "NOMESS") < 1 Then
'       Call COMUNI("No Existen Informes Para la O.F. Elegida")
'       Exit Sub
'    End If
    '
    Call CARGALISEL("OPERXFAB ", "INFOFAB", "IDINFOFAB/A10;NUMEOPER/F6;DESCOPER/A48;COD_EXTE/A16", CONDI$, "NOMESS")
    TXT$ = "INFORMES SEGUN " & NORFA$
    If NORFA$ = "" Then TXT$ = "INFORMES SIN ORDEN DE FABRICACIÓN"
20  Call SELECHO("OPERXFAB/" & TXT$)
    '
    'LEO Y VALIDO
    IDINF$ = TRIM$(VALACT1$("OPERXFAB"))
    NOPE% = XVALO(Left$(VALACT2("OPERXFAB"), 6))
    CONDI$ = "IDINFOFAB = '" & IDINF$ & "'"
    DESOPE$ = VALOBADA("INFOFAB", "DESCOPER", CONDI$, "NOMESS")
    CI% = XVALO(VALOBADA("INFOFAB", "COD_INTE", CONDI$, "NOMESS"))
    CANREAL = XVALO(VALOBADA("INFOFAB", "CANTREAL", CONDI$, "NOMESS"))
    'PRESENTA LISTA DE O/f
    CONDI1$ = "COD_INTE=" & CI% & " AND STATUORF < 3"
    If NORFA$ <> "" Then CONDI1$ = CONDI1$ + " AND IDSUBOR<>'" & NORFA$ & "'"
    Call COPYSTRU("ECORDEF", "LISOFAB")
    Call CARGALISEL("LISOFAB", "ORDEFABR", "IDSUBOR/A12;REFEREN/A48", CONDI1$)
    '
    If NOPE% < 1 Or IDINF$ = "" Then Exit Sub
    '
    'PRESENTAR OBJETO
     VDEF$ = ""
     Call REPLA(VDEF$, NORFA$, 1, 12)
     Call REPLA(VDEF$, MKI$(NOPE%), 13, 2)
     Call REPLA(VDEF$, DESOPE$, 29, 26)
30   XX$ = OBJPLA$("MODIINFO", VDEF$)
     If XX$ = "" Then GoTo 20
     '
     NORFASELEC$ = TRIM$(Mid$(XX$, 17, 12))
     '
     'VALIDACIONES
     'valido que se seleccione orden
     If NORFASELEC$ = "" Then
        Call COMUNI("Falta Seleccionar Orden de Fabricación")
        VDEF$ = XX$
        GoTo 30
     End If
     CONDI$ = "NUMEOPER=" & NOPE% & " AND IDSUBOR='" & NORFASELEC$ & "'"
     CANPEND = XVALO(VALOBADA("OPERFAB", "CANTPEND", CONDI$, "NOMESS"))
     If CANREAL > CANPEND Then
        Call COMUNI("Imposible Asignar a O/F. Cantidad del Informe Supera la Pendiente de la O/F Elegida")
        VDEF$ = XX$
        GoTo 30
     End If
     '
     'ACTUALIZO EL REGISTRO
     CONDI$ = " IDINFOFAB = '" & IDINF$ & "'"
     Call ACTUREGISTRO("INFOFAB", "IDSUBOR", CONDI$, NORFASELEC$, REGAF&, "NOMESS")
     '
     'ACTUALIZO CANTIDAD PARA LA ORDEN ORIGINAL EN OPERFAB
     If NORFA$ <> "" Then 'EN EL CASO QUE SEA CON LA OPCION CON ORDEN DE FABRICACION
        CONDI$ = "IDSUBOR =  '" & NORFA$ & "' AND NUMEOPER = " & NOPE%
        CANTIRECALCU = XVALO(VALOBADA("INFOFAB", "SUM(CANTREAL)", CONDI$, "NOMESS"))
        Call ACTUREGISTRO("OPERFAB", "CANTREAL", CONDI$, FORMATNUM$(CANTIRECALCU, "F16.5"), REGAF&, "NOMESS")
        CANTPROG = XVALO(VALOBADA("OPERFAB", "CANTPROG", CONDI$, "NOMESS"))
        CANTREAL = XVALO(VALOBADA("OPERFAB", "CANTREAL", CONDI$, "NOMESS"))
        Call ACTUREGISTRO("OPERFAB", "CANTPEND", CONDI$, (CANTPROG - CANTREAL), REGAF&, "NOMESS")
     End If
     '
     'ACTUALIZO CANTIDAD PARA LA ORDEN ASIGNADA EN OPERFAB
     CONDI1$ = "IDSUBOR =  '" & NORFASELEC$ & "' AND NUMEOPER = " & NOPE%
     CANTIRECALCU1 = XVALO(VALOBADA("INFOFAB", "SUM(CANTREAL)", CONDI1$, "NOMESS"))
     Call ACTUREGISTRO("OPERFAB", "CANTREAL", CONDI1$, FORMATNUM$(CANTIRECALCU1, "F16.5"), REGAF&, "NOMESS")
     CANTPROG = XVALO(VALOBADA("OPERFAB", "CANTPROG", CONDI1$, "NOMESS"))
     CANTREAL = XVALO(VALOBADA("OPERFAB", "CANTREAL", CONDI1$, "NOMESS"))
     Call ACTUREGISTRO("OPERFAB", "CANTPEND", CONDI1$, (CANTPROG - CANTREAL), REGAF&, "NOMESS")
     '
     Call COMUNI("Reasignación Finalizada")
     
End Sub

Function VALOTXT$(TEXTO$, POS%)

   'PARA REGISTROS SEPARADOS CON PUNTO Y COMA ;
   'DEVUELVE EL VALOR DEL TEXTO EN LA POSICION INDICADA EN SEGUNDO PARAMETRO
   VALO$ = TEXTO$
   If POS% < 1 Or TRIM$(VALO$) = "" Then GoTo 99
   
   For i% = 1 To POS%
       POS1% = InStr(VALO$, ";")
       If POS1% < 1 Then
            VALO2$ = VALO$
            Exit For
       End If
       VALO2$ = Left$(VALO$, POS1% - 1)
       VALO$ = Mid$(VALO$, POS1% + 1)
   Next i%
   '
99 VALOTXT$ = VALO2$
End Function
Private Sub Command14_Click()
   Command14.Enabled = False
   STPPMOD07.Show 1
   Call COMUNI("Re-Inicie Ahora la Aplicación.")
   Call FINAL("")
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call INFOFA07.Command26_Click
   Command16.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   EMP_GES05.Show 1
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'VALIDACIONES
   HOII% = HOY(HOS$)
   If FECHANUM(Text4) < 33 Or FECHANUM(Text4) > HOII% Then
5     Call MENSERR(24, "Fecha no Válida")
      Text4.SetFocus
      GoTo 99
   End If
   '
   On Error Resume Next
   ABC12 = CDate(Text4)
   If Err Then
     On Error GoTo 0
     GoTo 5
   End If
   On Error GoTo 0
   '
   LEGAI% = XVALO(Text1)
   LEGANAME$ = ""
   If LEGAI% > 0 Then
        LEGACELU$ = MKI$(LEGAI%)
        LEGANAME$ = ECOARCH$("PADRON", MKI$(LEGAI%))
      Else
        LEGACELU$ = ""
        TCELU$ = Label5 + "/"
8       PPXX% = InStr(TCELU$, "/")
        If PPXX% > 0 Then
          LEGAL& = XVALO(Left$(TCELU$, PPXX% - 1))
          TCELU$ = Mid$(TCELU$, PPXX% + 1)
          If LEGAL& > 0 Then
            If LEGAL& <= 32767 Then
              LEGAI% = LEGAL&
              If ECOARCH$("PADRON", MKI$(LEGAI%)) <> "" Then
                ' PARA QUE NO GRABE 2 VECES EL MISMO
                For KKII% = 1 To Len(LEGACELU$) Step 2
                  If CVI(Mid$(LEGACELU$, KKII%, 2)) = LEGAI% Then GoTo 8
                Next KKII%
                LEGACELU$ = LEGACELU$ + MKI$(LEGAI%)
                LEGANAME$ = LEGANAME$ + TRIM$(ECOARCH$("PADRON", MKI$(LEGAI%))) + "; "
                GoTo 8
              End If
            End If
          End If
          Call MENSERR(24, "Legajos No Válidos en Componentes de Célula")
        End If
   End If
   LEGAI% = CVI(Left$(LEGACELU$, 2))
   If LEGAI% < 1 Or ECOARCH$("PADRON", MKI$(LEGAI%)) = "" Or LEGACELU$ = "" Then
     Call MENSERR(24, "Operario / Célula Inexistente o no Válido")
     Text1.SetFocus
     GoTo 99
   End If
   '
   CID1% = CODINT%(Text10)
   If CID1% <= 0 Then
     Call MENSERR(24, "Código no Válido")
     Text10.SetFocus
     GoTo 99
   End If
   '
   If TRIM$(Label4) <> "" Then
     Call ABREORFAB
' este es el viejo codigo cambiado para ver si mejora la velocidad 30/julio 2007
'     OrdeFabr.FindFirst "IdSubOr = '" & TRIM$(Label4) & "' "
'     If OrdeFabr.NoMatch = False Then
'       If OrdeFabr!Cod_Inte = CID1% Then
'         GoTo 19    ' VALIDA NUMERO DE o/f Y CODIGO
'       End If
'     End If
'     Call MENSERR(24, "O/F no Corresponde a Código de Artículo")
'     GoTo 99

     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & TRIM$(Label4) & "' "
     'Set ORFABU = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim ORFABU As ADODB.Recordset
     Set ORFABU = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     With ORFABU
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          On Error GoTo 0
          If !Cod_Inte = CID1% Then
             GoTo 19    ' VALIDA NUMERO DE o/f Y CODIGO
          End If
       End If
       On Error GoTo 0
       Call MENSERR(24, "O/F no Corresponde a Código de Artículo")
       GoTo 99
      End With
      Set ORFABU = Nothing
   End If
   '
19 NOPXX$ = TRIM$(Text3)
   If NOPXX$ <> "" Then
     For KKU& = 1 To ULTREG&("!OPERINF")
       XOP$ = REGLEIDO$("!OPERINF", KKU&)
       If TRIM$(Mid$(XOP$, 3, 6)) = NOPXX$ Then GoTo 21
     Next KKU&
   End If
   Call MENSERR(24, "Operación no Válida")
   Text3.SetFocus
   GoTo 99
   '
21 PPXX% = InStr(NOPXX$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NOPXX$)
   NOPEX% = 0: ALTEX% = 0
   On Error Resume Next
   NOPEX% = XVALO(Left$(NOPXX$, PPXX% - 1))
   ALTEX% = XVALO(Mid$(NOPXX$, PPXX% + 1))
   On Error GoTo 0
   '
   MAQUI1$ = Text5
   If ECOARCH$("PADMAQ", MAQUI1$) = "" Then
     Call MENSERR(24, "Máquina / Equipo no Válido")
     Text5.SetFocus
     GoTo 99
   End If
   '
   If XVALO(Text8.TEXT) < 0.005 Then
     Call MENSERR(24, "Cantidad no Válida")
     Text8.SetFocus
     GoTo 99
   End If
   '
   TDES% = Int(HORANUM(Text6))
   THAS% = Int(HORANUM(Text7))
   TFRI% = Int(HORANUM(Text2))
   If TRIM$(Text6) = "" Or TRIM$(Text7) = "" Or THAS% < TDES% Then
     Call MENSERR(24, "Horarios de Comienzo y Fin no Válidos")
     Text6.SetFocus
     GoTo 99
   End If
   '
   'GRABAR INFORME
   NROOF$ = Label4
   CIXXI% = CODINT%(Text10)
   NOPE% = NOPEX%
   DOPE$ = Label7
   CAINF# = XVALO(Text8)
   CANOC# = XVALO(Text11)
   MAQUI$ = Text5
   '
   TIEINF = 60 * (DIFHORA(TDES%, THAS%) - DIFHORA(0, TFRI%))
   If TIEINF < 0.01 Then
     Call MENSERR(24, "Horarios / Tiempos Informados no Válidos.")
     Text6.SetFocus
     GoTo 99
   End If
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   CAINFANT% = 0: CAREAL# = 0
   'Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim INFOFABTMP As ADODB.Recordset
   Set INFOFABTMP = New ADODB.Recordset
   INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         'If IsNull(!ReTraba) Or !ReTraba < 1 Then
         If XVALO(!RETRABA) < 1 Then
           CAREAL# = CAREAL# + !CANTREAL
         End If
         CAINFANT% = CAINFANT% + 1
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set INFOFABTMP = Nothing
   '
   On Error GoTo 0
   SALDO# = XVALO(Label18)
   STATUOP% = 3 ' Cumplida
   MALIB$ = "S"
   If CAINF# < SALDO# - 0.05 Then
     STATUOP% = 1   ' AUN PENDIENTE
     MALIB$ = " "
   End If
   '
   'TEXTO DE CELULA  / LE QUITO LA ULTIMA ; SI TIENE
   If InStr(Mid$(LEGANAME$, Len(LEGANAME$) - 2), ";") > 0 Then
      LEGANAME$ = Left$(LEGANAME$, Len(TRIM$(LEGANAME$)) - 1)
   End If
   
   Call ABREORFAB
   SQLX$ = " SELECT * FROM InfoFab "
   Dim INFOFFFAB As ADODB.Recordset
   Set INFOFFFAB = New ADODB.Recordset
   INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   'With InfoFab
   With INFOFFFAB
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !IdSubOr = NROOF$
      !DescriOF = Left$(ECOARCH$("ECORDEF", NROOF$), 64)
      !NuOrInf = CAINFANT%
      !Cod_Inte = CIXXI%
      !Cod_Exte = CODEXT$(CIXXI%)
      !Descrip = Left$(DESCRIT$(CIXXI%), 64)
      !NUMEOPER = NOPE%
      !AltOper = ALTEX%
      !DESCOPER = Left$(DOPE$, 64)
         MARETRA% = 0: If Check1.Value = 1 Then MARETRA% = 1
      !RETRABA = MARETRA%
         ATENMUL% = 0: If Check2.Value = 1 Then ATENMUL% = 1
      !AtMulti = ATENMUL%
      !CanOrPro = XVALO(Label19)
      !CANTREAL = CAINF#
      !CaNoConf = CANOC#
      !MAQUINA = Left$(MAQUI$, 6)
      !DescMaq = Left$(ECOARCH$("PADMAQ", MAQUI$), 64)
      !LegaOps = LEGACELU$
         FEX = FECHANUM(Text4)
         THAS1% = THAS%: If THAS1% > 2359 Then THAS1% = 0
         FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEXNOR$(Int(THAS1%)))
      !FeHorLib = FEHOLI
      !Hora_Des = TDES%
      !Hora_Has = THAS%
      !Hora_Fri = TFRI%
      !MaqLiber = MALIB$
      !TiemEjec = TIEINF
         MODO% = CAINFANT%
      !TiemStan = TIESTAN(CIXXI%, NOPE%, CAINF#, MODO%, ALTEX%)  ' 1 o mas es para que no considere preparacion
           'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
           ' A LOS CAMPOS DEL REGISTRO. "!"
           'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
         RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
         If RENDI > 999.9 Then RENDI = 999.9
      !RENDIM = RENDI
      !Observa = AJUSTI$(TRIM$(Text9), 30)
      !Referen = Text9
      !NroInfNC = TRIM$(Left$(Text12, 32))
      !RefeInf = TRIM$(Left$(Text13, 64))
      !UserControl = USERNAME$
      !OPERARIOS = LEGANAME$
      .Update
   End With
   Set INFOFFFAB = Nothing
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   CAREAL# = 0: TIEINF = 0
   'Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim INFOFABTMP1 As ADODB.Recordset
   Set INFOFABTMP1 = New ADODB.Recordset
   INFOFABTMP1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With INFOFABTMP1
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         'If IsNull(!ReTraba) Or !ReTraba < 1 Then
         If XVALO(!RETRABA) < 1 Then
           CAREAL# = CAREAL# + !CANTREAL
         End If
         TIEINF = TIEINF + !TiemEjec
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Set INFOFABTMP1 = Nothing
   '
' ESTE ES EL VIEJO CODIGO REEMPLAZADO POR TEMA VELOCIDAD ESCORIAL 30/07/07
'   With OperFab
'     .FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
'     If .NoMatch = False Then
'       .Edit
'       !CantReal = CAREAL#
'       !CantPend = !CantProg - CAREAL#
'         If !CantPend < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CantPend = 0
'         CANTPE# = !CantPend
'         modo% = CAINFANT%
'       !TiemEjec = TIEINF
'       !TiemPend = TIEMAQUI(CIXXI%, NOPE%, CANTPE#, modo%)  ' 1 es para que no considere preparacion
'           If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
'       !StatOper = STATUOP%
'         .Update
'     End If
'   End With
'   '
   SQLZ$ = " SELECT * FROM OPERFAB "
   SQLZ$ = SQLZ$ + " WHERE IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
   'Set OPERBU = Ofabric.OpenRecordset(SQLZ$, dbOpenDynaset)
   Dim OPERBU As ADODB.Recordset
   Set OPERBU = New ADODB.Recordset
   OPERBU.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   With OPERBU
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        On Error GoTo 0
       '.Edit
       !CANTREAL = CAREAL#
       !CANTPEND = !CANTPROG - CAREAL#
         If !CANTPEND < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CANTPEND = 0
         CANTPE# = !CANTPEND
         MODO% = CAINFANT%
       !TiemEjec = TIEINF
       !TiemPend = TIEMAQUI(CIXXI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
          If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
       !StatOper = STATUOP%
       .Update
     End If
     On Error GoTo 0
   End With
   Set OPERBU = Nothing
   'OPERBU.Close
   '
   If Check1.Value = 0 Then  ' PARA NO CONSIDERAR RE-TRABAJOS
     If TRIM$(NROOF$) <> "" Then
       If CAINF# >= 0.05 Then
         If NOPE% = ULTIMOP%(CIXXI%) Then
           TTXX$ = "La Operación Informada es la Última\de la Secuencia Standard de '" + TRIM$(CODEXT$(CIXXI%)) + "'."
           TTXX$ = TTXX$ + "\   \¿ Desea Informar el Cierre Parcial Correspondiente ?"
           If COMALTER%(TTXX$ + "\\Si, Cierre Parcial\No, Continuar") = 1 Then
             NROOF7$ = NROOF$
             CAINF7 = CAINF#
             FECHINFO$ = Text4
             Call ORFCIE(NROOF7$, CAINF7, FECHINFO$)
           End If
         End If
       End If
     End If
   End If
   '
   Text10 = ""
   Text2 = ""
   Text3 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text11 = ""
   Text12 = ""
   Text13 = ""
   Check1.Value = 0
   Check2.Value = 0
   Label25 = ""
   '
   Text10.SetFocus
   'AGREGUE15/09/06MH
   Check1.Value = 0
   Check2.Value = 0
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   Text10 = ""
   Label4 = ""
   MODIPAMOD07.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
   '
   NROOF$ = ""
   '
   If CODINT%(Text10) > 0 Then
     Call SELECHO("!ECORDEP/Ordenes de Trabajo en Curso")
     NROOF$ = AJUSTI$(VALACT1$("!ECORDEP"), 12)
     If TRIM$(NROOF$) = "" Then
       Label4 = ""
       Exit Sub
     End If
   End If
   '
   Label4 = NROOF$
   Text3.SetFocus
   '
End Sub

Private Sub Command4_Click()
  '
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Text1.TEXT = VDEV$
    Text10.SetFocus
  End If
  '
End Sub

Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command6_Click()
  '
  Call SELECHO("!OPERINF/Operaciones de Codigo: " + Text10)
  CODOP1$ = VALACT1$("!OPERINF")
  If CODOP1$ <> "" Then
    Y$ = FILTERGETRECORD$("!OPERINF", 1, MKI$(XVALO(CODOP1$)))
    CODOP1$ = Mid$(Y$, 3, 58)
    i% = InStr(1, CODOP1$, "-")
    CODOP2$ = TRIM$(Mid$(Y$, 1, i%))
    Text3.TEXT = CODOP2$
  End If
  '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text5.TEXT = VDEVU$
   End If
   '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELECHO("MASTER")
   If VALACT1$("MASTER") <> "" Then
     Text10 = VALACT1$("MASTER")
   End If
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  '
  Call SELECHO("ICELULAS")
  ICELU% = XVALO(VALACT1$("ICELULAS"))
  If ICELU% > 0 Then
    XX$ = FILTERGETRECORD$("ICELULAS", 1, MKI$(ICELU%))
    SLEGA$ = TRIM$(Mid$(XX$, 97, 160))
    Text1 = ""
    Label5 = SLEGA$
    LEGACELU$ = Mid$(XX$, 33, 64)
  End If
  '
  Command9.Enabled = True
End Sub

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Text4.TEXT = FECHATEX$(HOII%)
  Call INFOFA07.GENRESDAT
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  'AGREGUE15/09/06MH
  Check1.Value = 0
  Check2.Value = 0
  '
  If Control$("PADIAMOD", "INFDESCA") <> "SI" Then
      Label23.Visible = False
      Text2.Visible = False
      Label24.Top = 3540
      Label24.FontUnderline = True
      Label25.Left = 3300
  End If
  '
  If Control$("PADIAMOD", "PINOCONF") <> "SI" Then
      Label26.Visible = False
      Text11.Visible = False
      Text8.Top = Label18.Top
      Label13.Top = 3540
      Label13.Caption = "Cantidad:"
      Label13.Left = 7100
      Label13.FontUnderline = True
  End If
  '
  '
  Call APLICACIONSKINS(Me, Picture14)
  
  Screen.MousePointer = 1
  '
End Sub



Private Sub Label4_Change()
    '
    Call GENECODOPEOF(Label4)
    '
    UREX& = ULTREG&("!OPERINF")
    If UREX& > 0 Then
       For KU& = 1 To UREX&
          XYX$ = REGLEIDO$("!OPERINF", KU&)
          If XVALO(Mid$(XYX$, 49, 12)) >= 0.5 Then
            Text3 = TRIM$(Mid$(XYX$, 3, 6))
            Exit Sub
          End If
       Next KU&
    End If
    '
End Sub

Private Sub Label5_DblClick()
    '
    JJ& = 0
    REBLA$ = REGBLAN$("DCELULAS")
    TCELU$ = TRIM$(Label5)
    If XVALO(Text1) > 0 Then TCELU$ = TRIM$(Text1)
    TCELU$ = TCELU$ + "/"
8   PPXX% = InStr(TCELU$, "/")
    If PPXX% > 0 Then
       LEGAL& = XVALO(Left$(TCELU$, PPXX% - 1))
       TCELU$ = Mid$(TCELU$, PPXX% + 1)
       If LEGAL& > 0 Then
         If LEGAL& <= 32767 Then
           LEGAI% = LEGAL&
           If ECOARCH$("PADRON", MKI$(LEGAI%)) <> "" Then
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 3, 2)
             JJ& = JJ& + 1
             Call GRAREG("!DCELULAS", YY$, JJ&)
           End If
         End If
       End If
       GoTo 8
    End If
    Call SETULTREG("!DCELULAS", JJ&)
    Call CIERRARCH("!DCELULAS")
    '
    VTB% = VENTABU%("DCELULAS/NC/TITUPAN=Procesos por Célula ")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    TTXX2$ = ""
    For i& = 1 To ULTREG&("!DCELULAS")
       X$ = REGLEIDO$("!DCELULAS", i&)
       If CVI(Mid$(X$, 3, 2)) > 0 Then
         NLEGA% = CVI(Mid$(X$, 3, 2))
         For JJ& = 1 To Len(TTXX1$) Step 2
           If CVI(Mid$(TTXX1$, JJ&, 2)) = NLEGA% Then GoTo 19
         Next JJ&
         Call REPLA(X$, MKI$(ICELU%), 1, 2)
         J& = J& + 1
         Call GRAREG("!DCELULAS", X$, J&)
         If TTXX2$ <> "" Then TTXX2$ = TTXX2$ + " / "
         TTXX2$ = TTXX2$ + TRIM$(Str$(NLEGA%))
       End If
19  Next i&
    Call SETULTREG("!DCELULAS", J&)
    Call CIERRARCH("!DCELULAS")
    If ULTREG&("!DCELULAS") <= 1 Then
         Text1 = TRIM$(TTXX2$)
       Else
         Text1 = ""
         Label5 = TTXX2$
    End If
    '
End Sub

Private Sub Text1_Change()
  '
  If XVALO(Text1) < 32767 Then
      Label5.Caption = ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT)))
      LEGACELU$ = ""
    Else
      Call MENSERR(24, "Legajo no Válido.")
      Text1 = ""
  End If
  '
End Sub
Private Sub Text1_DblClick()
  '
  Call Command4_Click
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT))) <> "" Then
         Text10.SetFocus
      End If
   End If
End Sub

Private Sub Text10_Change()
   Label3 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label4 = "": CANOR% = 0
   Text3 = "": Text5 = ""
   Call SETULTREG("!ECORDEP", 0)
   If CIXXI% > 0 Then
      Call ABREORFAB
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXXI% & " "
      SQLX$ = SQLX$ + "AND StatuOrf < 3 "                ' excluye cerradas"
      SQLX$ = SQLX$ + "AND CanApro < CanProy "           ' no marca cumplidos"
      SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
      '
      'Set ORFASEL = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim ORFASEL As ADODB.Recordset
      Set ORFASEL = New ADODB.Recordset
      ORFASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

      On Error Resume Next
      ORFASEL.MoveLast
      If Err Then
         Exit Sub
      End If
      Screen.MousePointer = 11
      CANOR% = ORFASEL.RecordCount
      JJ& = 0
      With ORFASEL
        .MoveFirst
        Label4 = ORFASEL!IdSubOr
        Do While (.EOF = False)
          Y$ = REGBLAN("ECORDEP")
          Call REPLA(Y$, !IdSubOr, 1, 12)
          Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
          '
          JJ& = JJ& + 1
          Call GRAREG("!ECORDEP", Y$, JJ&)
          '
        .MoveNext
        Loop
      End With
      Set ORFASEL = Nothing
      Call SETULTREG("!ECORDEP", JJ&)
      Call CIERRARCH("!ECORDEP")
      Call FRESHECHO("!ECORDEP")
      Screen.MousePointer = 1
   End If
   '
End Sub

Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  'Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim OPFABTMP As ADODB.Recordset
  Set OPFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
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
  Set OPFABTMP = Nothing
  '
End Function

Private Sub Text10_DblClick()
   Call Command8_Click
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
     If CODINT%(Text10) < 1 Then Exit Sub
     If ULTREG&("!ECORDEP") = 1 Then
          Text3.SetFocus
        Else
          Call Command3_Click
     End If
   End If
   '
End Sub





Private Sub Text11_GotFocus()
  Text11.SelStart = 0
  Text11.SelLength = Len(Text11)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    Text9.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
  '
  Text11.TEXT = TRIM$(NUMSTRI$(XVALO(Text11.TEXT), 12, 1, 0))
  '
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Text13.SetFocus
   End If

End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
     If TRIM$(Text9) = "" Then
       Command2.SetFocus
     End If
   End If

End Sub

Private Sub Text2_Change()
   Call CALCUTIEMPO
End Sub

Private Sub Text2_GotFocus()
  Text2.SelStart = 0
  Text2.SelLength = Len(Text2)
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
  '
  If KeyAscii = 13 Then
    Text8.SetFocus
  End If
  '
End Sub

Private Sub Text2_LostFocus()
   '
   Text2 = HORATEX$(HORANUM(Text2))
   Exit Sub
   '
'  Static MODORA%
'  If MODORA% < 1 Then
'     MODORA% = 1
'     If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'        MODORA% = 2
'     End If
'  End If
'  '
'  If XVALO(Text2.TEXT) > 0 Then
'    Text2.TEXT = REPLACAR$(Text2.TEXT, ".", ":")
'    If InStr(1, Text2.TEXT, ".") = 0 Then
'      If InStr(1, Text2.TEXT, ":") = 0 Then
'        Text2.TEXT = Text2.TEXT + ":0"
'      End If
'    End If
'    '
'    If MODORA% = 2 Then
'       TTXX1$ = FORMATNUM$(REPLACAR$(Text2.TEXT, ":", "."), "F7.2")
'       Text2 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'       Exit Sub
'    End If
'    '
'    On Error Resume Next
'    DXX1 = CDate("0" + Text2.TEXT)
'    On Error GoTo 0
'    Text2.TEXT = Format$(DXX1, "hh:mm")
'    '
'  End If
'  '
End Sub

Private Sub Text3_Change()
  '
  CID1% = CODINT%(Text10)
  NOPEX% = 0: ALTEX% = 0
  NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
  NOPX$ = REPLACAR$(NOPX$, ".", "/")
  PPXX% = InStr(NOPX$, "/")
  If PPXX% < 1 Then PPXX% = 1 + Len(NOPX$)
  On Error Resume Next
  NOPEX% = XVALO(Left$(NOPX$, PPXX% - 1))
  ALTEX% = XVALO(Mid$(NOPX$, PPXX% + 1))
  On Error GoTo 0
  Label7 = DESCRIOP$(CID1%, NOPEX%, ALTEX%) ' XVALO(Text3.TEXT))
  Text5.TEXT = MAQDEFIOP$(Label4, NOPEX%)
  If ALTEX% > 0 Then
    Text5 = MAQUISTAN(CID1%, NOPEX%, ALTEX%)
  End If
  Text8.TEXT = TRIM$(CAPENOP$(Label4, NOPEX%, CANPROGRA))
  Label18 = Text8
  Label19 = TRIM$(FORMATNUM$(CANPROGRA, "F12.1"))
  '\\\OT-07-0028-RG-25/01/07///
  CADX$ = FORMATNUM(CADSTAN(CID1%, NOPEX%, ALTEX%), "F8.1")
  Label20.Caption = CADX$
  '\\\OT-07-0028-RG-FIN///
  '
End Sub
'
Sub GENECODOPEOF(NROOF$)
     '
     If TRIM$(NROOF$) = "" Then
       CIXXI% = CODINT%(Text10)
       If CIXXI% > 0 Then Call GENECOPER(CIXXI%)
       Exit Sub
     End If
     '
     Call BLANARCH("!OPERINF")
     Call SETULTREG("!OPERINF", 0)
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM OPERFAB "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
     SQLX$ = SQLX$ + "AND StatOper < 8 "                      ' no marca cumplidos"
     SQLX$ = SQLX$ + " ORDER BY NumeOper "
     '
     'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim FABRTMP As ADODB.Recordset
     Set FABRTMP = New ADODB.Recordset
     FABRTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With FABRTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Exit Sub
       End If
       '
       On Error GoTo 0
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
       Do While (.EOF = False)
          CAPROD# = !CANTPROG      'Cantidad Programada
          CAREA# = !CANTREAL       'Cantidad realizada
          SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
          '
          'If SALDO# > 0.05 Then
            NOPE% = !NUMEOPER      'numero de operación
            DOPE$ = !DESCOPER      'Descripcion de la operacion
            MAQUI$ = !MAQUINA      'Máquina
            '
            JJ& = JJ& + 1: JJI% = JJ&
            YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
            NOPX$ = FORMATNUM$(NOPE%, "I3") + "   - "
            NOPX$ = NOPX$ + AJUSTI$(DOPE$, 29) + " - "
            NOPX$ = NOPX$ + AJUSTI$(MAQUI$, 6)
            NOPX$ = NOPX$ + FORMATNUM$(SALDO#, "F12.1")
            Call REPLA(YY$, NOPX$, 3, 58)
              CAPROD1 = CAPROD#
            Call REPLA(YY$, MKS$(CAPROD1), 61, 4)
            Call GRAREG("!OPERINF", YY$, JJ&)
            '
            ' AGREGAR OPERACIONES ALTERNATIVAS
            CIX12% = CODINT%(Text10)
            SQLY$ = "SELECT * FROM SecOper "
            SQLY$ = SQLY$ + "WHERE CodIConj=" & CIX12% & " "
            SQLY$ = SQLY$ + "AND NumeOper = " & NOPE% & " "
            SQLY$ = SQLY$ + "AND AltOper > 0 "
            SQLY$ = SQLY$ + "AND StatOper = 1 "
            SQLY$ = SQLY$ + "ORDER BY AltOper ASC"
            'Set OPERALT = Articulos.OpenRecordset(SQLY$, dbOpenSnapshot)
            Dim OPERALT As ADODB.Recordset
            Set OPERALT = New ADODB.Recordset
            OPERALT.Open FILTSQL$(SQLY$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

            On Error Resume Next
            OPERALT.MoveFirst
            If Err < 1 Then
              On Error GoTo 0
              With OPERALT
                Do While Not .EOF
                   NOPE% = !NUMEOPER      'numero de operación
                   ALTE% = !AltOper
                   DOPE$ = !DESCOPER      'Descripcion de la operacion
                   MAQUI$ = !CodMaq0      'Máquina
                   '
                   JJ& = JJ& + 1: JJI% = JJ&
                   YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
                   NOPX$ = Left$(AJUSTD$(Str$(NOPE%), 3) + "/" + TRIM$(Str$(ALTE%)) + "   ", 6) + "- "
                   NOPX$ = NOPX$ + AJUSTI$(DOPE$, 29) + " - "
                   NOPX$ = NOPX$ + AJUSTI$(MAQUI$, 6)
                   NOPX$ = NOPX$ + FORMATNUM$(SALDO#, "F12.1")
                   Call REPLA(YY$, NOPX$, 3, 58)
                   CAPROD1 = CAPROD#
                   Call REPLA(YY$, MKS$(CAPROD1), 61, 4)
                   Call GRAREG("!OPERINF", YY$, JJ&)
                .MoveNext
                Loop
              End With
              Set OPERALT = Nothing
            End If
            On Error GoTo 0
            '
          'End If
       .MoveNext
       Loop
     End With
     Set FABRTMP = Nothing
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
End Sub
'
Sub GENECOPER(CIJK%)
     '
     Call BLANARCH("!OPERINF")
     Call SETULTREG("!OPERINF", 0)
     Call FRESHECHO("!OPERINF") 'nuevo
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     Call ABREMASTER
     'Set SecopTemp = Articulos.OpenRecordset("SELECT * FROM SecOper WHERE CodIConj=" & CIJK% & " AND StatOper = 1 ORDER BY NumeOper ASC, AltOper ASC", dbOpenSnapshot)
     
     On Error Resume Next
     SECOPTEMP.MoveFirst
     If Err < 1 Then
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
       With SECOPTEMP
         Do While Not .EOF
           CAPROD# = 0      'Cantidad Programada
           CAREA# = 0       'Cantidad realizada
           SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
           '
           NOPE% = !NUMEOPER      'numero de operación
           ALTE% = !AltOper
           NO3$ = FORMATNUM(NOPE%, "F3.0")
           NOPX$ = AJUSTD$(Str$(NOPE%), 3)
             If ALTE% > 0 Then NOPX$ = NOPX$ + "/" + TRIM$(Str$(ALTE%))
           NOPX$ = Left$(NOPX$ + "      ", 6) + "- "
           DOPE$ = !DESCOPER      'Descripcion de la operacion
           MAQUI$ = !CodMaq0      'Máquina
           '
           JJ& = JJ& + 1: JJI% = JJ&
           YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
           NOPX$ = NOPX$ + AJUSTI$(DOPE$, 29) + " - "
           NOPX$ = NOPX$ + AJUSTI$(MAQUI$, 6)
           NOPX$ = NOPX$ + FORMATNUM$(SALDO#, "F12.1")
           Call REPLA(YY$, NOPX$, 3, 58)
              CAPROD1 = CAPROD#
           Call REPLA(YY$, MKS$(CAPROD1), 61, 4)
           Call GRAREG("!OPERINF", YY$, JJ&)
           '
         .MoveNext
         Loop
       End With
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
End Sub

Function CAPENOP$(NROOF$, NOPE1%, CANPROGRA)
  '
  CAPENOP$ = ""
  Call ABREORFAB
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  SQLX$ = SQLX$ + "AND NumeOper = " & NOPE1% & " "
  SQLX$ = SQLX$ + " ORDER BY NumeOper "
  '
  'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim FABRTMP As ADODB.Recordset
  Set FABRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  '
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    CAPROD# = !CANTPROG      'Cantidad Programada
    CAREA# = !CANTREAL       'Cantidad realizada
    SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
    CAPENOP$ = NUMSTRI$(SALDO#, 12, 1, 0)
    CANPROGRA = CAPROD#
  End With
  Set FABRTMP = Nothing
  '
End Function
'
Function MAQDEFIOP$(NROOF$, NOPE1%)
  '
  MAQDEFIOP$ = ""
  Call ABREORFAB
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  SQLX$ = SQLX$ + "AND NumeOper = " & NOPE1% & " "
  SQLX$ = SQLX$ + " ORDER BY NumeOper "
  '
  'Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim FABRTMP As ADODB.Recordset
  Set FABRTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 20
    End If
    On Error GoTo 0
    MAQDEFIOP$ = !MAQUINA
  End With
  Set FABRTMP = Nothing
  '
20 If MAQDEFIOP$ = "" Then
     CIXI% = CODINT%(Text10)
     MAQDEFIOP$ = MAQUISTAN$(CIXI%, NOPE1%)
   End If
  '
End Function

Private Sub Text3_DblClick()
  '
  Call Command6_Click
  '
End Sub

Private Sub Text3_GotFocus()
  Text3.SelStart = 0
  Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Text5.SetFocus
 End If
End Sub

Private Sub Text3_LostFocus()
  NOPX$ = REPLACAR$(TRIM$(Text3), "-", "/")
  NOPX$ = REPLACAR$(NOPX$, ".", "/")
  Text3 = NOPX$
End Sub

Private Sub Text4_DblClick()
   Call Command5_Click
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text1.SetFocus
   End If
End Sub

  '
  '
Private Sub Text5_DblClick()
  '
  Call Command7_Click
  '
End Sub
Private Sub Text5_Change()
  '
  Label10.Caption = ECOARCH$("PADMAQ", Text5.TEXT)
  '
End Sub

Private Sub Text5_GotFocus()
  Text5.SelStart = 0
  Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
  Call CALCUTIEMPO
End Sub

Private Sub Text6_GotFocus()
  Text6.SelStart = 0
  Text6.SelLength = Len(Text6)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   Text7.SetFocus
  End If
End Sub

Private Sub Text6_LostFocus()
   '
   Text6 = HORATEX$(HORANUM(Text6))
   Exit Sub
   '
'   Static MODORA%
'   If MODORA% < 1 Then
'      MODORA% = 1
'      If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'         MODORA% = 2
'      End If
'   End If
'   '
'  If XVALO(Text6.TEXT) > 0 Then
'    Text6.TEXT = REPLACAR$(Text6.TEXT, ".", ":")
'    If InStr(1, Text6.TEXT, ".") = 0 Then
'      If InStr(1, Text6.TEXT, ":") = 0 Then
'        Text6.TEXT = Text6.TEXT + ":0"
'      End If
'    End If
'    '
'    If MODORA% = 2 Then
'       TTXX1$ = FORMATNUM$(REPLACAR$(Text6.TEXT, ":", "."), "F7.2")
'       Text6 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'       Exit Sub
'    End If
'    '
'    On Error Resume Next
'    DXX1 = CDate("0" + Text6.TEXT)
'    If DXX1 <> CDate("0" + Text6.TEXT) Then
'         Call MENSERR("24", "Hora No Válida")
'         Text6 = ""
'         Text6.SetFocus
'    End If
'    On Error GoTo 0
'    Text6.TEXT = Format$(DXX1, "hh:mm")
'  End If
'  '
End Sub

Private Sub Text7_Change()
  Call CALCUTIEMPO
End Sub

Private Sub Text7_GotFocus()
  Text7.SelStart = 0
  Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Text2.Visible = True Then
         Text2.SetFocus
      Else
         Text8.SetFocus
      End If
   End If
End Sub

Private Sub Text7_LostFocus()
   '
   Text7 = HORATEX$(HORANUM(Text7))
   Exit Sub
   '
'  Static MODORA%
'  If MODORA% < 1 Then
'     MODORA% = 1
'     If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
'        MODORA% = 2
'     End If
'  End If
'  '
'  If XVALO(Text7.TEXT) > 0 Then
'    Text7.TEXT = REPLACAR$(Text7.TEXT, ".", ":")
'    If InStr(1, Text7.TEXT, ".") = 0 Then
'      If InStr(1, Text7.TEXT, ":") = 0 Then
'        Text7.TEXT = Text7.TEXT + ":0"
'      End If
'    End If
'    '
'    If MODORA% = 2 Then
'       TTXX1$ = FORMATNUM$(REPLACAR$(Text7.TEXT, ":", "."), "F7.2")
'       Text7 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
'       Exit Sub
'    End If
'    '
'    TTXX77$ = TRIM$(Text7)
'    If Left$(TTXX77$, 2) = "24" Then
'      TTXX77$ = "0" + Mid$(TTXX77$, 3)
'    End If
'    On Error Resume Next
'    DXX1 = CDate(TTXX77$)
'    On Error GoTo 0
'    Text7.TEXT = Format$(DXX1, "hh:mm")
'    '
'  End If
'  '
End Sub

Private Sub Text8_GotFocus()
  Text8.SelStart = 0
  Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Text11.Visible = True Then
        Text11.SetFocus
     Else
        Text9.SetFocus
     End If
  End If
End Sub

Private Sub Text8_LOSTFOCUS()
  '
  Text8.TEXT = TRIM$(NUMSTRI$(XVALO(Text8.TEXT), 12, 1, 0))
  '
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      Text12.SetFocus
   End If
   '
End Sub

Sub CALCUTIEMPO()
   '
   HCOM% = Int(HORANUM(Text6))
   HFIN% = Int(HORANUM(Text7))
   HDESCAN% = Int(HORANUM(Text2))
   DIFDES = DIFHORA(0, HDESCAN%)
   '
   Label25 = ""
   If HFIN% >= HCOM% Then
     TDIF = (DIFHORA(HCOM%, HFIN%) - DIFDES) * 3600
     If TDIF >= 0 Then
        HMINX$ = HORMINSE$(TDIF)
        HMINY$ = Left$(HMINX$, Len(HMINX$) - 3)
        Label25.Caption = HMINY$
     End If
   End If
   '
End Sub

