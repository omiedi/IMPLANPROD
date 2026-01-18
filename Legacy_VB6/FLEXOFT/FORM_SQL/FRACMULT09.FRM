VERSION 5.00
Begin VB.Form FRACMULT09 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fraccionamiento Multiple"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2655
      Left            =   10800
      ScaleHeight     =   2595
      ScaleWidth      =   2760
      TabIndex        =   4
      Top             =   5280
      Width           =   2820
      Begin VB.CommandButton Command2 
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
         Height          =   600
         Left            =   135
         Picture         =   "FRACMULT09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Graba - Aprueba la Transacción"
         Top             =   900
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   135
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   5
         Top             =   105
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   570
         Width           =   1545
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
         TabIndex        =   12
         Top             =   570
         Width           =   1545
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -280
         Picture         =   "FRACMULT09.frx":030A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1650
      Left            =   10800
      ScaleHeight     =   1590
      ScaleWidth      =   3240
      TabIndex        =   43
      Top             =   0
      Width           =   3300
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
         Height          =   650
         Left            =   135
         Picture         =   "FRACMULT09.frx":222C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Salir"
         Top             =   120
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
         Height          =   650
         Left            =   120
         Picture         =   "FRACMULT09.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   820
         Width           =   690
      End
      Begin VB.Label Label28 
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
         TabIndex        =   47
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label26 
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
         TabIndex        =   46
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label25 
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
         TabIndex        =   45
         Top             =   1095
         Width           =   1545
      End
      Begin VB.Label Label24 
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
         TabIndex        =   44
         Top             =   45
         Width           =   1335
      End
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
      Left            =   2055
      TabIndex        =   25
      Top             =   1215
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
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   24
      Top             =   1215
      Width           =   600
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
      Left            =   6825
      TabIndex        =   19
      Top             =   1200
      Width           =   175
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
      Left            =   1455
      TabIndex        =   0
      Top             =   315
      Width           =   2160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fraccionar del Lote "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   1440
      TabIndex        =   16
      Top             =   5160
      Width           =   7605
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5715
         TabIndex        =   41
         Top             =   1140
         Width           =   1650
      End
      Begin VB.TextBox Text6 
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
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descrip.:"
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
         Left            =   -105
         TabIndex        =   42
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Asig.:"
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
         Left            =   -300
         TabIndex        =   40
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label7 
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
         Left            =   1080
         TabIndex        =   39
         Top             =   1125
         Width           =   1755
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
         Left            =   1080
         TabIndex        =   38
         Top             =   765
         Width           =   6300
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/P.:"
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
         Left            =   300
         TabIndex        =   37
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label5 
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
         TabIndex        =   36
         Top             =   390
         Width           =   1755
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
         Left            =   5610
         TabIndex        =   35
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C. Frac.:"
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
         Left            =   2475
         TabIndex        =   18
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label19 
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
         Height          =   210
         Left            =   4140
         TabIndex        =   17
         Top             =   435
         Width           =   1335
      End
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "O/P Asignada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Width           =   11490
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   11000
         TabIndex        =   52
         Top             =   405
         Width           =   30
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   9650
         TabIndex        =   50
         Top             =   405
         Width           =   30
         Begin VB.Frame Frame5 
            Caption         =   "Frame3"
            Height          =   2685
            Left            =   5000
            TabIndex        =   51
            Top             =   0
            Width           =   30
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   1860
         TabIndex        =   32
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   6900
         TabIndex        =   31
         Top             =   390
         Width           =   30
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   8300
         TabIndex        =   30
         Top             =   405
         Width           =   30
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   2750
         Left            =   3615
         TabIndex        =   29
         Top             =   420
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
         Height          =   2580
         ItemData        =   "FRACMULT09.frx":2680
         Left            =   75
         List            =   "FRACMULT09.frx":2682
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   555
         Width           =   11240
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"FRACMULT09.frx":2684
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
         Left            =   75
         TabIndex        =   34
         Top             =   285
         Width           =   11240
      End
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
      Left            =   2325
      TabIndex        =   27
      Top             =   1230
      Width           =   1890
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
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
      Left            =   510
      TabIndex        =   26
      Top             =   1275
      Width           =   840
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
      Left            =   8025
      TabIndex        =   23
      Top             =   1215
      Width           =   1440
   End
   Begin VB.Label Label12 
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
      Height          =   210
      Left            =   6900
      TabIndex        =   22
      Top             =   1275
      Width           =   1080
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
      Left            =   4305
      TabIndex        =   21
      Top             =   1245
      Width           =   915
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
      Left            =   5250
      TabIndex        =   20
      Top             =   1200
      Width           =   1650
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
      Left            =   6780
      TabIndex        =   11
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
      Left            =   8460
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   735
      Width           =   8025
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
      TabIndex        =   7
      Top             =   375
      Width           =   1545
   End
End
Attribute VB_Name = "FRACMULT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IndexGuia As Integer

Private Sub Command1_Click()
   '
   '
   For UI& = 0 To List1.ListCount
      LINEA$ = List1.List(UI&)
      CODIGO$ = TRIM$(Text7)
      CIXI% = CODINT(CODIGO$)
      CANCON# = XVALO(Mid$(LINEA$, 91, 13))
      If CANCON# >= FRACMIN(CIXI%) Then
         Call MENSERR(24, "Imposible Cerrar. Debe Registrar Movimientos\con Click sobre Botón 'Save'")
         Exit Sub
      End If
   Next UI&
   '
10 Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command2_Click()
    '
   Command2.Enabled = False
   Screen.MousePointer = 11
   NDEP% = XVALO(Text1)
   FF% = FECHANUM(Label13)
   NORI% = 0
   SUMAFRAC# = 0
   CODIGO$ = TRIM(Text7)
   CIXI% = CODINT(CODIGO$)
   DESCRI$ = DESCRIT0(CIXI%)
   IDLOT$ = Label36
   '
   DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
   If NDEP% = DEPOTRANS% Then
       Call COMUNI("Depósito Seleccionado No Válido\Configurado como Material en Transito ")
       GoTo 99
   End If

   For UI& = 0 To List1.ListCount - 1
      LINEA$ = List1.List(UI&)
      '
      OP$ = Mid(LINEA$, 1, 16)
      
      CANTASIG# = XVALO(Mid(LINEA$, 65, 13))
      CANCON# = XVALO(Mid(LINEA$, 91, 13))
    
      '
      If CANCON# >= FRACMIN(CIXI%) Then
        CMO$ = "CF"
        DOPRI$ = "OP-" + OP$
        CONDI$ = "NUORPROD ='" & OP$ & "'"
        SF$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$, "NOMESS")
        DOSEC$ = SF$
        REFE$ = "Fracc.Central de Pesadas"
        SUMAFRAC# = SUMAFRAC# + CANCON#
        '
        CONDI$ = "IDENLOTE = '" & IDLOT$ & "' AND DESTINO ='" & SF$ & "'"
        IDREQ$ = VALOBADA("BALAREQMO", "IDENREQUE", CONDI$) 'TRIM$(Mid$(LINEA$, 151, 32))
        '
        NORI% = NORI% + 1
        Call MOVISTOK(FF%, CIXI%, IDLOT$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, 0, 0, 0, NDEP%, (-1) * CANCON#, , , , , IDREQ$)
      End If
      '
   Next UI&
   '
   SALDO# = XVALO(Label11) - SUMAFRAC#
   'Imprime Etiqueta por Parte Remanente
   If SALDO# > 0 Then
       Call PRINSTIKAPRO1(CIXI%, IDLOT$, NDEP%, SALDO#)
   End If
   '
   Call ARMALILOT(CODINT(Text7.TEXT), NDEP%, , , 1)
   Call LIMPCONMAT
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub
'
Private Sub Command30_Click()
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Text1 = VALACT1$("TADEPOSI")
End Sub

Private Sub Command4_Click()
   IndexGuia = List1.ListIndex
   LINEA$ = List1.TEXT
   '
   If TRIM(Label4) = "" Then
      Call COMUNI("Debe Seleccionar un Registro de la Lista")
      GoTo 99
   End If
   '
   SUMA# = XVALO(Text6)
   For I = 0 To List1.ListCount - 1
     CANCON# = XVALO(Mid(LINEA$, 91, 13))
     SUMA# = SUMA# + CANCON#
   Next I
   '
   If SUMA# > XVALO(Label11) Then
      Call COMUNI("Total de Cant. Fraccionada es Mayor a Cant. de Saldo")
      GoTo 99
   End If
   '
   OP$ = Label5
   CODIGO$ = Text7
   CIXI% = CODINT(CODIGO$)
   DESCRI$ = Label16
   CANASI# = XVALO(Label7)
   FRA# = XVALO(Text6)
   LOTE$ = Label36
   DEP% = XVALO(Text1)
   '
   If XVALO(Text6) < 0 Then
      COMUNI ("La Cant. Fraccionada Debe ser Mayor a Cero")
      GoTo 99
   End If
   '
   'Imprime Etiqueta por Parte Fraccionada
   Call PRINSTIKAPRO1(CIXI%, LOTE$, DEP%, FRA#)
   '
   FRACC$ = FORMATNUM(FRA#, "F13." & CNTDC$(CIXI%))
   CANTASIG$ = FORMATNUM(CANASI#, "F13." & CNTDC$(CIXI%))
   '
   Call REPLA(LINEA$, OP$, 1, 16)
   Call REPLA(LINEA$, CODIGO$, 18, 16)
   Call REPLA(LINEA$, DESCRI$, 35, 29)
   Call REPLA(LINEA$, CANTASIG$, 65, 13)
   Call REPLA(LINEA$, FRACC$, 91, 13)
   List1.List(IndexGuia) = LINEA$
   
99 Command3.Enabled = True
    
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Text7 = VALACT1$("MASTER")
       A = CODINT(Text7)
    End If
    '
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
    If CODINT(Text7) < 1 Then
       Call COMUNI("Codigo No Valido")
       GoTo 99
    End If
    '
    If XVALO(Text1) < 1 Then
       Call COMUNI("Deposito No Valido")
       GoTo 99
    End If
    '
    Call SELECHO("!LILOTSE")
    If VALACT1$("!LILOTSE") <> "" Then
       Label36.Caption = VALACT1$("!LILOTSE")
    End If
99
    '
End Sub

Private Sub Form_Load()
   Label13 = FECHATEX(HOY(HO$))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   For UI& = 0 To List1.ListCount
      LINEA$ = List1.List(UI&)
      CODIGO$ = TRIM$(Text7)
      CIXI% = CODINT(CODIGO$)
      CANCON# = XVALO(Mid$(LINEA$, 91, 13))
      If CANCON# >= FRACMIN(CIXI%) Then
         Cancel = 1
         Call MENSERR(24, "Imposible Cerrar. Debe Registrar Movimientos\con Click sobre Botón 'Save'")
         Exit Sub
      End If
   Next UI&
   Cancel = 0
   '
End Sub

Private Sub Label36_Change()
  '
   Call LIMPFRA
   CIXXI% = XVALO(CODINT(Text7))
   Label24 = UBILOTE$(CIXXI%, Label36, 1)
   Label11 = FORMATNUM(SALNOCONS(Label36, CIXXI%), "F13." & CNTDC$(CODINT%(Text7)))
   LOTE$ = Label36
   If LOTE$ <> "" Then
      List1.Clear
      SQLX$ = "SELECT * FROM BALAREQMO "
      SQLX$ = SQLX$ + " WHERE (IDENLOTE ='" & LOTE$ & "' OR IDENLOTE = '') "
      SQLX$ = SQLX$ + " AND KILO_ASIG > 0"
      Dim OPROD As ADODB.Recordset
      Set OPROD = FLEXCONN.Execute(FILTSQL$(SQLX$))
      
      With OPROD
        Do While Not (.EOF)
           LINEA$ = ""
           '
           SF$ = SAFETEXT$(!DESTINO)
           IDREQUE$ = SAFETEXT$(!IdenReque)
           CONDI$ = "IDSUBOR='" & SF$ & "'"
           CIXI% = VALOBADA("ORDEFABR", "COD_INTE", CONDI$)
           OP$ = TRIM$(VALOBADA("ORDEFABR", "NUORPROD", CONDI$))
           If OP$ <> "" Then
              '
              CODIGO$ = CODEXT$(CIXI%)
              DESCRI$ = DESCRIT0$(CIXI%)
              CANTASIG$ = FORMATNUM$(XVALO(!Kilo_Asig), "F13." & CNTDC$(CIXXI%))
              '
              CONDI$ = "Cod_Inte = " & CIXXI%
              CONDI$ = CONDI$ + " AND IdenLote = '" & LOTE$ & "'"
              CONDI$ = CONDI$ + " AND IdenMovi = '" & IDREQUE$ & "'"
              CANYACON# = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
              CANYACON# = CANYACON# - XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
              CANTYACO$ = FORMATNUM$(CANYACON#, "F13." & CNTDC$(CIXXI%))
              '
              'FRACC$ = FORMATNUM$(12, "F10." & CNTDC$)
              '
              Call REPLA(LINEA$, OP$, 1, 16)
              Call REPLA(LINEA$, CODIGO$, 18, 16)
              Call REPLA(LINEA$, DESCRI$, 35, 29)
              Call REPLA(LINEA$, CANTASIG$, 65, 13)
              Call REPLA(LINEA$, CANTYACO$, 78, 13)
              List1.AddItem LINEA$
           End If
        .MoveNext
        Loop
      End With
   End If
   '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text6.SetFocus
   End If
End Sub

Private Sub Text5_LostFocus()
    Text5 = FORMATNUM$(XVALO(Text5), "F13." & CNTDC$(CODINT%(Text7)))
End Sub

Private Sub Label4_Change()
 Label16 = DESCRIT0(CODINT(Label4))
End Sub



Private Sub List1_DblClick()
     IndexGuia = List1.ListIndex
     LINEA$ = List1.List(IndexGuia)
     '
     OP$ = Mid(LINEA$, 1, 16)
     CODIGO$ = Mid(LINEA$, 18, 16)
     DESCRI$ = Mid(LINEA$, 35, 29)
     CANTASIG$ = Mid(LINEA$, 65, 13)
     CANTYACO$ = Mid(LINEA$, 78, 13)
     FRACC$ = Mid$(LINEA$, 91, 13)
     '
     CI% = CODINT(Text7)
     Label5 = OP$
     Label4 = CODIGO$
     Label7 = FORMATNUM(CANTASIG$, "F13." & CNTDC$(CI%))
     Text6 = FORMATNUM(FRACC$, "F13." & CNTDC$(CI%))
     '
End Sub

Private Sub Text1_Change()
   DEP% = XVALO(Text1)
   Label38 = ECOARCH$("TADEPOSI", Chr$(DEP%))
   Call ARMALILOT(CODINT(Text7.TEXT), DEP%, , , 1)
End Sub

Private Sub Text6_LostFocus()
  Text6 = FORMATNUM(XVALO(Text6), "F13." & CNTDC$(CODINT%(Text7)))
End Sub

Private Sub Text7_Change()
   '
   Call LIMPCONMAT
   DEP% = XVALO(Text1)
   CI% = CODINT%(Text7)
   Call ARMALILOT(CI%, DEP%, , , 1)
   '
   If (CI%) > 0 Then
      CONDI$ = "COD_INTE = " & Str(CI%)
      Label30 = VALOBADA("MASTER", "Descripcion", "Codint = " & Str(CI%))
   End If
   '
End Sub

Private Sub Text7_DblClick()
   Call Command5_Click
End Sub

Sub LIMPCONMAT()
    '
    Label36 = ""
    Label24 = ""
    Label11 = ""
    Text5 = ""
    Text6 = ""
    List1.Clear
    Call LIMPFRA
    '
End Sub
Sub LIMPFRA()
    Label5 = ""
    Label4 = ""
    Label16 = ""
    Label7 = ""
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command5_Click
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command8.SetFocus
   End If
End Sub

Private Sub Text7_LostFocus()
    If Text7 <> "" Then
       CI% = CODINT(Text7)
       Text7 = CODEXT(CI%)
    End If
End Sub
