VERSION 5.00
Begin VB.Form RECMATRAN09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepcion Material en Transito"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5085
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
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
      Left            =   7140
      TabIndex        =   5
      Top             =   720
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
      Height          =   300
      Left            =   3225
      TabIndex        =   3
      Top             =   720
      Width           =   175
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
      Left            =   900
      TabIndex        =   2
      Top             =   720
      Width           =   2325
   End
   Begin VB.TextBox Text1 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
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
      Height          =   810
      Left            =   7560
      Picture         =   "RECMATRAN09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Graba - Aprueba la Transacción"
      Top             =   3720
      Width           =   830
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ubicacion de Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   3840
      TabIndex        =   22
      Top             =   3480
      Width           =   3525
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
         Left            =   3060
         TabIndex        =   7
         Top             =   645
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
         Left            =   1680
         TabIndex        =   6
         Top             =   645
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "!"
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
         Left            =   3255
         TabIndex        =   8
         Top             =   645
         Width           =   175
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deposito"
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
         Left            =   -650
         TabIndex        =   44
         Top             =   320
         Width           =   1545
      End
      Begin VB.Label Label20 
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
         Left            =   1440
         TabIndex        =   43
         Top             =   280
         Width           =   1980
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
         Left            =   960
         TabIndex        =   40
         Top             =   280
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Ubicación:"
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
         TabIndex        =   23
         Top             =   690
         Width           =   1770
      End
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
      Height          =   1050
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   3645
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricaion"
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
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Retest:"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
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
         Left            =   2400
         TabIndex        =   19
         Top             =   300
         Width           =   1335
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
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   915
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
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Width           =   915
      End
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
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   3765
      Left            =   7560
      ScaleHeight     =   3705
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   -120
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
         Height          =   650
         Left            =   120
         Picture         =   "RECMATRAN09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Listado de Material en Transito"
         Top             =   2220
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "RECMATRAN09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1560
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "RECMATRAN09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   13
         Top             =   3000
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "RECMATRAN09.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   900
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "RECMATRAN09.frx":0D72
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3240
         Width           =   1515
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Ident:"
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
      Left            =   4560
      TabIndex        =   42
      Top             =   765
      Width           =   690
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
      Left            =   -360
      TabIndex        =   41
      Top             =   765
      Width           =   1155
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barras >>>>>>>>:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   39
      Top             =   240
      Width           =   5160
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
      Left            =   5520
      TabIndex        =   38
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito Nº"
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
      Left            =   3840
      TabIndex        =   37
      Top             =   1740
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label5 
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
      Left            =   5760
      TabIndex        =   36
      Top             =   2640
      Width           =   1620
   End
   Begin VB.Label Label3 
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
      Height          =   270
      Left            =   4080
      TabIndex        =   35
      Top             =   2640
      Width           =   1545
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
      Left            =   5760
      TabIndex        =   34
      Top             =   3120
      Width           =   1620
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precinto:"
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
      Left            =   4080
      TabIndex        =   33
      Top             =   3120
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
      Left            =   1440
      TabIndex        =   32
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
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
      Left            =   -240
      TabIndex        =   31
      Top             =   3120
      Width           =   1545
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
      Left            =   0
      TabIndex        =   30
      Top             =   2640
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
      Left            =   1440
      TabIndex        =   29
      Top             =   2640
      Width           =   1995
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
      Left            =   -720
      TabIndex        =   28
      Top             =   1740
      Width           =   1545
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
      Left            =   960
      TabIndex        =   27
      Top             =   1680
      Width           =   2220
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
      Left            =   960
      TabIndex        =   26
      Top             =   2160
      Width           =   6420
   End
   Begin VB.Label Label13 
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
      Left            =   6360
      TabIndex        =   25
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recepcion"
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
      Left            =   4680
      TabIndex        =   24
      Top             =   1350
      Width           =   1545
   End
End
Attribute VB_Name = "RECMATRAN09"
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
   Label10 = ""
   Label13 = ""
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command11_Click()
    '
    Command11.Enabled = False
    '
    DEPTRAN% = CONTROL$("BOLRECEP", "DEPTRANS")  ' depósito de transito
    If DEPTRAN% < 1 Then
        Call COMUNI("Depósito de Transporte no Configurado")
        GoTo 99
    End If
    CONDI$ = "DEPOUBIC = " & DEPTRAN%
    Call CARGALISEL("!MASTRAN", "LOTINGRE", "COD_INTE/C16+MASTER/A40", CONDI$, "DISTINCT")
    Call SELECHO("!MASTRAN")
    Text8 = TRIM$(VALACT1("!MASTRAN"))
    '
99  Command11.Enabled = True
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   CIXI% = CODINT(TRIM$(Label28))
   LOTE$ = TRIM$(Label18)
   DEPP% = XVALO(Label16)
   If XVALO(CIXI%) > 0 Then
      Call UBIRECOM(CIXI%, LOTE$, DEPP%)
      Call CommSel("!TAUBICA", Command2, Text3)
   End If
   '
   Command2.Enabled = True
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    DEPMATRAN$ = CONTROL$("BOLRECEP", "DEPTRANS")
    ' Por defecto carga  deposito material en transito
    Call FINAL("?MATETRANS(" + TRIM$(DEPMATRAN$) + ")")
    Command20.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    CIXI% = CODINT%(Text8)
    If CIXI% < 1 Then
        Call COMUNI("Debe Seleccionar el Codigo")
        GoTo 99
    End If
    DEPTRAN% = CONTROL$("BOLRECEP", "DEPTRANS")  ' depósito de transito
    CONDI$ = "DEPOUBIC = " & DEPTRAN% & " and Cod_Inte = " & CIXI%
    Call CARGALISEL("!LOTRAN", "LOTINGRE", "IDENLOTE/A16;FechAlta/D8;CanAlta/F14." & CNTDC$ & ";(CanAlta-CanConsu)/F14." & CNTDC$ & ";LoteCol_Prov/A16", CONDI$, "NOMESS")
    Call SELECHO("!LOTRAN")
    Text2 = TRIM$(VALACT1("!LOTRAN"))
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    DEPP% = XVALO(Label16)
    CIXI% = CODINT%(Label28)
    If XVALO(CIXI%) > 0 Then
       Call GENTAUBI(DEPP%)
       If ULTREG&("!TAUBICA") < 1 Then GoTo 99
       Call SELECHO("!TAUBICA")
        
       UBISELE$ = VALACT1$("!TAUBICA")
       If UBISELE$ <> "" Then
          Text3 = UBISELE$
       End If
    End If
    '
99  Command4.Enabled = True
End Sub


Private Sub Command5_Click()
    '
    Command5.Enabled = False
    '
    Static DOCULAR$
    If DOCULAR$ = "" Then
        DOCULAR$ = TRIM$(Label13)
        NUORIT% = 0
    End If
    '
    CIXI% = CODINT%(Label28)
    LOTE$ = TRIM$(Text2)
    If CIXI% < 1 Then
        COMUNI ("Codigo de Articulo No Válido")
        GoTo 99
    End If
    '
    If CantRegistros&("LOTINGRE", "IDENLOTE = '" & LOTE$ & "'") < 1 Then
        COMUNI ("Lote Ingresado No Válida")
        Text2.SetFocus
        GoTo 99
    End If
    '
    DEPACT% = XVALO(Label16)
    If CantRegistros("UBIPRE", "DEPOUBIC = " & DEPACT%) > 0 Then ' HAY UBICACIONES DEFINIDAS
      If CantRegistros&("UBIPRE", "DEPOUBIC = " & DEPACT% & "AND Codiubic = '" & TRIM$(Text3) & "'") < 1 Then
        COMUNI ("Ubicación No Válida")
        GoTo 99
      End If
    End If
    '
    UBICAFIS$ = TRIM$(Text3)                  'Ubicacion Propuesta
    Destino$ = TRIM$(Label2)                         'Destino
    '
    If PERMITEUBI%(CIXI%, Destino$, UBICAFIS$) <> 1 Then    'VALIDA LA UBICACION SELECCIONADA
        Text3.SetFocus
        GoTo 99
    End If
    '
    If COMALTER%("Revise los Datos Ingresados.\¿ Confirma la Recepción ?\\Si, Confirmar\No, Cancelar") <> 1 Then GoTo 99
    '
    FEMOVI% = HOY(HO$)
    COMP$ = CONTROL("", "CODMBREC")        'Tipo de Ingreso BR o IR
    If TRIM$(COMP$) = "" Then COMP$ = "BR"
    DOCUPRI$ = COMP$ & "." & TRIM$(Label13)
    DOCUSEC$ = TRIM$(Label10)                       'NUMERO DEL REMITO DE TRASLADO
    REFEREN1$ = "Entrega de Material en Transito"
    REFEREN2$ = "Recepcion de Material en Transito"
    DEPP% = XVALO(Label16)  'DEPOSITO DE INGRESO
    DEPTRANS% = CONTROL$("BOLRECEP", "DEPTRANS")
    CANTI = XVALO(Label5)
    '
    If DOCUPRI$ = DOCULAR$ Then
        NUORIT% = NUORIT% + 1
    End If
    '
    'REGISTRO MOVIMIENTO DE SALIDA EN DEPOSITO DE MATERIAL EN TRANSITO
    Call MOVISTOK(FEMOVI%, CIXI%, LOTE$, "TR", DOCULAR$, DOCULAR$, DOCUSEC$, DOCUSEC$, REFEREN1$, NUORIT%, 0, 0, 0, DEPTRANS%, (-1) * CANTI)
    'REGISTRO MOVIMIENTO DE ENTRADA A DEPOSITO DE RECEPCION
    Call MOVISTOK(FEMOVI%, CIXI%, LOTE$, "TR", DOCUPRI$, DOCUPRI$, DOCUSEC$, DOCUSEC$, REFEREN2$, NUORIT%, 0, 0, 0, DEPP%, CANTI)
    CONDI$ = "COD_INTE = " & CIXI% & "AND IDENLOTE ='" & LOTE$ & "'"
    '
    Call ACTUREGISTRO("LOTINGRE", "DEPOUBIC", CONDI$, DEPP%, REGS&, "NOMESS")       'Actualizo deposito
    Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBICAFIS$, REGS&, "NOMESS") 'Actualizo ubicacion
    ' LINE DE ABAJO COMENTADA PARA QUE NO CAMBIE EL ESTADO DEL CAMPO LIBERA Y MANTENGA EL QUE TRAE ORIGINALMENTE
'    Call ACTUREGISTRO("LOTINGRE", "LIBERA", CONDI$, "CUAR", REGS&, "NOMESS")        'Actualizo estado
    '
    Call ACTUREGISTRO("UBIPRE", "Status", "codiUbic = '" & UBICAFIS$ & "'", 1, nroREG&)
    '
    Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
    '
    Call PRINSTIKAPRO1(CIXI%, LOTE$)
    '
    COMUNI ("Recepcion de Material en Transito Procesada")
    '
    Call LIMPTEXT
    Text8 = ""
    Text1 = ""
    '
99  Command5.Enabled = True
    '
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    FIKARDEX09.Text1 = TRIM$(Label28)
    FIKARDEX09.Show 1
    Command8.Enabled = True
End Sub

Private Sub Text1_Change()
    '
    If Text1 <> "" Then Text8 = ""
    CIXI% = CODINT%(Mid$(TRIM$(Text1), 1, 16))
    If CIXI% > 1 Then
        LOTE$ = TRIM$(Mid$(Text1, 17, 12))
        Label28 = CODEXT$(CI%)
        Label30 = DESCRIT$(CI%)
        If CantRegistros&("LOTINGRE", "IDENLOTE = '" & LOTE$ & "'AND COD_INTE = " & CIXI%) > 0 Then
            Call CARGDATOS(CIXI%, LOTE$)
            Exit Sub
        End If
        Label18 = ""
        Label2 = ""
        Label5 = ""
        Label4 = ""
        Label38 = ""
        Label37 = ""
        Label36 = ""
    End If
    Call LIMPTEXT
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii <> 13 Or TRIM$(Text1) = "" Then
      Exit Sub
    End If
    '
    Text1 = CONVERLECBAR(Text1)
    Call FRACLECBAR(Text1, Text8, Text2, Label5)
    '
    'Text2 = ""
    'Text8 = ""
    '
    DEPTRANS% = CONTROL$("BOLRECEP", "DEPTRANS")   ' depósito de transito
    '
  
    '
    'CIXI% = CODINT%(TRIM$(Mid$(TENTRA$, 1, 16)))
    'LOTE$ = TRIM$(Mid$(TENTRA$, 17, 12))
    '
    CIXI% = CODINT(Text8)
    LOTE$ = TRIM$(Text2)
    CONDI$ = "COD_INTE = " & CIXI% & "AND IDENLOTE ='" & LOTE$ & "'"
    '
    DEPUBIC% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$))
    '
    If DEPUBIC% <> DEPTRANS% Then
        Call MENSERR(24, "El Lote a Ingresar no es Válido o no se Encuentra en Tránsito")
        Text1 = ""
        Exit Sub
    End If
    '
    Call CARGDATOS(CIXI%, LOTE$)
    '
End Sub
Sub LIMPTEXT()
    '
    Text3 = ""
    Text2 = ""
    Label28 = ""
    Label30 = ""
    Label18 = ""
    Label2 = ""
    Label5 = ""
    Label4 = ""
    Label38 = ""
    Label37 = ""
    Label36 = ""
    '
End Sub

Private Sub Text2_Change()
    '
    CIXI% = CODINT%(Text8)
    LOTE$ = TRIM$(Text2)
    
    Call CARGDATOS(CIXI%, LOTE$)
    '
End Sub

Private Sub Text2_DblClick()
    Call Command3_Click
End Sub

Private Sub Text8_Change()
    Text2 = ""
    If Text8 <> "" Then Text1 = ""
    CIXI% = CODINT%(Text8)
    If CIXI < 1 Then
        Call LIMPTEXT
        Exit Sub
    End If
    Label28 = CODEXT$(CIXI%)
    Label30 = DESCRIT$(CIXI%)
    '
End Sub

Private Sub Text8_DblClick()
    Call Command11_Click
End Sub

Sub CARGDATOS(CI%, LOTE$)
    '
    Label28 = CODEXT$(CI%)
    Label30 = DESCRIT$(CI%)
    'Label18 = LOTE$
    '
    CONDI$ = "COD_INTE = " & CI% & "AND IDENLOTE ='" & LOTE$ & "'"
    Label18 = SAFETEXT(VALOBADA("LOTINGRE", "LoteCol_Prov", CONDI$))                       'lote proveedor
    Label4 = VALOBADA("LOTINGRE", "PRECINTO", CONDI$)                                      'asigna precinto
    Label5 = FORMATNUM$(XVALO(VALOBADA("LOTINGRE", "(CANALTA - CANCONSU)", CONDI$)), "F13." & CNTDC)   'saldo del lote
    Label38 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechFab", CONDI$)))                    'fecha de fabricacion
    Label37 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "Fechco_Rep", CONDI$)))                 'fecha de retest
    Label36 = FECHATEX$(CVINT(VALOBADA("LOTINGRE", "FechVenc", CONDI$)))                   'fecha de vencimiento
    Label2 = VALOBADA("LOTINGRE", "COD_DESTI", CONDI$)
    '
End Sub
