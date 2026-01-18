VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STOCOMPO07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Stocks de Componentes"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3030
      Index           =   0
      Left            =   8700
      TabIndex        =   15
      Top             =   1900
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3030
      Index           =   4
      Left            =   7110
      TabIndex        =   7
      Top             =   1900
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3030
      Index           =   3
      Left            =   5470
      TabIndex        =   6
      Top             =   1900
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3030
      Index           =   1
      Left            =   1890
      TabIndex        =   5
      Top             =   1900
      Width           =   15
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8970
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -30
         TabIndex        =   12
         Top             =   810
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -10
         TabIndex        =   11
         Top             =   400
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6080
         TabIndex        =   10
         Top             =   400
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1470
         TabIndex        =   9
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7050
         TabIndex        =   8
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1470
         TabIndex        =   4
         Top             =   720
         Width           =   7350
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      ItemData        =   "STOCOMPO07.frx":0000
      Left            =   180
      List            =   "STOCOMPO07.frx":0007
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   2160
      Width           =   8865
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   5130
      Left            =   9240
      ScaleHeight     =   5100
      ScaleWidth      =   2025
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   18
         Top             =   4560
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "STOCOMPO07.frx":0068
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "COMPONENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   8985
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"STOCOMPO07.frx":01B2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   60
         TabIndex        =   14
         Top             =   320
         Width           =   8865
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "STOCOMPO07.frx":0247
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STOCOMPO07.frx":047B
      TabIndex        =   16
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6810
      Y1              =   4650
      Y2              =   4650
   End
End
Attribute VB_Name = "STOCOMPO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture3)
    '
End Sub

Private Sub Label4_Change()
   '
   'DESCRIPCION DE ARTICULO SELECCIONADO
   'DATOS DE ENCABEZADO
   CI% = CODINT%(Label4)
   Label6 = DESCRIT$(CI%)
   CODDX$ = Label4
   Label5 = FORMATNUM$(STOCKACT(CI%), "F15." & CNTDC$(CI%) & "")
   
   'CARGA DE STOCK EN EL LIST
   Call LEEEXPLO(CODDX$, 1)
   '
   List1.Clear
   '
   For J% = 1 To CAIT%
          CIY% = CIJ%(J%)
          CODIGOEXT$ = CODEXT$(CIY%)
          DESCRIPCIOX$ = DESCRIT$(CIY%)
          STOCKAC = STOCKACT(CIY%)
          STOCKDI = STOCKDIS(CIY%)
          If STOCKDI < 0 Then STOCKDI = 0
          '
          Call REPLA(TX$, CODIGOEXT$, 1, 16)
          Call REPLA(TX$, DESCRIPCIOX$, 17, 32)
          Call REPLA(TX$, FORMATSTAN$(STOCKAC, "F15", CIY%), 51, 15)
          Call REPLA(TX$, FORMATSTAN$(STOCKDI, "F15", CIY%), 67, 15)
          '
          List1.AddItem TX$
   Next J%

    '

End Sub

Private Sub List1_DblClick()
   Label4 = TRIM$(Left$(List1.TEXT, 16))
End Sub
