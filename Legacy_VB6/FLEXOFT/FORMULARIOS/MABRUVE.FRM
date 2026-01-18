VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form MABRUVE 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis de Margen Bruto de Ventas"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
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
      Height          =   855
      Left            =   5460
      Picture         =   "MABRUVE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton Command29 
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
      Height          =   855
      Left            =   5460
      Picture         =   "MABRUVE.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1350
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5460
      Picture         =   "MABRUVE.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2265
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Margen Bruto Mensual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   6
      Top             =   300
      Width           =   5055
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   11
         Top             =   1155
         Width           =   3690
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pantalla"
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
            Left            =   420
            TabIndex        =   13
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impresión"
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
            Left            =   1995
            TabIndex        =   12
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.TextBox MESAN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1470
         TabIndex        =   9
         Text            =   " "
         Top             =   630
         Width           =   2250
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
         Height          =   400
         Left            =   3700
         TabIndex        =   8
         Top             =   620
         Width           =   175
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4095
         Picture         =   "MABRUVE.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   1240
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y año:"
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
         Index           =   10
         Left            =   -840
         TabIndex        =   10
         Top             =   735
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Costos de Reposición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   210
      TabIndex        =   0
      Top             =   2415
      Width           =   5055
      Begin VB.CommandButton Command6 
         Caption         =   "Listar"
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
         Left            =   2625
         TabIndex        =   17
         Top             =   1680
         Width           =   960
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Archivo de E/S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   315
         TabIndex        =   4
         Top             =   525
         Width           =   4425
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
            Height          =   330
            Left            =   210
            TabIndex        =   5
            Top             =   420
            Width           =   4005
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Edit."
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
         Left            =   3780
         TabIndex        =   3
         Top             =   1680
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Export."
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
         Left            =   1470
         TabIndex        =   2
         Top             =   1680
         Width           =   960
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Import."
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
         Left            =   315
         TabIndex        =   1
         Top             =   1680
         Width           =   960
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   5460
      TabIndex        =   18
      Top             =   3675
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   291
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5250
      Picture         =   "MABRUVE.frx":0A68
      Stretch         =   -1  'True
      Top             =   4170
      Width           =   4935
   End
End
Attribute VB_Name = "MABRUVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTREC_Click()
    '
    Call CALMABRU
    '
End Sub

Private Sub command1_Click()
    '
    ARIMPO$ = TRIM$(Text1.Text)
    If ARIMPO$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Archivo de Importación")
        Exit Sub
    End If
    '
    If EXISTE%(ARIMPO$) < 1 Then
        Call MENSERR(24, "No existe Archivo " + ARIMPO$)
        Exit Sub
    End If
    '
    FECO% = FECHANUM("01/01/00")
    LU$ = LUDAT$
    N1% = FILEOPEN%(ARIMPO$, 1, LG%)
    '
    Screen.MousePointer = 11
    While Not EOF(N1%)
        Line Input #N1%, XXX$
        CODL$ = Left$(XXX$, 11)
        COUL$ = REPLACAR$(Right$(XXX$, 12), ",", ".")
        
    
        COD$ = AJUSTI$(CODL$, 16)
        COU = Val(TRIM$(COUL$))
        '
        If COU >= 0.005 Then
           CI% = CODINT%(COD$)
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 REGMA& = CI% + 1
                 YYY$ = REGLEIDO$("MASTER", REGMA&)
                 Call REPLA(YYY$, MKS$(COU), 111, 4)
                 Call REPLA(YYY$, MKI$(FECO%), 115, 2)
                 Call GRAREG("MASTER", YYY$, REGMA&)
              End If
           End If
        End If
    Wend
    Close #N1%
    Call CIERRARCH("MASTER")
    Screen.MousePointer = 1
    Call COMUNI("Importación Completa")
    '
End Sub

Private Sub Command2_Click()
    '
    Call EXCOSUN
    '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    Call ACCDIR("ANAMABRU")
    Command29.Enabled = True
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    Call HELPONLINE("ANAMABRU")
    Command29.Enabled = True
End Sub

Private Sub Command3_Click()
    Call FRESHALL
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    MESAN.Text = MESASEL$
End Sub

Private Sub Command4_Click()
    '
    Call CARCOSUN
    '
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CIERRARCH("*.*")
    Call FINAL("")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Call LICOSUN
    '
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call VERJOBID
    '
End Sub

