VERSION 5.00
Begin VB.Form ALTAPORFAM 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Código de Artículo por Familia"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7950
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   16
      Top             =   3570
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   7980
      Picture         =   "ALTAPORFAM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Acceso a Maestro de Artículos"
      Top             =   1890
      Width           =   750
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
      Height          =   330
      Left            =   1680
      TabIndex        =   13
      Top             =   1155
      Width           =   6105
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Códigos Existentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   105
      TabIndex        =   12
      Top             =   1575
      Width           =   7680
      Begin VB.PictureBox Picture1 
         Height          =   2220
         Left            =   105
         ScaleHeight     =   2160
         ScaleWidth      =   7410
         TabIndex        =   18
         Top             =   450
         Width           =   7470
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1890
            IntegralHeight  =   0   'False
            ItemData        =   "ALTAPORFAM.frx":030A
            Left            =   -10
            List            =   "ALTAPORFAM.frx":030C
            TabIndex        =   20
            Top             =   315
            Width           =   7440
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Código         Descripción"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   30
            TabIndex        =   19
            Top             =   105
            Width           =   7425
         End
      End
   End
   Begin VB.CommandButton Command3 
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
      Height          =   770
      Left            =   7980
      Picture         =   "ALTAPORFAM.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.CommandButton BOTOK 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   7980
      Picture         =   "ALTAPORFAM.frx":0458
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Alta del Nuevo Código"
      Top             =   1050
      Width           =   750
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
      Height          =   330
      Left            =   3150
      Picture         =   "ALTAPORFAM.frx":0762
      TabIndex        =   1
      Top             =   210
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
      Height          =   330
      Left            =   2730
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7380
      Picture         =   "ALTAPORFAM.frx":0A6C
      Top             =   3885
      Width           =   2010
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      Caption         =   "Denominación"
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
      Left            =   -840
      TabIndex        =   14
      Top             =   1250
      Width           =   2430
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      Caption         =   "Tipo"
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
      Left            =   3700
      TabIndex        =   9
      Top             =   735
      Width           =   540
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4410
      TabIndex        =   8
      Top             =   630
      Width           =   435
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4410
      TabIndex        =   7
      Top             =   210
      Width           =   435
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4935
      TabIndex        =   6
      Top             =   630
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      Caption         =   "U.Med."
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
      Left            =   3255
      TabIndex        =   5
      Top             =   315
      Width           =   1065
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4935
      TabIndex        =   4
      Top             =   210
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      Caption         =   "Codigo de Familia"
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
      Left            =   210
      TabIndex        =   3
      Top             =   315
      Width           =   2430
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   105
      TabIndex        =   2
      Top             =   630
      Width           =   3220
   End
End
Attribute VB_Name = "ALTAPORFAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTOK_Click()
    '
    BOTOK.Enabled = False
    '
    If Len(Text1) <> 2 Then
      Call COMUNI("Código de Familia No Válido")
      GoTo 99
    End If
    '
    X$ = ""
    X$ = FILTERGETRECORD$("TAFAMIL", 1, TRIM$(Text1))
    If Len(X$) < 4 Then
      Call COMUNI("Código de Familia No Válido")
      GoTo 99
    End If
    '
    If TRIM$(Text2) = "" Then
      Call COMUNI("Falta Descripción del Artículo")
      GoTo 99
    End If
    '
    Screen.MousePointer = 11
    PRONULI& = 1
    ABUS$ = UCase$(TRIM$(Text1)) + "-"
    FIN& = NUMAS%
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MASTER", U& + 1)
      CODILEIDO$ = UCase$(TRIM$(Mid$(XX$, 77, 15)))
      If Left$(CODILEIDO$, 3) = ABUS$ Then
        NUITN& = Val(Mid$(CODILEIDO$, 4))
        If NUITN& >= PRONULI& Then PRONULI& = 1 + NUITN&
      End If
    Next U&
    '
    PRONUX$ = TRIM$(Str$(PRONULI&))
    While Len(PRONUX$) < 5
      PRONUX$ = "0" + PRONUX$
    Wend
    CODINUE$ = ABUS$ + PRONUX$
    '
    DESCRINUE$ = Text2
    TIMA% = Val(Label7)
    UNID$ = TRIM$(Label6)
    Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
    If OKX% < 1 Then
      Call MENSERR(24, "El Alta de Codigo\no pudo Completarse.\  \Ejecute Control de Maestro\y Consulte a su Soporte Técnico.")
    Else
      Call Text1_Change
      Text2 = ""
      Hide
    End If
    '
99  Screen.MousePointer = 1
    BOTOK.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    AMASTO00.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    X$ = REGSEL$("TAFAMIL")
    '
    If Len(X$) > 4 Then
      Text1 = Mid$(X$, 1, 2)
    End If
    '
End Sub

Private Sub Label6_Change()
    '
    X$ = ECOARCH("UMEDIDA", TRIM$(Label6))
    Label3 = TRIM$(X$)
    '
End Sub

Private Sub Label7_Change()
    '
    X$ = ECOARCH("TATIMAT", Chr(Val(Label7)))
    Label5 = TRIM$(X$)
    '
End Sub

Private Sub Text1_Change()
    '
    Text2 = ""
    Label1 = ""
    Label6 = ""
    Label7 = ""
    List4.Clear
    '
    If ECOARCH$("TAFAMIL", Text1) = "" Then
      Exit Sub
    End If
    '
    X$ = REGBLAN$("TAFAMIL")
    X$ = FILTERGETRECORD$("TAFAMIL", 1, TRIM$(Text1))
    '
    Label1 = TRIM$(Mid$(X$, 3, 27))
    Label6 = TRIM$(Mid$(X$, 31, 2))
    Label7 = Asc(Mid(X$, 30, 1))
    ABUS$ = UCase$(TRIM$(Text1)) + "-"
    FIN& = NUMAS%
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MASTER", U& + 1)
      CODILEIDO$ = UCase$(TRIM$(Mid$(XX$, 77, 15)))
      If Left$(CODILEIDO$, 3) = ABUS$ Then
        List4.AddItem Mid$(XX$, 77, 15) + Mid$(XX$, 5, 64)
      End If
    Next U&
    '
End Sub

Private Sub Text1_DblClick()
    '
    Command9_Click
    '
End Sub
