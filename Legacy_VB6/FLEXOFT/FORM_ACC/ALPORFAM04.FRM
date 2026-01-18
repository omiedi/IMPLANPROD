VERSION 5.00
Begin VB.Form ALPORFAM04 
   BackColor       =   &H00FCD7B6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Código de Artículo por Familia"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Height          =   4500
      Left            =   8160
      ScaleHeight     =   4440
      ScaleWidth      =   1020
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1080
      Begin VB.CommandButton Command5 
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
         Height          =   690
         Left            =   240
         Picture         =   "ALPORFAM04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Guardar "
         Top             =   3600
         Width           =   640
      End
      Begin VB.Image Image3 
         Height          =   4600
         Left            =   0
         Picture         =   "ALPORFAM04.frx":030A
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4515
      Left            =   8160
      ScaleHeight     =   4455
      ScaleWidth      =   1020
      TabIndex        =   16
      Top             =   0
      Width           =   1080
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
         Height          =   690
         Left            =   120
         Picture         =   "ALPORFAM04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Alta del Nuevo Código"
         Top             =   790
         Width           =   640
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
         Height          =   690
         Left            =   120
         Picture         =   "ALPORFAM04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   100
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   17
         Top             =   3600
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -240
         Picture         =   "ALPORFAM04.frx":2AF6
         Top             =   3950
         Width           =   1515
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
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   1050
      Width           =   6210
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCD7B6&
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
      Left            =   210
      TabIndex        =   10
      Top             =   1575
      Width           =   7680
      Begin VB.PictureBox Picture1 
         Height          =   2220
         Left            =   105
         ScaleHeight     =   2160
         ScaleWidth      =   7365
         TabIndex        =   13
         Top             =   450
         Width           =   7430
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
            ItemData        =   "ALPORFAM04.frx":4A18
            Left            =   -30
            List            =   "ALPORFAM04.frx":4A1A
            TabIndex        =   15
            Top             =   315
            Width           =   7440
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción"
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
            TabIndex        =   14
            Top             =   105
            Width           =   7425
         End
      End
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
      Height          =   300
      Left            =   3255
      Picture         =   "ALPORFAM04.frx":4A1C
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
      Height          =   300
      Left            =   2865
      MaxLength       =   15
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Denominación:"
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
      Left            =   -840
      TabIndex        =   12
      Top             =   1140
      Width           =   2430
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
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
      Left            =   3885
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
      Height          =   300
      Left            =   4515
      TabIndex        =   8
      Top             =   585
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
      Height          =   300
      Left            =   4515
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
      Height          =   300
      Left            =   5040
      TabIndex        =   6
      Top             =   585
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "U.Med:"
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
      Left            =   3360
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
      Height          =   300
      Left            =   5040
      TabIndex        =   4
      Top             =   210
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Familia:"
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
      Left            =   315
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
      Height          =   300
      Left            =   210
      TabIndex        =   2
      Top             =   585
      Width           =   3225
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MnuGuardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnusalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "ALPORFAM04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOK_Click()
    '
    BOTOK.Enabled = False
    '
    '////////**********///////////////
    'HAY QUE HACER ESTE PROC. CON LAS BASES NUEVAS.
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
    Call ABREMASTER
    SQLX$ = "SELECT * FROM MASTER "
    Set ARTLITMP = Articulos.OpenRecordset(SQLX$, 4)
    With ARTLITMP
        On Error Resume Next
        .MoveFirst
        If Err Then
           On Error GoTo 0
           Exit Sub
        End If
        On Error GoTo 0
        '
        Do Until .EOF = True
          On Error Resume Next
          CODILEIDO$ = UCase$(TRIM$(!CODIGO)) 'Mid$(XX$, 77, 15)))
          If Err Then
            On Error GoTo 0
            GoTo 19
          End If
          On Error GoTo 0
          If Left$(CODILEIDO$, 3) = ABUS$ Then
            NUITN& = XVALO(Mid$(CODILEIDO$, 4))
            If NUITN& >= PRONULI& Then PRONULI& = 1 + NUITN&
          End If
19        .MoveNext
        Loop
    End With
    '
    PRONUX$ = TRIM$(Str$(PRONULI&))
    While Len(PRONUX$) < 5
      PRONUX$ = "0" + PRONUX$
    Wend
    CODINUE$ = ABUS$ + PRONUX$
    '
    DESCRINUE$ = Text2
    TIMA% = XVALO(Label7)
    UNID$ = TRIM$(Label6)
    
    Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
    '/////////***********/////////////////
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


Private Sub Command5_Click()
  Call Command3_Click
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

Private Sub Form_Load()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Picture3.Visible = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Cancel = 0 Then
       Cancel = 1
       Call Command3_Click
    End If
End Sub

Private Sub Label6_Change()
    '
    X$ = ECOARCH("UMEDIDA", TRIM$(Label6))
    Label3 = TRIM$(X$)
    '
End Sub

Private Sub Label7_Change()
    '
    X$ = ECOARCH("TATIMAT", Chr(XVALO(Label7)))
    Label5 = TRIM$(X$)
    '
End Sub



Private Sub MnuGuardar_Click()
    Call BOTOK_Click
End Sub

Private Sub mnusalir_Click()
   Call Command3_Click
End Sub

Private Sub text1_lostfocus()
    '
    XJ$ = ECOARCH("TAFAMIL", TRIM$(Text1))
    If XJ$ = "" Then
        Text1 = ""
    End If
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
    Call ABREMASTER
    SQLX$ = "SELECT * FROM MASTER "
    Set ARTLITMP = Articulos.OpenRecordset(SQLX$, 4)
    With ARTLITMP
        On Error Resume Next
        .MoveFirst
        If Err Then
           On Error GoTo 0
           Exit Sub
        End If
        On Error GoTo 0
        '
        Do Until .EOF = True
          On Error Resume Next
          CODILEIDO$ = UCase$(TRIM$(!CODIGO)) 'Mid$(XX$, 77, 15)))
          If Err Then
            On Error GoTo 0
            GoTo 19
          End If
          On Error GoTo 0
          If Left$(CODILEIDO$, 3) = ABUS$ Then
             List4.AddItem TRIM$(!CODIGO) + Chr(9) + TRIM$(!DESCRIPCION)
          End If
19        .MoveNext
        Loop
    End With
    '
End Sub

Private Sub Text1_DblClick()
    '
    Command9_Click
    '
End Sub
