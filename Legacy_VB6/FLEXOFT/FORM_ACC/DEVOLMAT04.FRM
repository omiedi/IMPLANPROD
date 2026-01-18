VERSION 5.00
Begin VB.Form DEVOLMAT04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolución de Materiales"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture15 
      Height          =   2100
      Left            =   5985
      ScaleHeight     =   2040
      ScaleWidth      =   1035
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   2200
         Left            =   0
         Picture         =   "DEVOLMAT04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
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
      Left            =   2400
      TabIndex        =   0
      Top             =   210
      Width           =   1755
   End
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
      Left            =   4515
      TabIndex        =   1
      Top             =   1455
      Width           =   1230
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
      Left            =   4120
      TabIndex        =   14
      Top             =   210
      Width           =   175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004040&
      Height          =   6000
      Left            =   5985
      ScaleHeight     =   5940
      ScaleWidth      =   2160
      TabIndex        =   11
      Top             =   0
      Width           =   2220
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   1400
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command29 
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
         Height          =   650
         Left            =   120
         Picture         =   "DEVOLMAT04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Aprueba y graba "
         Top             =   720
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   105
         Picture         =   "DEVOLMAT04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   0
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "DEVOLMAT04.frx":27EC
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1410
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
      Left            =   2850
      TabIndex        =   4
      Top             =   1455
      Width           =   175
   End
   Begin VB.Label Label8 
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
      TabIndex        =   10
      Top             =   210
      Width           =   495
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   255
      Width           =   1020
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote o Partida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   8
      Top             =   1350
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
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
      Left            =   -840
      TabIndex        =   7
      Top             =   405
      Width           =   1965
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
      Left            =   420
      TabIndex        =   6
      Top             =   630
      Width           =   5325
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
      Left            =   1170
      TabIndex        =   5
      Top             =   1455
      Width           =   1860
   End
   Begin VB.Label Label43 
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
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   1500
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
      Left            =   315
      TabIndex        =   2
      Top             =   255
      Width           =   1965
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "DEVOLMAT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call Command29_Click
End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
      Text1 = VALACT1$("MASTER")
    End If
End Sub

Private Sub Command29_Click()
    '
    Command29.Enabled = False
    CI% = CODINT%(Text1)
    If CI% < 1 Then
       Call MENSERR(24, "Código Inexistente o no Válido")
       GoTo 99
    End If
    '
    If TRIM$(Label28) = "" Then
      Call MENSERR(24, "No se Ha Elegido el Lote")
      GoTo 99
    End If
    '
    If XVALO(Text4) < 0.05 Then
      Call MENSERR(24, "No se Ha Ingresado Cantidad")
      GoTo 99
    End If
    '
    SALOT$ = FORMATNUM$(SALOTE(CI%, TRIM$(Label28)), "F10.1")
    '
    If XVALO(Text4) > XVALO(SALOT$) + 0.05 Then
       Call MENSERR(24, "Saldo Insuficiente en el Lote Elegido.")
       GoTo 99
    End If
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE IDENLOTE= '" & TRIM$(Label28) & "'"
    Set LOTETemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LOTETemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
           On Error GoTo 0
           DEP% = !DepoENT
           PRO% = !NUMPROV
           REMIRECH$ = !NUMREM_PROV
           'NUBORE$ = !NUMEBORE
           '
       End If
       On Error GoTo 0
    End With
    PROV% = -1 * PRO%
    FF% = HOY(HOS$)
    LOTE$ = TRIM$(Label28)
    If TRIM$(Control$("", "REMIRECH")) <> "" Then
      CODRECHA$ = TRIM$(Control$("", "CODRRECH"))
      If CODRECHA$ = "" Then CODRECHA$ = "RD"
    End If
    NUBORE$ = XVALO(Mid$(LOTE$, 4)) ' SE QUEDA SOLO CON EL NUMERO
    DOSEC$ = "RT." + TRIM$(REMIRECH$)
    REMIRE$ = CODRECHA$ + "." + NUBORE$
    CANTI = XVALO(Text4) * (-1) ' cantidad
    REFE$ = "RECHAZO - Rto." + TRIM$(REMIRECH$)
    If DEP% = 99 Then DEP% = DEPOPRE%(CI%)
    Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRE$, REMIRE$, DOSELA$, DOSEC$, REFE$, 1, VUNI#, MONEII%, PROV%, DEP%, CANTI)
    Call ACSALOTE(CI%, LOTE$)
    '
    Call SETULTREG("!BOLREDET", 0) 'ESCRIBE EN EL BOLREDET
    YY$ = REGBLAN$("BOLREDET")
    Call REPLA(YY$, MKI$(CI%), 83, 2)
    Call REPLA(YY$, MKD$(CANTI * -1), 171, 8) ' LO VUELVE A POSITIVO
    Call GRAREG("!BOLREDET", YY$, 1)
    Call SETULTREG&("!BOLREDET", 1)
    Screen.MousePointer = 1
    '
    Call REMIRECHA(PROV%, DEP%, REMIRE$)
    '
    Call COMUNI("Devolución Procesada")
    Call limpiar
    Command29.Enabled = True
    Hide
    Exit Sub
    '
99  Screen.MousePointer = 1
    Command29.Enabled = True
End Sub

Private Sub Command5_Click()
    CI% = CODINT%(Text1)
    If CI% > 0 Then
      Call SELECHO("!LILOTSE")
      VDEVU$ = VALACT1$("!LILOTSE")
      If VDEVU$ <> "" Then
        Label28 = VDEVU$
      End If
    End If
End Sub


Private Sub Label46_Click()

End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture15.Visible = True
      Me.BackColor = &HFCD7B6
      MARCOBARRA.Top = 1800
      Set MARCOBARRA.Container = Me.Picture15
    End If

End Sub

Private Sub Label28_DblClick()
   Call Command5_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text1_Change()
   CI% = CODINT%(Text1.TEXT)
   Label2.Caption = DESCRIT0$(CI%)
   Label8.Caption = UNIMED$(CI%)
   Label28 = ""
   If CI% > 0 Then
      Call ARMALILOT(CI%)
      Screen.MousePointer = 1
   End If

End Sub

Private Sub Text1_DblClick()
   Call Command2_Click
End Sub

Private Sub Text4_LostFocus()
  Text4 = FORMATNUM(Text4, "F9.1")
End Sub

Sub limpiar()
    Text1 = ""
    Label28 = ""
    Text4 = ""
End Sub
