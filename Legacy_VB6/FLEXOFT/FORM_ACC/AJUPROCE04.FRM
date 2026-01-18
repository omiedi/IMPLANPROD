VERSION 5.00
Begin VB.Form AJUPROCE04 
   BackColor       =   &H00F0F0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuste Saldos en Proceso por O/F"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3675
      TabIndex        =   3
      Top             =   840
      Width           =   175
   End
   Begin VB.TextBox Text10 
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
      TabIndex        =   2
      Top             =   840
      Width           =   2010
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "AJUPROCE04.frx":0000
      Top             =   2835
      Width           =   5685
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
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
      Left            =   4095
      TabIndex        =   5
      Text            =   " "
      Top             =   2310
      Width           =   1170
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
      Left            =   2310
      TabIndex        =   4
      Top             =   1785
      Width           =   175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7140
      TabIndex        =   13
      Top             =   840
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
      Left            =   7170
      TabIndex        =   1
      Top             =   315
      Width           =   175
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
      Left            =   6105
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   1065
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8460
      Left            =   7560
      ScaleHeight     =   8400
      ScaleWidth      =   1005
      TabIndex        =   9
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   105
         Picture         =   "AJUPROCE04.frx":0002
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "O.K."
         Top             =   2910
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   1890
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
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
         Height          =   640
         Left            =   105
         Picture         =   "AJUPROCE04.frx":030C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "AJUPROCE04.frx":0456
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "AJUPROCE04.frx":0760
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2205
         Width           =   1515
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ajuste:"
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
      Left            =   4725
      TabIndex        =   27
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   6130
      TabIndex        =   26
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1680
      TabIndex        =   25
      Top             =   1785
      Width           =   645
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1680
      TabIndex        =   23
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Teorica:"
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
      Left            =   315
      TabIndex        =   22
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/F Nro:"
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
      Left            =   4095
      TabIndex        =   21
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones:"
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
      Left            =   210
      TabIndex        =   20
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant Ajustada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2870
      TabIndex        =   19
      Top             =   2205
      Width           =   1125
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación:"
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
      Left            =   -735
      TabIndex        =   18
      Top             =   1845
      Width           =   2325
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   2625
      TabIndex        =   17
      Top             =   1785
      Width           =   4695
   End
   Begin VB.Label Label6 
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
      Left            =   315
      TabIndex        =   16
      Top             =   915
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   5475
      TabIndex        =   15
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      TabIndex        =   14
      Top             =   1260
      Width           =   5655
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   4725
      TabIndex        =   12
      Top             =   375
      Width           =   1335
   End
End
Attribute VB_Name = "AJUPROCE04"
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
   'VALIDACIONES
   If Abs(XVALO(Label5)) < 0.05 Then
     Call MENSERR(24, "Sin Cantidad Informada para Ajuste")
     GoTo 99
   End If
   '
   HOII% = HOY(HOS$)
   If FECHANUM(Text4) < 33 Or FECHANUM(Text4) > HOII% Then
      Call MENSERR(24, "Fecha no Válida")
      Text4.SetFocus
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
   If TRIM$(Label4) = "" Then
     Call MENSERR(24, "Falta Número de O/Fabricación")
     Command3.SetFocus
     GoTo 99
   End If
   '
   If TRIM$(Label4) <> "" Then
     Call ABREORFAB
     OrdeFabr.FindFirst "IdSubOr = '" & TRIM$(Label4) & "' "
     If OrdeFabr.NoMatch = False Then
       If OrdeFabr!Cod_Inte = CID1% Then
         GoTo 19    ' VALIDA NUMERO DE o/f Y CODIGO
       End If
     End If
     Call MENSERR(24, "O/F no Corresponde a Código de Artículo")
     GoTo 99
   End If
   '
19 If XVALO(Label2) < 1 Then
      Call MENSERR(24, "Falta Número de Operación")
      GoTo 99
   End If
   '
   'GRABAR INFORME
   NROOF$ = Label4
   CIXXI% = CODINT%(Text10)
   NOPE% = XVALO(Label2)
   DOPE$ = Label7
   CAINF# = XVALO(Text8) - XVALO(Label18)
   MAQUI$ = "Ajuste"
   DEMAQ$ = TRIM$(USERNAME) + " - " + Format$(Now, "dd/mm/yy hh:mm")
   TIEINF = 0
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   CAINFANT% = 0
   Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         CAREAL# = CAREAL# + !CantReal
         CAINFANT% = CAINFANT% + 1
       .MoveNext
       Loop
     End If
   End With
   '
   Call ABREORFAB
   With InfoFab
      .AddNew
      !IdSubOr = NROOF$
      !DescriOF = ECOARCH$("ECORDEF", NROOF$)
      !NuOrInf = CAINFANT%
      !Cod_Inte = CIXXI%
      !Cod_Exte = CODEXT$(CIXXI%)
      !Descrip = DESCRIT$(CIXXI%)
      !NumeOper = NOPE%
      !DescOper = DOPE$
      !CanOrPro = XVALO(Label19)
      !CantReal = CAINF#
      !Maquina = MAQUI$
      !DescMaq = DEMAQ$
      !LegaOps = MKI$(0)
         FEX = FECHANUM(Text4)
         FEHOLI = CDate(FECHATEX$(FEX))
      !FeHorLib = FEHOLI
         If FEX = HOY(HOS$) Then
             TDES% = HORANUM(Time$)
           Else
             TDES% = 2400
         End If
      !Hora_Des = TDES%
      !Hora_Has = TDES%
      !MaqLiber = " "
      !TiemEjec = 0
         MODO% = CAINFANT%
      !TiemStan = 0
      !RENDIM = 0   ' RENDI
      !Observa = Text9
      .Update
   End With
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   '
   CAREAL# = 0: TIEINF = 0
   Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         CAREAL# = CAREAL# + !CantReal
         TIEINF = TIEINF + !TiemEjec
       .MoveNext
       Loop
     End If
   End With
   '
   With OperFab
     .FindFirst "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NOPE%
     If .NoMatch = False Then
       .Edit
       !CantReal = CAREAL#
       !CantPend = !CantProg - CAREAL#
         If !CantPend < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CantPend = 0
         CANTPE# = !CantPend
         MODO% = CAINFANT%
       !TiemEjec = TIEINF
       !TiemPend = TIEMAQUI(CIXXI%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
           If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
       !StatOper = STATUOP%
         .Update
     End If
   End With
   '
   RETEX10$ = Text10
   Text10 = ""
   Label2 = ""
   Label3 = ""
   Label4 = ""
   Label7 = ""
   Label18 = ""
   Text8 = ""
   Text9 = ""
   Text10 = RETEX10$
   '
   Text10.SetFocus
   '
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
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
     If TRIM$(NROOF$) = "" Then GoTo 90
   End If
   '
90 Label4 = NROOF$
   Screen.MousePointer = 1
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
    I% = InStr(1, CODOP1$, "-")
    CODOP2$ = TRIM$(Mid$(Y$, 1, I%))
    Label2 = CODOP2$
    Label7 = TRIM$(Mid$(CODOP1$, I% + 1, 46 - I%))
    Label18 = TRIM$(Mid$(CODOP1$, 47))
    Text8 = Label18
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

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Text4.TEXT = FECHATEX$(HOII%)
  'Call INFOFA04.GENRESDAT
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  Screen.MousePointer = 1
  '
End Sub

Private Sub Label18_Change()
   Call Text8_Change
End Sub

Private Sub Label4_Change()
    '
    Call GENECODOPEOF(Label4)
    '
End Sub

Private Sub Text10_Change()
   Label3 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label4 = "": CANOR% = 0
   Text3 = "": Text5 = ""
   JJ& = 0
   Call SETULTREG("!ECORDEP", 0)
   '
   If CIXXI% > 0 Then
      Call ABREORFAB
      SQLX$ = " SELECT * FROM ORDEFABR "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXXI% & " "
      SQLX$ = SQLX$ + "AND StatUORF < 8 "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + "AND CanApro < CanProy "                      ' no marca cumplidos"
      SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC "
      '
      Set ORFASEL = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      On Error Resume Next
      ORFASEL.MoveLast
      If Err Then
         GoTo 90
      End If
      Screen.MousePointer = 11
      CANOR% = ORFASEL.RecordCount
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
      '
90    Call SETULTREG("!ECORDEP", JJ&)
      Call CIERRARCH("!ECORDEP")
      Call FRESHECHO("!ECORDEP")
      Screen.MousePointer = 1
      '
      If JJ& < 1 Then
        Call MENSERR(24, "No hay Ordenes de Fabricación en Curso\para el Código Seleccionado")
      End If
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

Private Sub Text10_DblClick()
   Call Command8_Click
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Command3.SetFocus
   End If
End Sub
'
Sub GENECODOPEOF(NROOF$)
     '
     Dim NOPY1%(1024), DOPY1$(1024), CACU1(1024), CSAL1(1024)
     '
     Call BLANARCH("!OPERINF")
     Call SETULTREG("!OPERINF", 0)
     Call FRESHECHO("!OPERINF") 'nuevo
     '
     If TRIM$(NROOF$) = "" Then
       Exit Sub
     End If
     '
     FECHAJU% = FECHANUM(Text4)
     'ACA OBTENGO LOS REGISTROS DE OPERFAB, CUYO IdSubOr = NROOF$
     'Luego cambio las variable por los campos correspondientes
     '
     SQLX$ = " SELECT * FROM InfoFab "
     SQLX$ = SQLX$ + " WHERE  IdSubOr = '" & NROOF$ & "'"
     SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
     Set IFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     JJ& = 0: CANOPSTA% = 0: NOPEA% = 0
     With IFATEMP
       On Error Resume Next
       .MoveFirst '
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No hay Informes de Fabricación\para la Orden de Fabricación Elegida.")
         Exit Sub
       End If
       On Error GoTo 0
       '
       Do While .EOF = False
         '
         If IsNull(!ReTraba) Or !ReTraba < 1 Then
            If CVINT(!FeHorLib) <= FECHAJU% Then
              NOP2% = !NumeOper
              CAN2 = !CantReal
              DOPE$ = !DescOper
              For JI1% = 1 To CANOPSTA%
                If NOPY1%(JI1%) = NOP2% Then
                  CACU1(JI1%) = CACU1(JI1%) + CAN2
                  GoTo 19
                End If
              Next JI1%
            End If
            '
            If NOP2% <> NOPEA% Then
               JJ& = JJ& + 1
               NOPY1%(JJ&) = NOP2%
               DOPY1$(JJ&) = DOPE$
               CACU1(JJ&) = CAN2
               CSAL1(JJ&) = 0
               NOPEA% = NOPE%
               CANOPSTA% = JJ&
            End If
         End If
         '
19     .MoveNext
       Loop
     End With
     '
     JJ& = JJ& + 1
     NOPY1%(JJ&) = 9999
     DOPY1$(JJ&) = "Parcial O/F - Ingreso Stock"
     CACU1(JJ&) = 0
     CSAL1(JJ&) = 0
     CANOPSTA% = JJ&
     '
20   ' AGREGA ENTRADAS A STOCK POR CIERRE DE o/fABRIC
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM Movisto "
     SQLX$ = SQLX$ + " WHERE  CodiMovi = 'OF' "
     SQLX$ = SQLX$ + "AND DoPriLar = '" & NROOF$ & "'"
     Set Movisel = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With Movisel
        On Error Resume Next
        .MoveFirst '
        If Err Then
          On Error GoTo 0
          GoTo 30
        End If
        On Error GoTo 0
        '
        Do While .EOF = False
          CACU1(CANOPSTA%) = CACU1(CANOPSTA%) + !CantIngre
        .MoveNext
        '
        Loop
     End With
     '
30   JJ& = 0
     For JI1% = 1 To CANOPSTA% - 1
       CSAL1(JI1%) = CACU1(JI1%) - CACU1(JI1% + 1)
       YY$ = MKI$(JI1%) + Space$(58) + String(4, 0)
       NOPX$ = AJUSTI$(TRIM$(Str$(NOPY1%(JI1%))) + " - " + AJUSTI$(TRIM$(Left$(DOPY1$(JI1%), 48)), 48), 46)
       NOPX$ = NOPX$ + FORMATNUM$(CSAL1(JI1%), "F12.1")
       Call REPLA(YY$, NOPX$, 3, 58)
       JJ& = JJ& + 1
       Call GRAREG("!OPERINF", YY$, JJ&)
     Next JI1%
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
     Set SecopTemp = Articulos.OpenRecordset("SELECT * FROM SecOper WHERE CodIConj=" & CIJK% & " AND StatOper = 1 ORDER BY NumeOper ASC", dbOpenSnapshot)
     On Error Resume Next
     SecopTemp.MoveFirst
     If Err < 1 Then
       JJ& = 0
       YY$ = REGBLAN("OPERINF")
       With SecopTemp
         Do While Not .EOF
           CAPROD# = 0      'Cantidad Programada
           CAREA# = 0       'Cantidad realizada
           SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
           '
           NOPE% = !NumeOper      'numero de operación
           NO3$ = FORMATNUM(NOPE%, "F3.0")
           DOPE$ = !DescOper      'Descripcion de la operacion
           MAQUI$ = !CodMaq0      'Máquina
           '
           JJ& = JJ& + 1: JJI% = JJ&
           YY$ = MKI$(JJI%) + Space$(58) + String(4, 0)
           NOPX$ = AJUSTI$(TRIM$(Str$(NOPE%)) + " - " + AJUSTI$(TRIM$(Left$(DOPE$, 31)), 30) + " - " + MAQUI$, 46)
           NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
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
  Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    CAPROD# = !CantProg      'Cantidad Programada
    CAREA# = !CantReal       'Cantidad realizada
    SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
    CAPENOP$ = NUMSTRI$(SALDO#, 12, 1, 0)
    CANPROGRA = CAPROD#
  End With
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
  Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With FABRTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 20
    End If
    On Error GoTo 0
    MAQDEFIOP$ = !Maquina
  End With
  '
20 If MAQDEFIOP$ = "" Then
     CIXI% = CODINT%(Text10)
     MAQDEFIOP$ = MAQUISTAN$(CIXI%, NOPE1%)
   End If
  '
End Function

Private Sub Text4_DblClick()
   Call Command5_Click
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text10.SetFocus
   End If
End Sub

Private Sub Text8_Change()
  Label5 = FORMATNUM(XVALO(Text8) - XVALO(Label18), "F12.1")
End Sub

Private Sub Text8_GotFocus()
  Text8.SelStart = 0
  Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    Text9.SetFocus
   End If
End Sub

Private Sub Text8_LOSTFOCUS()
  '
  Text8.TEXT = NUMSTRI$(XVALO(Text8.TEXT), 12, 1, 0)
  '
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    Command2.SetFocus
   End If
End Sub

