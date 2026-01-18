VERSION 5.00
Begin VB.Form CONPROD04 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parte Diario  M.O.D."
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   10110
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
      Left            =   5145
      TabIndex        =   3
      Top             =   315
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
      Left            =   3150
      TabIndex        =   2
      Top             =   315
      Width           =   2010
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
      Left            =   2100
      TabIndex        =   5
      Top             =   1155
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
      Left            =   1470
      TabIndex        =   4
      Top             =   1155
      Width           =   630
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
      Left            =   3495
      TabIndex        =   1
      Top             =   3360
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
      Left            =   2430
      TabIndex        =   0
      Text            =   " "
      Top             =   3360
      Width           =   1065
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00004080&
      Height          =   8460
      Left            =   9225
      ScaleHeight     =   8400
      ScaleWidth      =   1005
      TabIndex        =   9
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CONPROD04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Listado de Informe de Produccion por Fecha"
         Top             =   1680
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Setup"
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
         Picture         =   "CONPROD04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   3885
         Width           =   650
      End
      Begin VB.CommandButton Command26 
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
         Picture         =   "CONPROD04.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Corrección Partes de Fabricación"
         Top             =   2310
         Width           =   650
      End
      Begin VB.CommandButton Command19 
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
         Picture         =   "CONPROD04.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "A-B-M Padrón de Operarios"
         Top             =   3245
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   4620
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
         Picture         =   "CONPROD04.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "CONPROD04.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   845
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CONPROD04.frx":0FF4
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5025
         Width           =   1515
      End
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
      Left            =   -945
      TabIndex        =   16
      Top             =   1215
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
      Left            =   2415
      TabIndex        =   15
      Top             =   1155
      Width           =   2955
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Artículo:"
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
      TabIndex        =   14
      Top             =   390
      Width           =   2910
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
      Left            =   210
      TabIndex        =   13
      Top             =   735
      Width           =   5130
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   1050
      TabIndex        =   12
      Top             =   3420
      Width           =   1335
   End
End
Attribute VB_Name = "CONPROD04"
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

Private Sub Command14_Click()
   Command14.Enabled = False
   STPPMOD04.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Call INFOFA04.Command26_Click
   Command16.Enabled = True
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   'opxx% = ALTERNA%("Padron de Operarios\Lineas de Producción\Integrantes de Línea")
   'If opxx% = 1 Then
        Call FINAL("*PADRON")
        AppActivate Caption
        SendKeys Chr$(1)
   '   ElseIf opxx% = 2 Then
   '     Call MENSERR(24, "No Disponoble")
   '   ElseIf opxx% = 3 Then
   '     Call MENSERR(24, "No Disponible")
   'End If
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   'VALIDACIONES
   HOII% = HOY(HOS$)
   If FECHANUM(Text4) < 33 Or FECHANUM(Text4) > HOII% Then
      Call MENSERR(24, "Fecha no Válida")
      Text4.SetFocus
      GoTo 99
   End If
   '
   LEGAI% = Val(Text1)
   If ECOARCH$("PADRON", MKI$(LEGAI%)) = "" Then
     Call MENSERR(24, "Operario Inexistente o no Válido")
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
19 If DESCRIOP$(CID1%, XVALO(Text3.TEXT)) = "" Then
     Call MENSERR(24, "Operación no Válida")
     Text3.SetFocus
     GoTo 99
   End If
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
   TDES% = HORANUM%(Text6)
   THAS% = HORANUM%(Text7)
   If THAS% < TDES% Or (THAS% - TDES%) < 1 Then
     Call MENSERR(24, "Horarios de Comienzo y Fin no Válidos")
     Text6.SetFocus
     GoTo 99
   End If
   '
   'GRABAR INFORME
   NROOF$ = Label4
   CIXXI% = CODINT%(Text10)
   NOPE% = XVALO(Text3)
   DOPE$ = Label7
   CAINF# = XVALO(Text8)
   MAQUI$ = Text5
   TIEINF = 60 * DIFHORA(TDES%, THAS%)
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
   SALDO# = XVALO(Label18)
   STATUOP% = 3 ' Cumplida
   MALIB$ = "S"
   If CAINF# < SALDO# - 0.05 Then
     STATUOP% = 1   ' AUN PENDIENTE
     MALIB$ = " "
 '      TTXX$ = "Queda un Saldo Pendiente\para esta Operación\\Mantener Pendiente\Cancelar Saldo"
 '      If COMALTER%(TTXX$) = 2 Then
 '        STATUOP% = 8 ' CANCELADO
 '      End If
   End If
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
      !DescMaq = ECOARCH$("PADMAQ", MAQUI$)
      !LegaOps = MKI$(Int(XVALO(Text1)))
         FEX = FECHANUM(Text4)
         FEHOLI = CDate(FECHATEX$(FEX) + " " + HORATEXNOR$(Int(THAS%)))
      !FeHorLib = FEHOLI
      !Hora_Des = TDES%
      !Hora_Has = THAS%
      !MaqLiber = MALIB$
      !TiemEjec = TIEINF
         MODO% = CAINFANT%
      !TiemStan = TIESTAN(CIXXI%, NOPE%, CAINF#, MODO%)   ' 1 o mas es para que no considere preparacion
           'COMENTE LA SIG. LINEA XQ LE FALTABA HACER REFERENCIA
           ' A LOS CAMPOS DEL REGISTRO. "!"
           'RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * TiemEjec / TiemStan
         RENDI = 999.9: If !TiemStan > 0 Then RENDI = 100 * !TiemEjec / !TiemStan
         If RENDI > 999.9 Then RENDI = 999.9
      !RENDIM = RENDI
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
   Text10 = ""
   Text3 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   '
   Text10.SetFocus
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   MODIPAMOD04.Show 1
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
   Call SELECHO("!ECORDEP/Ordenes de Trabajo en Curso")
   NROOF$ = AJUSTI$(VALACT1$("!ECORDEP"), 12)
   If TRIM$(NROOF$) = "" Then
     Label4 = ""
     Exit Sub
   End If
   '
   Label4 = NROOF$
   '
End Sub

Private Sub Command4_Click()
  '
  Call SELECHO("PADRON")
  VDEV$ = VALACT1$("PADRON")
  If VDEV$ <> "" Then
    Text1.TEXT = VDEV$
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
  Call SELECHO("!OPERINF/Operaciones de Codigo: " + Label4.Caption)
  CODOP1$ = VALACT1$("!OPERINF")
  If CODOP1$ <> "" Then
    Y$ = FILTERGETRECORD$("!OPERINF", 1, MKI$(XVALO(CODOP1$)))
    CODOP1$ = Mid$(Y$, 3, 58)
    I% = InStr(1, CODOP1$, "-")
    CODOP2$ = TRIM$(Mid$(Y$, 1, I%))
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

Sub Form_Load()
  '
  Screen.MousePointer = 11
  Call CENTRAFORM(Me)
  HOII% = HOY(HOS$)
  Text4.TEXT = FECHATEX$(HOII%)
  Call INFOFA04.GENRESDAT
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  Screen.MousePointer = 1
  '
End Sub

Private Sub Label4_Change()
    '
    Call GENECODOPEOF(Label4)
    '
End Sub

Private Sub Text1_Change()
  '
  Label5.Caption = ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT)))
  '
End Sub
Private Sub Text1_DblClick()
  '
  Call Command4_Click
  '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Text10.SetFocus
End If
End Sub

Private Sub Text10_Change()
   Label3 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Label4 = "": CANOR% = 0
   Text3 = "": Text5 = ""
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
      Call SETULTREG("!ECORDEP", JJ&)
      Call CIERRARCH("!ECORDEP")
      Call FRESHECHO("!ECORDEP")
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub TEXT2_DBLCLICK()
  '
  Call Command3_Click
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
  Text3.SetFocus
End If


End Sub

Private Sub Text3_Change()
  '
  CID1% = CODINT%(Text10)
  Label7 = DESCRIOP$(CID1%, XVALO(Text3.TEXT))
  Text5.TEXT = MAQDEFIOP$(Label4, XVALO(Text3.TEXT))
  Text8.TEXT = CAPENOP$(Label4, XVALO(Text3.TEXT), CANPROGRA)
  Label18 = Text8
  Label19 = TRIM$(FORMATNUM$(CANPROGRA, "F12.1"))
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
     Set FABRTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
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
          CAPROD# = !CantProg      'Cantidad Programada
          CAREA# = !CantReal       'Cantidad realizada
          SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
          '
          If SALDO# > 0.05 Then
            NOPE% = !NumeOper      'numero de operación
            DOPE$ = !DescOper      'Descripcion de la operacion
            MAQUI$ = !Maquina      'Máquina
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
          End If
       .MoveNext
       Loop
     End With
        
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   Text5.SetFocus
 End If
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
Private Sub Text5_dblclick()
  '
  Call Command7_Click
  '
End Sub
Private Sub text5_change()
  '
  Label10.Caption = ECOARCH$("PADMAQ", Text5.TEXT)
  '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   Text7.SetFocus
  End If
End Sub

Private Sub Text6_lostfocus()
  '
   Static MODORA%
   If MODORA% < 1 Then
      MODORA% = 1
      If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
         MODORA% = 2
      End If
   End If
   '
  If XVALO(Text6.TEXT) > 0 Then
    Text6.TEXT = REPLACAR$(Text6.TEXT, ".", ":")
    If InStr(1, Text6.TEXT, ".") = 0 Then
      If InStr(1, Text6.TEXT, ":") = 0 Then
        Text6.TEXT = Text6.TEXT + ":0"
      End If
    End If
    '
    If MODORA% = 2 Then
       TTXX1$ = FORMATNUM$(REPLACAR$(Text6.TEXT, ":", "."), "F7.2")
       Text6 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
       Exit Sub
    End If
    '
    On Error Resume Next
      D1 = CDate(Text6.TEXT)
    On Error GoTo 0
    Text6.TEXT = Format$(D1, "hh:mm")
  End If
  '
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    Text8.SetFocus
   End If
End Sub

Private Sub Text7_lostfocus()
  '
  Static MODORA%
  If MODORA% < 1 Then
     MODORA% = 1
     If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
        MODORA% = 2
     End If
  End If
  '
  If XVALO(Text7.TEXT) > 0 Then
    Text7.TEXT = REPLACAR$(Text7.TEXT, ".", ":")
    If InStr(1, Text7.TEXT, ".") = 0 Then
      If InStr(1, Text7.TEXT, ":") = 0 Then
        Text7.TEXT = Text7.TEXT + ":0"
      End If
    End If
    '
    If MODORA% = 2 Then
       TTXX1$ = FORMATNUM$(REPLACAR$(Text7.TEXT, ":", "."), "F7.2")
       Text7 = TRIM$(REPLACAR$(TTXX1$, ".", ","))
       Exit Sub
    End If
    '
    On Error Resume Next
      D1 = CDate(Text7.TEXT)
    On Error GoTo 0
    Text7.TEXT = Format$(D1, "hh:mm")
    '
  End If
  '
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
