VERSION 5.00
Begin VB.Form INFOPIE 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Por Pieza y Operación"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3675
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "O.K."
      Top             =   2310
      Width           =   1485
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
      Height          =   285
      Left            =   5595
      TabIndex        =   15
      Top             =   1680
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
      Left            =   2205
      TabIndex        =   14
      Top             =   1680
      Width           =   175
   End
   Begin VB.TextBox Text3 
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
      Left            =   1155
      TabIndex        =   1
      Top             =   1155
      Width           =   630
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
      Left            =   1785
      TabIndex        =   11
      Top             =   1155
      Width           =   175
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
      Left            =   4530
      TabIndex        =   3
      Text            =   " "
      Top             =   1680
      Width           =   1065
   End
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
      Height          =   300
      Left            =   2835
      TabIndex        =   6
      Top             =   105
      Width           =   175
   End
   Begin VB.TextBox Text10 
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
      Left            =   1155
      TabIndex        =   0
      Top             =   105
      Width           =   1710
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1365
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "O.K."
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   2
      Text            =   " "
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label2 
      Height          =   540
      Left            =   2625
      TabIndex        =   16
      Top             =   1995
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label7 
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
      Left            =   2100
      TabIndex        =   13
      Top             =   1155
      Width           =   3675
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
      Left            =   -1260
      TabIndex        =   12
      Top             =   1230
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta: "
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
      Left            =   2730
      TabIndex        =   10
      Top             =   1770
      Width           =   1755
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
      Left            =   -270
      TabIndex        =   9
      Top             =   165
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1155
      TabIndex        =   8
      Top             =   525
      Width           =   4620
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde: "
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
      Left            =   -630
      TabIndex        =   7
      Top             =   1770
      Width           =   1755
   End
End
Attribute VB_Name = "INFOPIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
    Command2.Enabled = False
    'VALIDACIONES
    Label2.Caption = "OK"
    
    HOII% = HOY(HOS$)
    DES1% = FECHANUM(Text4)
    HAS1% = FECHANUM(Text1)
    If DES1% < 33 Or HAS1% > HOII% Or HAS1% < DES1% Then
       Call MENSERR(24, "Rango de Fechas no Válido")
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
19  NOPSEL% = XVALO(Text3.TEXT)
    If DESCRIOP$(CID1%, NOPSEL%) = "" Then
      Call MENSERR(24, "Operación no Válida")
      Text3.SetFocus
      GoTo 99
    End If
    
    '
    CANOPERS% = 0
    RENDIMA = XVALO(CONTROL("PADIAMOD", "RENDIMAX"))
    If RENDIMA < 1 Then RENDIMA = 99999
    '
    Call BLOQARCH("INFOPEFA")
    JJ& = 0
    '
    Screen.MousePointer = 11
    FEC1 = CDbl(CVDAT(DES1%))
    FEC2 = CDbl(CVDAT(HAS1%))
    '
    Call ABREORFAB
    SQLX$ = " SELECT * FROM InfoFab "
    SQLX$ = SQLX$ + " WHERE  INT(CDBL(FeHorLib))  >=  " & FEC1 & " "
    SQLX$ = SQLX$ + " AND  INT(CDBL(FeHorLib))  <=  " & FEC2 & " "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CID1% & " "
    SQLX$ = SQLX$ + "AND NumeOper = " & NOPSEL% & " "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set PROXFECTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PROXFECTMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No Existen Informes de Producción\para el Criterio Seleccionado")
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 99
       End If
       On Error GoTo 0
       '
       JJ& = 0
       XX$ = REGBLAN$("INFOPEFA")
       Do While .EOF = False
         '
         ALTOP% = 0
         On Error Resume Next
         ALTOP% = !AltOper
         On Error GoTo 0
         '
         LEGAJOP% = CVI(!LegaOps)
         Call REPLA(XX$, MKI$(LEGAJOP%), 1, 2)
         Call REPLA(XX$, !IdSubOr, 3, 12)
         Call REPLA(XX$, MKI$(!NumeOper), 15, 2)
            DOPE$ = TRIM$(Str$(!NumeOper))
            If ALTOP% > 0 Then DOPE$ = DOPE$ + "/" + TRIM$(Str$(ALTOP%))
            DOPE$ = AJUSTI$(DOPE$, 7) + !DescOper
         Call REPLA(XX$, DOPE$, 17, 42)
         On Error Resume Next
         Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 59, 2)
         Call REPLA(XX$, MKI$(!Hora_Des), 61, 2)
         Call REPLA(XX$, MKI$(!Hora_Has), 63, 2)
         On Error GoTo 0
         Call REPLA(XX$, !Maquina, 65, 6)
           HHORA$ = Format(!FeHorLib, "hh:mm:ss")
           If HHORA$ = "00:00:00" Or TRIM$(HHORA$) = "" Then HHORA$ = HORATEXNOR$(!Hora_Has)
           FFECH$ = Format(!FeHorLib, "dd/mm/yy")
         Call REPLA(XX$, MKI$(CVINT(!FeHorLib)), 71, 2)
         Call REPLA(XX$, MKI$(HORANUM(HHORA$)), 73, 2)
         Call REPLA(XX$, !IdSubOr, 75, 12)
         Call REPLA(XX$, MKI$(!Cod_Inte), 87, 2)
         Call REPLA(XX$, MKI$(!NumeOper), 89, 2)
         Call REPLA(XX$, DOPE$, 91, 48)
         'Se suma Cant. Real y Cant. No Conf
         NoConf# = XVALO(!CaNoConf)
         Real# = XVALO(!CantReal)
         TotCant# = NoConf# + Real#
         Call REPLA(XX$, MKD$(TotCant#), 139, 8)
         'Call REPLA(XX$, MKD$(!CantReal), 139, 8)
         Call REPLA(XX$, MKS$(60 * !TiemEjec), 147, 4)
         TSTAN = 60 * TIESTAN(!Cod_Inte, !NumeOper, !CantReal, !NuOrInf, ALTOP%)
         Call REPLA(XX$, MKS$(TSTAN), 155, 4)
         '
         ' ENJUAGUE ESCORIAL
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "ESCORIAL") > 0 Then
            'TSTANBA = TIESTAN(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
            TSTANBA = TIEMAQUI(!Cod_Inte, !NumeOper, 1, 1, ALTOP%)
            TSTAN = 60 * TIEMAQUI(!Cod_Inte, !NumeOper, TotCant#, !NuOrInf, ALTOP%)
            CADEN = 0
            If TSTANBA > 0 Then CADEN = 60 / TSTANBA
            Call REPLA(XX$, MKS$(CADEN), 155, 4)
         End If
           '
           If !TiemEjec <= 0 Then
               RENDIM = 0
             Else
               RENDIM = (100 * TSTAN) / (!TiemEjec * 60)
           End If
         Call REPLA(XX$, MKS$(RENDIM), 163, 4)
         Call REPLA(XX$, MKI$(LEGAJOP%), 167, 2)
         '
         JJ& = JJ& + 1
         Call GRAREG("INFOPEFA", XX$, JJ&)
         '
29      .MoveNext
       Loop
    End With
    '
    Call SETULTREG("INFOPEFA", JJ&)
    Call CIERRARCH("INFOPEFA")
    '
    FEX = FF%
    Call HEADERS("INFOPEFA", "")
    'AGREGADO
    Call HEADERS("INFOPEFA", "Corresponde a: " + TRIM$(CODEXT$(CID1%)) + " " + TRIM$(DESCRIT0$(CID1%)) + " - " + TRIM$(Str$(NOPSEL%)) + " " + TRIM$(DESCRIOP$(CID1%, NOPSEL%)))
    Call FILESORT("INFOPEFA", "", 6, 8, "ASC")
    Call FILESORT("INFOPEFA", "", 1, 2, "ASC")
    Call CONECRUN("*INFOPIE", "Informes de Producción por Pieza")
    Call CIERRARCH("INFOPEFA")
    Screen.MousePointer = 1
    '
99  Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text1.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command4_Click()
Unload Me
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
  Call SELECHO("!OPERINF/Operaciones de Codigo: " + Text10) '+ Label4.Caption)
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
  Text4.TEXT = FECHATEX(COMESACT)
  Text1.TEXT = FECHATEX$(HOII%)
  Call INFOFA04.GENRESDAT
  Call BLANARCH("!OPERINF")
  Call SETULTREG("!OPERINF", 0)
  Call FRESHECHO("!OPERINF")
  Screen.MousePointer = 1
  '
End Sub
'
'Private Sub Text1_Change()
'  '
'  Label5.Caption = ECOARCH$("PADRON", MKI$(XVALO(Text1.TEXT)))
'  '
'End Sub
'Private Sub Text1_DblClick()
'  '
'  ' Call Command4_Click
'  '
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Command2.SetFocus
  End If
End Sub

Private Sub Text10_Change()
   Label3 = DESCRIT0$(CODINT%(Text10))
   '
   CIXXI% = CODINT%(Text10)
   Text3 = ""
   If CIXXI% > 0 Then
     Call GENECOPER(CIXXI%)
   End If
   '
End Sub

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
           If IsNull(!AltOper) = True Or !AltOper < 1 Then
             SALDO# = 0
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
           End If
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
'
Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text4.SetFocus
   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text1.SetFocus
   End If
End Sub
  
  '
