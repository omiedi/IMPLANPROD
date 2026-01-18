VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DIARGEN07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario General"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3285
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   8600
      Left            =   2400
      ScaleHeight     =   8565
      ScaleWidth      =   1245
      TabIndex        =   22
      Top             =   -105
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         Height          =   195
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   27
         Top             =   3120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   28
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
         Height          =   690
         Left            =   120
         Picture         =   "DIARGEN07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command15 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "DIARGEN07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Genera Reporte Solicitado"
         Top             =   2160
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   600
      TabIndex        =   19
      Top             =   4800
      Width           =   2445
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
         Caption         =   "Cuentas c/Movim."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0F0&
         Caption         =   "Todas las Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1980
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H0095C5E1&
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Doble-Click para Editar Rango de Fechas"
      Top             =   120
      Width           =   2115
      Begin VB.CommandButton Command21 
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
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   175
      End
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
         Left            =   1560
         TabIndex        =   13
         Top             =   1455
         Width           =   175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   450
         TabIndex        =   18
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   450
         TabIndex        =   17
         Top             =   465
         Width           =   645
      End
      Begin VB.Label label5 
         Alignment       =   2  'Center
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
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Left            =   450
         TabIndex        =   15
         Top             =   1455
         Width           =   1110
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2150
      Left            =   0
      TabIndex        =   3
      Top             =   5775
      Width           =   4005
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
         Left            =   3675
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
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
         Left            =   1890
         TabIndex        =   6
         Top             =   360
         Width           =   1800
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
         Height          =   285
         Left            =   3675
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox Text2 
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
         Left            =   1890
         TabIndex        =   4
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label1 
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
         TabIndex        =   11
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   840
         TabIndex        =   10
         Top             =   420
         Width           =   960
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
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   3600
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   840
         TabIndex        =   8
         Top             =   1380
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0095C5E1&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   2115
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   2
         Top             =   680
         Width           =   1260
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "DIARGEN07.frx":0454
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "DIARGEN07.frx":0688
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "DIARGEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command15_Click()
   
   Command15.Enabled = False
   If FECHANUM(Label5) > FECHANUM(Label6) Or FECHANUM(Label5) < COMEJ% Or FECHANUM(Label6) > FINEJ% Then
     Call MENSERR(24, "Rango de Fechas Imposible")
     GoTo 99
   End If
   fexii% = FECHANUM(Label5)
   DESYY = FETEXCOR1$(fexii%)
   fexii% = FECHANUM(Label6)
   HASYY = FETEXCOR1$(fexii%)
   '
   FEINI = FETEXCOR1$(COMEJ%)
   FEFIN = FETEXCOR1$(FINEJ%)
   '
   Screen.MousePointer = 11
   'Call CIERRARCH("DIARPRIM")
   Call BLANARCH("DIARPRIM/NOERR")
   EMPREI% = CodiEmp%
   JJ& = 0
   REBLA$ = REGBLAN$("DIARPRIM")
   NICONFLIM% = NIVCONFI%
   '
   Screen.MousePointer = 11
   Call DIARPAN07.Form_Load
   '
   ''''Dim PASEASIEN As ADODB.Recordset
   ''''Set PASEASIEN = New ADODB.Recordset
   SQLX$ = "SELECT SUM(IDEBEPASE) AS SUDEBE, SUM(IHABEPASE) AS SUHABE FROM DetaAsien WITH(NOLOCK)"
   ' SQLX$ = SQLX$ + " WHERE SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " WHERE  Fechasien >= '" & FEINI & "'"
   SQLX$ = SQLX$ + " AND Fechasien < '" & DESYY & "'"
   SQLX$ = SQLX$ + " AND statpase >0  "
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
   ''''PASEASIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   Set PASEASIEN = READSET(SQLX$)
   '
   TRANSPOR# = 0: TRANDEBE# = 0: TRANHABE# = 0
   On Error Resume Next
   PASEASIEN.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     If Not IsNull(PASEASIEN!SUDEBE) Then TRANDEBE# = PASEASIEN!SUDEBE
     If Not IsNull(PASEASIEN!SUHABE) Then TRANHABE# = PASEASIEN!SUHABE
   End If
   On Error GoTo 0
   PASEASIEN.Close
   '
'   ' VALIDACION DE ASIENTOS CONTABLES POR ESTADO
'   Inconsistencia% = 0
'   SQLX$ = "SELECT * FROM EncabeAsien WITH(NOLOCK) WHERE "
'   SQLX$ = SQLX$ + " Fechasien >= '" & DESYY & "'"
'   SQLX$ = SQLX$ + CONDI$ + " AND Fechasien <= '" & HASYY & "'"
'   SQLX$ = SQLX$ + CONDI$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
'   SQLX$ = SQLX$ + " ORDER BY FECHASIEN, ORFEASIEN ASC;"
'   Set ValidacionAsientos = READSET(SQLX$)
'   With ValidacionAsientos
'        Do While Not .EOF
'            NumeroAsiento$ = SAFETEXT$(!NUIDASIEN)
'            EstadoEncabeAsien% = XVALO(!StatAsien)
'            EstadoDetaAsien% = XVALO(VALOBADA("DETAASIEN", "STATPASE", "NUIDASIEN='" & NumeroAsiento$ & "'"))
'            If EstadoEncabeAsien% < 0 And EstadoDetaAsien > 0 Then
'                Call ACTUREGISTRO("DETAASIEN", "STATPASE", "NUIDASIEN='" & NumeroAsiento$ & "'", "-1", REGAF&)
'                Call COMUNI("Se Ha Detectado una Inconsistencia en el Asiento '" & NumeroAsiento$ & "'\y Se Ha Procedido a Corregirlo\ \Debe Volver a Solicitar el Diario General.")
'                Inconsistencia% = 1
'            ElseIf EstadoEncabeAsien% > 0 And EstadoDetaAsien < 0 Then
'                Call ACTUREGISTRO("DETAASIEN", "STATPASE", "NUIDASIEN='" & NumeroAsiento$ & "'", "1", REGAF&)
'                Call COMUNI("Se Ha Detectado una Inconsistencia en el Asiento '" & NumeroAsiento$ & "'\y Se Ha Procedido a Re-Activarlo\ \Verificar el Asiento\ y Volver a Solicitar el Diario General.")
'                Inconsistencia% = 1
'            End If
'            .MoveNext
'        Loop
'   End With
'
'   If Inconsistencia% = 1 Then GoTo 99
   
   CONDI$ = "Fechasien >= '" & DESYY & "'"
   CONDI$ = CONDI$ + " AND Fechasien <= '" & HASYY & "'"
   CONDI$ = CONDI$ + " AND statpase >0  "
   CONDI$ = CONDI$ + " AND Niv_Clasi <= " & NICONFLIM% & " "
   FIN& = CantRegistros("DETAASIEN", CONDI$)
   '
   SQLX$ = "SELECT * FROM DetaAsien WITH(NOLOCK) WHERE " & CONDI$
   SQLX$ = SQLX$ + " ORDER BY FECHASIEN, ORFEASIEN, OrNupase ASC;"
   Set PASEASIEN = READSET(SQLX$)
   '
   JJ& = 0
10 With PASEASIEN
''''     On Error Resume Next
''''     .MoveLast
''''     If Err < 1 Then
''''       On Error GoTo 0
''''       FIN& = .RecordCount
''''       U& = 0
''''       .MoveFirst
       Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         IDEBE# = !IDebePase
         IHABE# = !IHabePase
         If Abs(IDEBE#) + Abs(IHABE#) < 0.005 Then GoTo 19
         Z$ = REBLA$
         Call REPLA(Z$, MKI$(CVINT(!FECHASIEN)), 1, 2)
         Call REPLA(Z$, !NUIDASIEN, 3, 16)
           ' AQUI FALTA NUMERO DE ASIENTO DEFINITIVO
             REFEX$ = REFGENAS(!NUIDASIEN)
         Call REPLA(Z$, REFEX$, 23, 64)
           ' AQUI FALTA REFERENCIA GENERAL DEL ASIENTO
         Call REPLA(Z$, CODCUE$(!CUECOINT), 87, 16)
           DEPASE$ = !RefePase
         Call REPLA(Z$, DEPASE$, 103, 64)
         Call REPLA(Z$, MKD$(IDEBE#), 167, 8)
         Call REPLA(Z$, MKD$(IHABE#), 175, 8)
         JJ& = JJ& + 1
         Call GRAREG("DIARPRIM", Z$, JJ&)
19     .MoveNext
       Loop
''''     End If
   End With
   PASEASIEN.Close
   Set PASEASIEN = Nothing
   ''''On Error GoTo 0
   '
   'GENERA ARCHIVO ASIDIARE CON SOLO LOS ASIENTOS DE DIARIO DEL RANGO ELEGIDO
   Call SETULTREG("DIARPRIM", JJ&)
   Call COPYSTRU("DIARPRIM", "ASDIARE")
   Call TRALOCAL("DIARPRIM", "ASDIARE")
   Call TRACENTRAL("ASDIARE", "")
   '
   Call ABRECABAN
''''   Dim CABATEMP As ADODB.Recordset
''''   Set CABATEMP = New ADODB.Recordset
   SQLX$ = "SELECT SUM(IDEBECOMP) AS SUDEBE, SUM(IHABECOMP) AS SUHABE FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or NuAsiDiar IS NULL) "
   SQLX$ = SQLX$ + " AND FechConta >= '" & FEINI & "'"
   SQLX$ = SQLX$ + " AND FechConta < '" & DESYY & "'"
   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
   SQLX$ = SQLX$ + " AND CodiEmpr = " & EMPREI%
''''   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   Set CABATEMP = READSET(SQLX$)
   '
   On Error Resume Next
   CABATEMP.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     If Not IsNull(CABATEMP!SUDEBE) Then TRANDEBE# = TRANDEBE# + CABATEMP!SUDEBE
     If Not IsNull(CABATEMP!SUHABE) Then TRANHABE# = TRANHABE# + CABATEMP!SUHABE
   End If
   On Error GoTo 0
   CABATEMP.Close
   '
   CONDI$ = "(NuAsidiar = '' or NuAsiDiar IS NULL) "
   CONDI$ = CONDI$ + " AND FechConta >= '" & DESYY & "'"
   CONDI$ = CONDI$ + " AND FechConta <= '" & HASYY & "'"
   CONDI$ = CONDI$ + " AND Niv_Clasi <= " & NICONFLIM%
   CONDI$ = CONDI$ + " AND CodiEMPR = " & EMPREI% & " "
   '
   FIN& = CantRegistros("CAJABANC", CONDI$)
''''   Set CABATEMP = New ADODB.Recordset
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " + CONDI$
''''   SQLX$ = SQLX$ + " WHERE (NuAsidiar = '' or NuAsiDiar IS NULL) "
''''   SQLX$ = SQLX$ + " AND FechConta >= '" & DESYY & "'"
''''   SQLX$ = SQLX$ + " AND FechConta <= '" & HASYY & "'"
''''   SQLX$ = SQLX$ + " AND Niv_Clasi <= " & NICONFLIM%
''''   SQLX$ = SQLX$ + " AND CodiEMPR = " & EMPREI% & " "
   SQLX$ = SQLX$ + " ORDER BY FECHCONTA ASC, FEHOREAL ASC, CODICOM_LAR;"
''''   CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   Set CABATEMP = READSET(SQLX$)
   '
   With CABATEMP
''''     On Error Resume Next
''''     .MoveLast
''''     If Err < 1 Then
''''       On Error GoTo 0
''''       FIN& = .RecordCount
''''       U& = 0
''''       .MoveFirst
''''       '
       Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         IDEBE# = !Idebecomp
         IHABE# = !IHABECOMP
         If Abs(IDEBE#) + Abs(IHABE#) < 0.005 Then GoTo 29
         Z$ = REBLA$
         Call REPLA(Z$, MKI$(CVINT(!FECHCONTA)), 1, 2)
         Call REPLA(Z$, !CodiCom_Cor, 3, 12)
         ' AQUI FALTA NUMERO DE ASIENTO DEFINITIVO
         ' AQUI FALTA REFERENCIA GENERAL DEL ASIENTO
         Call REPLA(Z$, CODCUE$(!CUECONTAI), 87, 16)
           DEPASE$ = !RefeCom
         Call REPLA(Z$, DEPASE$, 103, 64)
         Call REPLA(Z$, MKD$(IDEBE#), 167, 8)
         Call REPLA(Z$, MKD$(IHABE#), 175, 8)
           DEMOVI$ = !DESCRIMOV
         Call REPLA(Z$, DEMOVI$, 191, 40)
         JJ& = JJ& + 1
         Call GRAREG("DIARPRIM", Z$, JJ&)
29     .MoveNext
       Loop
''''     End If
   End With
   CABATEMP.Close
   Set CABATEMP = Nothing
   
   Screen.MousePointer = 1
   '
   'Call FILEDUMP("DIARPRIM")
   Call SETULTREG("DIARPRIM", JJ&)
   Call FILESORT("DIARPRIM", "", 1, 2, "ASC")
   TRANSPOR# = TRANDEBE# - TRANHABE#
   '
   DESX% = FECHANUM(Label5)
   HASX% = FECHANUM(Label6)
   FIN& = ULTREG&("DIARPRIM")
   SUMASI# = 0: ASIAN$ = ""
   HUBORECU% = 0
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("DIARPRIM", U&)
     FEAS% = CVI(Left$(X$, 2))
     '
     If FEAS% < DESX% Then
       TRANDEBE# = TRANDEBE# + CVD(Mid$(X$, 167, 8))
       TRANHABE# = TRANHABE# + CVD(Mid$(X$, 175, 8))
       TRANSPOR# = TRANDEBE# - TRANHABE#
     End If
     '
     If FEAS% >= DESX% Then
       If FEAS% <= HASX% Then
         If DIARPAN07.LIST1.ListCount = 0 Then
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, FECHATEX$(DIANTE(DESX%)), 1, 8)
           Call REPLA(TTXX$, "Transporte ...............................", 47, 32)
           Call REPLA(TTXX$, FORMATNUM$(TRANDEBE#, "G14.2"), 80, 14)
           Call REPLA(TTXX$, FORMATNUM$(TRANHABE#, "G14.2"), 94, 14)
           DIARPAN07.LIST1.AddItem TTXX$
           '
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, " ............ ", 80, 14)
           Call REPLA(TTXX$, "............. ", 94, 14)
           DIARPAN07.LIST1.AddItem TTXX$
         End If
         TTXX$ = Space$(128 + 40)
         Call REPLA(TTXX$, FECHATEX$(FEAS%), 1, 8)
           NIDAS$ = TRIM$(Mid$(X$, 3, 16))
           Call REPLA(TTXX$, NIDAS$, 11, 12)
           If NIDAS$ <> ASIAN$ Then
             If Abs(SUMASI#) >= 0.005 Then
               '
               If Asc(Left$(ASIAN$ + " ", 1)) > 57 Then ' SOLO ASIENTOS DE C&B
                 'Call RECUMOCAJ(ASIAN$, CVDAT(FEASAN%), OKEYY%)
                 Call RECUMOCAJ(ASIAN$, FEASAN%, OKEYY%)
                 If OKEYY% > 0 Then
                   HUBORECU% = 1
                   GoTo 89
                 End If
               End If
               '
               If COMALTER%("Comprobante " + ASIAN$ + " No Balanceado.\\Continuar\Verificar y Recuperar") = 2 Then
                 Call RECUNOBAL(ASIAN$, FEASAN%)
                 HUBORECU% = 1
                 ' Call COMUNI("Debe Ahora Volver a Pedir el Reporte.")
                 ' GoTo 99
                 Screen.MousePointer = 11
               End If
             End If
89           SUMASI# = 0
           End If
           ASIAN$ = NIDAS$
           FEASAN% = FEAS%
           NASDEF& = CVS(Mid$(X$, 19, 4))
         Call REPLA(TTXX$, FORMATNUM$(NASDEF, "F7.0"), 24, 7)
           COCUE$ = TRIM$(Mid$(X$, 87, 16))
         Call REPLA(TTXX$, COCUE$, 33, 12)
           RefPase$ = TRIM$(Mid$(X$, 103, 64))
         Call REPLA(TTXX$, RefPase$, 47, 32)
 'atencion
           IDEBE# = ROUND#(CVD(Mid$(X$, 167, 8)), 2)
           IHABE# = ROUND#(CVD(Mid$(X$, 175, 8)), 2)
           If IDEBE# < 0 Then IHABE# = IHABE# + Abs(IDEBE#): IDEBE# = 0
           If IHABE# < 0 Then IDEBE# = IDEBE# + Abs(IHABE#): IHABE# = 0
           SUMASI# = SUMASI# + IDEBE# - IHABE#
         Call REPLA(TTXX$, FORMATNUM$(IDEBE#, "G14.2"), 80, 14)
         Call REPLA(TTXX$, FORMATNUM$(IHABE#, "G14.2"), 94, 14)
           DEMOVI$ = Mid$(X$, 191, 40)
         Call REPLA(TTXX$, DEMOVI$, 129, 40)
         DIARPAN07.LIST1.AddItem TTXX$
       End If
     End If
   Next U&
   Call CIERRARCH("DIARPRIM")
   Screen.MousePointer = 1
   '
   If HUBORECU% > 0 Then
     Call COMUNI("Movimientos No Balanceados Recuperados.\  \Debe Volver a Pedir el Reporte.")
     GoTo 99
   End If
   '
   If Option1(1).Value = True Then
       Call REMPLAGRID(DIARPAN07, DIARPAN07.GRID2, DIARPAN07.Picture1, DIARPAN07.Label3, DIARPAN07.LIST1)
       DIARPAN07.Show 1
     Else
       Call DIARPAN07.LIDIARPRIN
   End If
   '
99 Call CIERRARCH("DIARPRIM")
   DIARPAN07.LIST1.Clear
   Command15.Enabled = True
   Screen.MousePointer = 1
   
End Sub

Sub RECUNOBAL(ASIAN$, FEASAN%)
   ' RECUPERACION COMPROBANTES NO BALANCEADOS
   If DERACCE%("RECUNOBA", "Recuperación Comprobantes no Balanceados") < 2 Then
     Exit Sub
   End If
   '
   FEASEL = FETEXCOR1$(FEASAN%)
   
   Dim CABASELE As ADODB.Recordset
   Set CABASELE = New ADODB.Recordset

   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + " WHERE CodiCom_Cor = '" & TRIM$(ASIAN$) & "' "
   SQLX$ = SQLX$ + " AND FechConta = '" & FEASEL & "' "
   '
   CABASELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   RECUNOBA.LIST1.Clear
   With CABASELE
     On Error Resume Next
     .MoveLast
     If Err > 0 Then
       On Error GoTo 0
       Call COMUNI("Imposible Recuperar Des-Balance - Consulte")
       Exit Sub
     End If
     On Error GoTo 0
     '
     NORI% = 0
     .MoveFirst
     ' LEE LOS DATOS ANTERIORES
     CoComCor$ = !CodiCom_Cor
     CoComLar$ = !CODICOM_LAR
     Referen$ = !RefeCom
     FeContab = !FECHCONTA
     FeAcredi = !fechacred
     CentroCos$ = SAFETEXT$(!CenCosto)
     ItConcep$ = SAFETEXT$(!ItemConc)
     NumeClie = 0
     NumeProv% = 0
     Descrip$ = SAFETEXT$(!DESCRIMOV)
     FechEmi = !FECHEMISI
     CodiEmp% = !CodiEmpr
     CodiSuc% = !CodiSucu
     NumeTal = !NumeTalo
     NumeCom = !NumeComp
     StatInit = 0
     MONEMI% = !Mone_Emis
     TICAI = !Tipo_Cam
     ValoDola# = 0
     BasImpo# = 0
     NumCai$ = !NROCAI
     InclIva$ = "NO"
     FechCai = !FEVTOCAI
     Observa$ = !OBSERGEN
     NiveCla% = !Niv_Clasi
     '
     Do While Not .EOF
       NORI% = NORI% + 1
       !ValAbsComp = NORI%
       .Update
       TTXX$ = Space$(80)
       Call REPLA(TTXX$, FORMATNUM$(NORI%, "I4"), 1, 4)
       Call REPLA(TTXX$, CODCUE$(!CUECONTAI), 7, 16)
       Call REPLA(TTXX$, !RefeCom, 25, 32)
         IDEBE# = !Idebecomp
         IHABE# = !IHABECOMP
       Call REPLA(TTXX$, FORMATNUM(IDEBE#, "F14.2"), 57, 14)
       Call REPLA(TTXX$, FORMATNUM(IHABE#, "F14.2"), 71, 14)
       RECUNOBA.LIST1.AddItem TTXX$
     .MoveNext
     Loop
   End With
   CABASELE.Close
   Set CABASELE = Nothing
   'jandy sql se cierra el recordset x estar en modo de bloqueo pesimista
   ' con lo cual no dejaba hacere el delete
   '
   RECUNOBA.List2.Clear
   RECUNOBA.Label6 = ""
   FEX = FEASAN%
   RECUNOBA.Caption = "Recuperación Comprobante No Balanceado " + TRIM$(ASIAN$) + " - " + FECHATEX$(FEX)
   Call RECUNOBA.TOTALI
   '
   Screen.MousePointer = 1
   RECUNOBA.Show 1
   '
   If RECUNOBA.Label6 = "OK" Then
      CUECOX$ = TRIM$(RECUNOBA.Label9)
      If CUECOX$ <> "" Then
         CuentaCo% = CUEINT%(CUECOX$)
         If CuentaCo% > 0 Then
           IMAJU# = XVALO(RECUNOBA.Label2) - XVALO(RECUNOBA.Label1)
           If Abs(IMAJU#) > 0 Then
              '
              TiMovi = "AJUSTE"            ' DEPOSITO BANCARIO
              OpMovi = "BALANCEO"           ' PLAZO DE ACREDITACION
              VaMovi = ""                   '
              Descrip = "Ajuste Balanceo de CC"
              ImpoDeb# = IMAJU#
              StatInit = 0
              Call GRACABAN
              '
            End If
         End If
      End If
      '
      For KU& = 1 To RECUNOBA.List2.ListCount
        TTXX$ = RECUNOBA.List2.List(KU& - 1)
        NORI% = Val(Mid$(TTXX$, 1, 4))
        'jandy sql
        strsql = "Delete from CajaBanc where ValAbsComp = " & NORI%
        strsql = strsql & " and  CodiCom_Cor = '" & TRIM$(ASIAN$) & "' "
        strsql = strsql & " and FechConta = '" & FEASEL & "' "
        Call EJECUTACOMANDO(strsql)
      Next KU&
   End If
   '
   strsql = "Update Cajabanc set ValAbsComp = Abs(IDEBECOMP - IHABECOMP)"
   strsql = strsql & " where  CodiCom_Cor = '" & TRIM$(ASIAN$) & "' "
   strsql = strsql & " and FechConta = '" & FEASEL & "' "
   Call EJECUTACOMANDO(strsql)
   '
End Sub



Private Sub Frame6_DblClick()
   DES1% = FECHANUM(Label5)
   HAS1% = FECHANUM(Label6)
   Call DESHASFECHA(DES1%, HAS1%, PERIO1$)
   If PERIO1$ <> "" Then
      Label5 = FECHATEX$(DES1%)
      Label6 = FECHATEX$(HAS1%)
   End If
End Sub

'
Private Sub label5_DblClick()
  Call Command21_Click
End Sub

Private Sub Label6_DblClick()
   Call Command4_Click
End Sub

Private Sub Text1_Change()
   Label1 = DENOCUE$(CUEINT%(Text1))
End Sub

Private Sub Text1_DblClick()
   Call Command8_Click
End Sub

Private Sub Text2_Change()
   Label2 = DENOCUE$(CUEINT%(Text2))
End Sub

Private Sub Command8_Click()
    Call SELECHO("ECUECON")
    PRICUEX$ = VALACT1$("ECUECON")
    If PRICUEX$ <> "" Then
        Text1 = PRICUEX$
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ECUECON")
    ULTCUEX$ = VALACT1$("ECUECON")
    If ULTCUEX$ <> "" Then
        Text2 = ULTCUEX$
    End If
End Sub

Private Sub Command21_Click()
   Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     FFII% = FECHANUM(FESELEC$)
     If FFII% >= COMEJ% Then
       If FFII% <= FINEJ% Then
         Label5 = FECHATEX$(FFII%)
         If FECHANUM(Label6) < FFII% Then
           Label6 = Label5
         End If
       End If
     End If
   End If
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   FESELEC$ = VALACT1$("SELFECHA")
   '
   If FESELEC$ <> "" Then
     FFII% = FECHANUM(FESELEC$)
     If FFII% >= COMEJ% Then
       If FFII% <= FINEJ% Then
         Label6 = FECHATEX$(FFII%)
         If FECHANUM(Label5) > FFII% Then
           Label5 = Label6
         End If
       End If
     End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub


Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Text2_DblClick()
   Call Command9_Click
End Sub


