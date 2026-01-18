VERSION 5.00
Begin VB.Form SELICEN2 
   BackColor       =   &H00D0EEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Licencias Vigentes"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0EEC0&
      Caption         =   "Validación de Ingreso x Usuarios para Sistemas  Multiempresas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   19
      Top             =   3300
      Width           =   6765
      Begin VB.CommandButton Command8 
         Caption         =   "Configurar Restricción de Ingreso de Usuarios Sistemas Multiempresa"
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
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Configuración para el caso de  multiempresas que usuarios pueden ingresar a  cada empresa."
         Top             =   240
         Width           =   6495
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D0EEC0&
         Caption         =   "Validar Ingreso de Usuarios para Sistemas Multiempresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   280
         Left            =   1080
         TabIndex        =   20
         ToolTipText     =   "Tildado el Sistema Valida En Sistemas Multiempresa los Derechos por Usuario"
         Top             =   240
         Width           =   5235
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Anular Licencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   4725
      TabIndex        =   14
      ToolTipText     =   "Asigna la Licencia Seleccionada a la Estación de Trabajo Actual "
      Top             =   2310
      Width           =   2115
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Re-Asignar Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Asigna la Licencia Seleccionada a la Estación de Trabajo Actual "
      Top             =   2310
      Width           =   2115
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cargar Nueva Licencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   210
      TabIndex        =   10
      ToolTipText     =   "Carga una Nueva Licencia a Partir de 'Diskette' de Habilitación de Terminal"
      Top             =   2310
      Width           =   2220
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4005
      Left            =   7035
      ScaleHeight     =   3945
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox Picture16 
         Height          =   3945
         Left            =   0
         ScaleHeight     =   3885
         ScaleWidth      =   1020
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   3900
            Left            =   0
            Picture         =   "SELICEN2.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
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
         Height          =   600
         Left            =   120
         Picture         =   "SELICEN2.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   860
         Width           =   600
      End
      Begin VB.CommandButton Command7 
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
         Left            =   120
         Picture         =   "SELICEN2.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Conexión a SQL-SERVER"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "SELICEN2.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   4200
         Width           =   650
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   3240
         Width           =   650
         Begin VB.Frame Frame2 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   120
         Picture         =   "SELICEN2.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   4755
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Height          =   600
         Left            =   120
         Picture         =   "SELICEN2.frx":2E00
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   250
         Width           =   600
      End
      Begin VB.CommandButton Command2 
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
         Height          =   600
         Left            =   120
         Picture         =   "SELICEN2.frx":2F4A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ingresa al Sistema"
         Top             =   1460
         Width           =   600
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -315
         Picture         =   "SELICEN2.frx":3254
         ToolTipText     =   "Doble-Click para (Re)-Habilitar Estación de Trabajo"
         Top             =   3540
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   8
         Top             =   1520
         Width           =   645
      End
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   525
      Width           =   6595
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nota:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   13
      Top             =   2730
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"SELICEN2.frx":5176
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   720
      TabIndex        =   12
      Top             =   2730
      Width           =   6105
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre de Terminal                Id-Ter  Nro-Serie"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   210
      TabIndex        =   9
      Top             =   250
      Width           =   6600
   End
   Begin VB.Line Line2 
      X1              =   6820
      X2              =   6820
      Y1              =   250
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   6820
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "SELICEN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call Command2_Click
End Sub

Private Sub Command1_Click()
   Close: End
End Sub

Private Sub Command2_Click()
   Hide
End Sub

Private Sub Command29_Click()
    'Call HELPONLINE("AJUSTOCK")
    MsgBox "Ayuda en Línea no Encontrada"
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CARLICENUE
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If List2.ListIndex < 0 Or List2.ListIndex >= List2.ListCount Then
      MsgBox "Asignación Imposible. Debe Seleccionar Licencia a Asignar."
      GoTo 99
   End If
   '
   SEVOLC$ = SERIVOL$("C:")
   For KKI% = 1 To List2.ListCount
     If TRIM$(Mid$(List2.List(KKI% - 1), 43, 12)) = SEVOLC$ Then
        MsgBox "Asignación Imposible !!!   La Estación Actual" + Chr$(13) + Chr$(10) + "ya Tiene Cargada Licencia en Entorno Activo."
        List2.ListIndex = KKI% - 1
        GoTo 99
     End If
   Next KKI%
   '
   IDX1$ = RANDSTRING$(-6)
   VDEF$ = Space$(52)
   Call REPLA(VDEF$, IDX1$, 1, 6)
   Call REPLA(VDEF$, SEVOLC$, 41, 12)
70 OBX$ = InputBox$("Nombre Identificador de Terminal", "Identificador Estación de Trabajo", NTX$)
   If OBX$ = "" Then Close: End
   '
   IDT$ = TRIM$(OBX$)
   If Len(IDT$) < 4 Then
      MsgBox ("Ingrese un Nombre para la Terminal (min. 4 pos.)")
      GoTo 70
   End If
   '
   NTID$ = Space$(38)
   Call REPLA(NTID$, Left$(OBX$, 32), 1, 32)
   Call REPLA(NTID$, IDX1$, 33, 6)
   IDENTER$ = TRIM$(IDX1$)
   NOMTER$ = TRIM$(Left$(OBX$, 32))
   IDENCRIP$ = ENCRIPTA$(NTID$)
   '
   SEVOLK$ = TRIM$(Mid$(List2.TEXT, 43, 12))
   IDST$ = LOADFILE$("IDENSYS.DRV")
   IDSX$ = Left$(IDST$, 960)
   While Len(IDSX$) < 960
     IDSX$ = IDSX$ + Chr$(0)
   Wend
   '
   JJ& = 0
   For KU& = 961 To Len(IDST$) Step 96
     RECID$ = Mid$(IDST$, KU&, 94)
     TTX1$ = TRIM$(DESENCRI$(Mid$(RECID$, 8, 12)))
     If TTX1$ <> SEVOLK$ Then
         JJ& = JJ& + 1
         Call REPLA(RECID$, Chr$(JJ&), 5, 1)
         SUCAR% = 0: For KI% = 1 To 94
                       SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
                     Next KI%
         RECID$ = RECID$ + MKI$(SUCAR%)
         IDSX$ = IDSX$ + RECID$
       Else
         SEVOLK$ = ""
     End If
   Next KU&
   RECID$ = RANDSTRING$(94)
   Call REPLA(RECID$, ENCRIPTA$(SEVOLC$ + "   "), 8, 12)
   Call REPLA(RECID$, IDENCRIP$, 38, 38)
   JJ& = JJ& + 1
   Call REPLA(RECID$, Chr$(JJ&), 5, 1)
   SUCAR% = 0: For KI% = 1 To 94
                 SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
               Next KI%
   RECID$ = RECID$ + MKI$(SUCAR%)
   IDSX$ = IDSX$ + RECID$
   Call SAVEFILE("IDENSYS.DRV", IDSX$)
   '
   MsgBox "Operación Completada - Reinicie la Aplicación", 16
   Close: End
   '
99 Command5.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If List2.ListIndex < 0 Or List2.ListIndex >= List2.ListCount Then
      MsgBox "Operación Imposible. Debe Seleccionar Licencia a Anular."
      GoTo 99
   End If
   '
   SEVOLK$ = TRIM$(Mid$(List2.TEXT, 43, 12))
   RESPU% = MsgBox("Desea  Confirmar  la Anulación  de la  Licencia" + Chr$(13) + Chr$(10) + "Correspondiente  a  la  Terminal  '" + SEVOLK$ + "'" + Chr$(13) + Chr$(10) + "Dejándola Reservada para Asignación Futura ?", 512 + 35)
   If RESPU% = 6 Then   ' RESPUESTA = 'SI'
      IDST$ = LOADFILE$("IDENSYS.DRV")
      IDSX$ = Left$(IDST$, 960)
      While Len(IDSX$) < 960
        IDSX$ = IDSX$ + Chr$(0)
      Wend
      '
      JJ& = 0
      For KU& = 961 To Len(IDST$) Step 96
        RECID$ = Mid$(IDST$, KU&, 94)
        TTX1$ = TRIM$(DESENCRI$(Mid$(RECID$, 8, 12)))
        If TTX1$ <> SEVOLK$ Then
          JJ& = JJ& + 1
          Call REPLA(RECID$, Chr$(JJ&), 5, 1)
          SUCAR% = 0: For KI% = 1 To 94
                        SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
                      Next KI%
          RECID$ = RECID$ + MKI$(SUCAR%)
          IDSX$ = IDSX$ + RECID$
        End If
      Next KU&
      RECID$ = RANDSTRING$(94)
      Call REPLA(RECID$, ENCRIPTA$("0000-0000   "), 8, 12)
      Call REPLA(RECID$, ENCRIPTA$("* Licencia Reservada Uso Futuro 000000"), 38, 38)
      JJ& = JJ& + 1
      Call REPLA(RECID$, Chr$(JJ&), 5, 1)
      SUCAR% = 0: For KI% = 1 To 94
                    SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
                  Next KI%
      RECID$ = RECID$ + MKI$(SUCAR%)
      IDSX$ = IDSX$ + RECID$
      Call SAVEFILE("IDENSYS.DRV", IDSX$)
      '
      MsgBox "Operación Completada - Reinicie la Aplicación", 16
      Close: End
      '
   End If
   '
99 Command6.Enabled = True
   '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If TRIM$(LUDAT$) = "" Then
         MsgBox "Acceso Imposible - Primero Establezca Carpeta de Datos."
      Else
         CONN_SQLS.Label1 = LUDAT$
         CONN_SQLS.Show 1
         Unload Me
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   If Command8.Caption = "Grabar" Then
     COEMP$ = "00" & CodiEmp%
     COEMP$ = Right(COEMP$, 2)
     Call FLEXLOG3.GRABARCONTROL("EMPRES" & COEMP$, COEMP$, TRIM$(Str$(Check2.Value)))
     Close: End
   End If
   '
   Static CAIN%
   
   CLAX$ = RANDSTRING$(-6)
   CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
   If CONTRASE$ <> CONTRACLA$(CLAX$) Then GoTo 40
   
   Command8.Width = 800
   Command8.Caption = "Grabar"
   GoTo 99
 
    
40  MsgBox "Usuario no Definido"
    CAIN% = CAIN% + 1
    If CAIN% > 3 Then
        MsgBox "Superado Limite Intentos de Acceso !!!"
        Close: End
    End If
'    CLA$ = "": CLB$ = "": CLAVE$ = ""

'   If CLACONTRA% < 1 Then GoTo 99
99 Command8.Enabled = True
   '
End Sub


Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    'se comenta el nueva aspecto por que en el flogon no existe la funcion control y no se puede compilar
'    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
'      Picture16.Visible = True
'      Picture16.ZOrder 0
'      Me.BackColor = &HFCD7B6
'      Me.Frame2.BackColor = &HFCD7B6
''      MARCOBARRA.Top = 2950
''      Set MARCOBARRA.Container = Me.Picture16
''      MARCOBARRA.ZOrder 0
'      ' &H8000000F
'     End If

    IDST$ = LOADFILE$("IDENSYS.DRV")
    List2.Clear
    '
    JJ& = 0
    For KU& = 961 To Len(IDST$) Step 96
      RECID$ = Mid$(IDST$, KU&, 96)
      JJ& = JJ& + 1
      SUCAR% = 0: For KI% = 1 To 94
                    SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
                  Next KI%
      If CVI(Mid$(RECID$, 95, 2)) = SUCAR% Then
        TTX1$ = DESENCRI$(Mid$(RECID$, 8, 12))
        TTX2$ = DESENCRI$(Mid$(RECID$, 38, 38))
        TTX3$ = Mid$(TTX2$, 33, 6)
        TTX2$ = Left$(TTX2$, 32)
        If (TRIM$(TTX1$) = "0000-0000") Then
          If Asc(Mid$(RECID$, 5, 1)) = JJ& Then GoTo 14
          GoTo 19
        End If
        '
        For KK1% = 1 To List2.ListCount
          If Mid$(List2.List(KK1% - 1), 43, 12) = TTX1$ Then GoTo 19
        Next KK1%
        '
14      TTXX$ = Space$(54)
        Mid$(TTXX$, 1, 32) = TTX2$
        Mid$(TTXX$, 35, 6) = TTX3$
        Mid$(TTXX$, 43, 12) = TTX1$
        List2.AddItem TTXX$
      End If
19  Next KU&

20  CONNSQL$ = TRAPASS1$(TRIM$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV")))
    If CONNSQL$ <> "" Then
      ' VERIFICA LA CONEXION
      If FLEXLOG3.VERICONEC(CONNSQL$) > 0 Then
           COEMP$ = "00" & CodiEmp%
           COEMP$ = Right(COEMP$, 2)
           CONFI% = Val(FLEXLOG3.MUESTRACONTROL$("EMPRES" & COEMP$, COEMP$))
           Check2.Value = CONFI%
           Command8.Caption = Command8.Caption & " (" & TRIM$(Str$(CONFI%)) & ")" ' PARA QUE QUEDE VISIBLE SI ESTA MARCADO EL CHECK
      End If
   End If

    '
End Sub

Private Sub List2_DblClick()
   Call Command5_Click
End Sub

Sub CARLICENUE()    ' habilita puesto de trabajo
    '
    Dim BLEC As String * 127
    Dim IDS As String * 96
    Dim IDXX As String * 40
    '
    PORDISKETTE% = 0
    SEVOLC$ = Left$(SERIVOL$("C:") + Space$(12), 12)
    SVLC$ = ENCRIPTA$(TRIM$(SEVOLC$))
    TZXC$ = ""
    For KKI% = 1 To Len(SVLC$)
      TZXC$ = TZXC$ + Str$(Asc(Mid$(SVLC$, KKI%, 1)))
    Next KKI%
    'MsgBox SVLC$ & "    " & TZXC$
    If TRIM$(TZXC$) = "109 119 155 158 177 224 235 13 17" Then GoTo 60
    If UCase$(CMD$) = "FLCACHE" Then GoTo 60
    '
    DEXX$ = LOADFILE("IDENSYS.DRV")
    If Len(DEXX$) >= 96 Then
      DEYY$ = DESENCRI$(Left$(DEXX$, 94))
      EMPREAC$ = UCase$(TRIM$(Mid$(DEYY$, 13, 30)))
    End If
    '
    ' ALTA DE NUEVAS TERMINALES POR MEDIO DE CLAVE Y CONTRASEÑA EN LUGAR DE POR DISKETTE
    CLAX$ = RANDSTRING$(-6)
    CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
    If CONTRASE$ = CONTRACLA$(CLAX$) Then GoTo 60
'    If InStr(EMPREAC$, "BRASBOM") > 0 Then GoTo 60
'    If InStr(EMPREAC$, "GUIDI") > 0 Then GoTo 60
'    If InStr(EMPREAC$, "FUINYTER") > 0 Then GoTo 60
'    If InStr(EMPREAC$, "ACHERNAR") > 0 Then GoTo 60
'    If InStr(EMPREAC$, "ARCON") > 0 Then GoTo 60
'    GoTo 60
    Exit Sub
    '
10  Screen.MousePointer = 1
    PORDISKETTE% = 1
    OPEX% = MsgBox(FORMESS$("Coloque 'Diskette' de Instalación\en la Disketera 'A:'."), 17, "Habilitación Estación de Trabajo")
    If OPEX% <> 1 Then Exit Sub
    '
    FLEN& = 0
    On Error GoTo 25
    FLEN& = FileLen("A:FLEXKEY.KEY")
    On Error GoTo 0
    GoTo 30
    '
25  ERRX = Err.Number
    Resume 27
27  On Error GoTo 0
    If ERRX = 53 Then
         MsgBox ("Diskette Incorrecto.")
       Else
         MsgBox ("Error de Acceso a la Disketera.")
    End If
    GoTo 10
    '
30  If FLEN& < 127& * 499 Then
       MsgBox ("Diskette Incorrecto.")
       GoTo 10
    End If
    '
40  On Error GoTo 45
    N1% = FreeFile
    Open "A:FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
    Get #N1%, 1, BLEC$
    Put #N1%, 1, BLEC$
    On Error GoTo 0
    GoTo 50
    '
45  Resume 46
46  On Error GoTo 0
    MsgBox (FORMESS$("Diskette Protegido Contra Escritura.\Quite la Faja de Protección\e Intente de Nuevo."))
    GoTo 10
    '
50  Screen.MousePointer = 11
    SEVOLA$ = Left$(SERIVOL$("A:") + Space$(12), 12)
'    SEVOLC$ = Left$(SERIVOL$("C:") + Space$(12), 12)
    CULOG$ = Left$(CurDir, 2)
    SEVOLF$ = Left$(SERIVOL$(CULOG$) + Space$(12), 12)
    If TRIM$(SEVOLA$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox ("Diskette No Válido (Err.50)")
       GoTo 10
    End If
    '
    If TRIM$(SEVOLF$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox (FORMESS$("Error de Acceso a Disco '" + CULOG$ + "'.\Consulte a su Soporte de Sistemas."))
       Screen.MousePointer = 11
       GoTo 10
    End If
    '
    For UL& = 1 To FLEN& / 127
       Get #N1%, UL&, BLEC$
       SUCAR% = 0
       For KI% = 1 To 125
          SUCAR% = SUCAR% + Asc(Mid$(BLEC$, KI%, 1)) - 83
       Next KI%
       If CVI(Mid$(BLEC$, 126, 2)) <> SUCAR% Then
          Close #N1%
          Screen.MousePointer = 1
          MsgBox ("Diskette No Válido (Err.51)")
          GoTo 10
       End If
    Next UL&
    Screen.MousePointer = 1
    '
    Get #N1%, 1, BLEC$
    If DESENCRI$(Mid$(BLEC$, 17, 12)) <> SEVOLA$ Then
       Close #N1%
       MsgBox ("Diskette No Válido (Err.52)")
       GoTo 10
    End If
    '
    Get #N1%, 3, BLEC$
    SDES$ = DESENCRI$(Mid$(BLEC$, 71, 12))
    If SDES$ <> String$(12, 23) Then
       Close #N1%
       MsgBox (FORMESS("Diskette no Válido o ya Utilizado."))
       GoTo 10
    End If
    '
    Get #N1%, 2, BLEC$
    SDES$ = DESENCRI$(Mid$(BLEC$, 29, 12))
    If SDES$ <> String$(12, 19) Then
       If SDES$ <> SEVOLC$ Then
          N2% = FILEOPEN%("IDENSYS.DRV", 0, 96)
          For i% = 11 To LOF(N2%) / 96
            Get #N2%, i%, IDS$
            SUCAR% = 0
            For KI% = 1 To 94
               SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
            Next KI%
            If CVI(Mid$(IDS$, 95, 2)) = SUCAR% Then
               SVLX$ = DESENCRI$(Mid$(IDS$, 8, 12))
               If SVLX$ = SDES$ Then
                 IDENCRIP$ = Mid$(BLEC$, 59, 38)
                 NOTERVIE$ = TRIM$(Left$(DESENCRI$(IDENCRIP$), 32))
                 MENSA$ = "El Diskette de Instalación Utilizado"
                 MENSA$ = MENSA$ + "\Corresponde a la Terminal Denominada"
                 MENSA$ = MENSA$ + "\'" + NOTERVIE$ + "'." + "\  "
                 MENSA$ = MENSA$ + "\Si Continúa con la Instalación,"
                 MENSA$ = MENSA$ + "ese Puesto de Trabajo Quedará DESHABILITADO.\   "
                 MENSA$ = MENSA$ + "\Desea Cancelar o Continuar ?"
                 If MsgBox(FORMESS$(MENSA$), 273, "Habilitación de Estación de Trabajo") <> 1 Then
                   Close: End
                 End If
                 '
                 IDS$ = String$(96, 0)
                 Put #N2%, i%, IDS$
                 '
                 Get #N1%, 2, BLEC$
                 TTXX$ = Left$(BLEC$, 125)
                 Mid$(TTXX$, 29, 12) = ENCRIPTA$(String$(12, 19))
                 SUCAR% = 0
                 For KI% = 1 To 125
                   SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
                 Next KI%
                 BLEC$ = TTXX$ + MKI$(SUCAR%)
                 Put #N1%, 2, BLEC$
                 Close #N2%
                 GoTo 60
               End If
            End If
          Next i%
          '
          Close #N1%
          Close #N2%
          MsgBox (FORMESS$("Diskette no Válido\Ya Utilizado en Otra Instalación."))
          GoTo 10
          '
       End If
    End If
    '
60  ID = Timer: While ID > 9999.99: ID = ID - 9000: Wend
    ID = Int(100 * ID)
    IDX$ = Mid$(Str$(ID), 2)
    While Len(IDX$) < 6: IDX$ = "0" + IDX$: Wend
    NTX$ = ""
    '
    If EXISTE%(LUCOM$ + "IDENTER.DRV") > 0 Then
      IDTX$ = Left$(LOADFILE$(LUCOM$ + "IDENTER.DRV"), 40)
      If Len(IDTX$) >= 40 Then
        SUDI% = 0: For i% = 1 To 38
                     SUDI% = SUDI% + Asc(Mid$(IDTX$, i%, 1))
                   Next i%
        If CVI(Mid$(IDTX$, 39, 2)) = SUDI% Then
          IDTY$ = DESENCRI$(Left$(IDTX$, 38))
          IDX$ = Mid$(IDTY$, 33, 6)
          NTX$ = Left$(IDTY$, 32)
        End If
      End If
    End If
    '
    VDEF$ = Space$(52)
    Call REPLA(VDEF$, IDX$, 1, 6)
    Call REPLA(VDEF$, SEVOLC$, 41, 12)
70  OBX$ = InputBox$("Nombre Identificador de Terminal", "Identificador Estación de Trabajo", NTX$)
    If OBX$ = "" Then Close: End
    '
    IDT$ = TRIM$(OBX$)
    If Len(IDT$) < 4 Then
       MsgBox ("Ingrese un Nombre para la Terminal (min. 4 pos.)")
       GoTo 70
    End If
    '
    NTID$ = Space$(38)
    Call REPLA(NTID$, Left$(OBX$, 32), 1, 32)
    Call REPLA(NTID$, IDX$, 33, 6)
    IDENTER$ = TRIM$(IDX$)
    NOMTER$ = TRIM$(Left$(OBX$, 32))
    IDENCRIP$ = ENCRIPTA$(NTID$)
    '
    If PORDISKETTE% > 0 Then
      Get #N1%, 2, BLEC$
      TTXX$ = Left$(BLEC$, 125)
      Call REPLA(TTXX$, ENCRIPTA$(SEVOLC$), 29, 12)
      Call REPLA(TTXX$, IDENCRIP$, 59, 38)
      SUCAR% = 0
      For KI% = 1 To 125
         SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
      Next KI%
      BLEC$ = TTXX$ + MKI$(SUCAR%)
      Put #N1%, 2, BLEC$
      Close #N1%
    End If
    '
    N2% = FILEOPEN%("IDENSYS.DRV", 0, 96)
    REI% = 1 + LOF(N2%) / 96: If REI% < 11 Then REI% = 11
    TTXX$ = RANDSTRING$(94)
    Call REPLA(TTXX$, ENCRIPTA$(SEVOLC$), 8, 12)
    Call REPLA(TTXX$, IDENCRIP$, 38, 38)
    SUCAR% = 0
    For KI% = 1 To 94
       SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
    Next KI%
    IDS$ = TTXX$ + MKI$(SUCAR%)
    Put #N2%, REI%, IDS$
    Close #N2%
    '
    N3% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    X$ = ENCRIPTA$(NTID$)
    SUDI% = 0: For i% = 1 To Len(X$)
                  SUDI% = SUDI% + Asc(Mid$(X$, i%, 1))
               Next i%
    Y$ = String$(40, 0)
    Call REPLA(Y$, X$, 1, 38)
    Call REPLA(Y$, MKI$(SUDI%), 39, 2)
    IDXX$ = Y$
    Put #N3%, 1, IDXX$
    '
    Y$ = String$(40, 0)
    Call REPLA(Y$, "B800", 1, 4)
    Call REPLA(Y$, MKS$(47104), 5, 4)
    Call REPLA(Y$, "SI", 9, 2)
    X$ = ENCRIPTA$(SEVOLC$)
    Call REPLA(Y$, X$, 11, 12)
    SUDI% = 0: For i% = 1 To 38
                 SUDI% = SUDI% + Asc(Mid$(Y$, i%, 1))
               Next i%
    Call REPLA(Y$, MKI$(SUDI%), 39, 2)
    IDXX$ = Y$
    Put #N3%, 2, IDXX$
    Close #N3%: N3% = 0
    '
    MsgBox (FORMESS$("Habilitación de Terminal '" + NOMTER$ + "' Completada Exitosamente.\  \Re-Inicie Ahora el Sistema."))
    '
99 End Sub


Private Sub salir_Click()
   Call Command1_Click
End Sub
