VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SIMUCARM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simulaciones de Fabricación"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   12750
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1815
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
      Height          =   330
      Left            =   11880
      TabIndex        =   22
      Top             =   1680
      Width           =   160
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Genera OF's de la Simulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   20
      ToolTipText     =   "Grabar Detalle de Simulación"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.TextBox TXTCOMPRA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      MaxLength       =   64
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nueva Simulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      ToolTipText     =   "Grabar Detalle de Simulación"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.TextBox TXTREFERENCIA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      MaxLength       =   64
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2400
      TabIndex        =   14
      ToolTipText     =   "Despliega Lista de Selección"
      Top             =   240
      Width           =   160
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar y Ejecutar Cálculo de Carga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7680
      TabIndex        =   10
      Top             =   4800
      Width           =   1605
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Grabar Simulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9360
      TabIndex        =   9
      ToolTipText     =   "Grabar Detalle de Simulación"
      Top             =   4800
      Width           =   1605
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
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
      Left            =   12000
      TabIndex        =   6
      Top             =   240
      Width           =   160
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   7800
      ScaleHeight     =   165
      ScaleWidth      =   990
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "SIMUCARM.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   6240
      OleObjectBlob   =   "SIMUCARM.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click Sobre Columna 0 Para Seleccionar Proyecto"
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   5
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10800
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10800
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. de Simulación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   160
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Última Simulación Calculada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   4840
      Width           =   1455
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Entrega General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   9480
      TabIndex        =   8
      Top             =   160
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   4080
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   4560
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "SIMUCARM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONSIMULACION%
Sub BLANFORMU()
   
   Me.Text17 = ""
   Me.Text2 = ""
   Me.TXTREFERENCIA = ""
   MSF0.Rows = 1
   TXTCOMPRA = ""
   Text3 = ""

End Sub

Sub GuardaSimulacion(Modo As Integer)

      SIMULA& = XVALO(Text2)
      '
      Dim CILOTE%(256), CALOTE#(256), CALOTE1#(256), PUNMAT$(256), PUNOPE$(256)
      Dim CISUMI%(512), CASUMI#(512)
      Dim MODOEXPLO%(256)

      CONDI$ = ""
      If MSF0.Rows <= 1 Then
         Call COMUNI("Falta Cargar Detalle de Simulación")
         Call BLOQUEAR_CONTROLES(Me, 1, 1)
         GoTo 99
      End If
      '
      CONDI2$ = " IdSimulOR =" & SIMULA&
      If CantRegistros("SIMU_ORDEFABR", CONDI2$, "NOMESS") < 1 Then ESNUEVO% = 1
      '
      'CONFIRMACION DE TRANSACCION
      If ESNUEVO% > 0 Then
'         OPX% = COMALTER%("Atención !!\Se va a Generar Una Nueva Simulación\Segun Detalle en Pantalla\\Cancelar\Continuar")
'         If OPX% < 2 Then GoTo 99
      Else
'         OPX% = COMALTER%("Atención !!\Se va a Modificar la Simulación Activa\Segun Detalle en Pantalla\\Cancelar\Continuar")
'         If OPX% < 2 Then GoTo 99
        CONDI$ = "IdSimulOR = " & SIMULA&
        Call ACTUREGISTRO("SIMU_ORDEFABR", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        CONDI$ = "IdSimulOP = " & SIMULA&
        Call ACTUREGISTRO("SIMU_OPERFAB", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")

      End If
      '
      '
      GESUBOR1% = 0
      If CONTROL$("", "GESUBOR") = "SI" Then
           GESUBOR1% = 1
      End If
      '
      'CASO SI ES NUEVO
      If ESNUEVO% > 0 Then
        SIMULA& = NuevaSimulacion&
      End If
      '
      '
      Screen.MousePointer = 11
      SQLX$ = " SELECT * FROM SIMU_ORDEFABR WHERE NUMEORFA  = " & SIMULA&
      Dim ORDEFFFABR1 As ADODB.Recordset
      Set ORDEFFFABR1 = New ADODB.Recordset
25    On Error Resume Next
      ORDEFFFABR1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      '
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      '
      For I& = 1 To MSF0.Rows - 1
      
        CI1% = MSF0.TextMatrix(I&, 1)
        CAN# = XVALO(MSF0.TextMatrix(I&, 4))
        FENTRE% = CVINT(MSF0.TextMatrix(I&, 5))
        NORDE% = 0
        CAPEN# = CAN#
        LOTE# = CAN#
        '
        While CAPEN# > 0
            CALOT# = LOTE#
            If CALOT# > CAPEN# Then
                CALOT# = CAPEN#
            End If
            '
            NORDE% = NORDE% + 1
            CILOTE%(NORDE%) = CI1%
            CALOTE#(NORDE%) = CALOT#
            CAPEN# = CAPEN# - CALOT#
        Wend

        If GESUBOR1% < 1 Then NORDE% = 1
        If GESUBOR1% > 0 Then
                Call EXPLOMUL(CI1%)
                For II% = 1 To CONTAMUL%
                    PAD%(II%) = 1
                Next II%
                PAD%(CONTAMUL%) = 0
                '
                For II% = 1 To CONTAMUL% - 1
                    If NIVE%(II% + 1) <= NIVE%(II%) Then
                        PAD%(II%) = 0
                        Debug.Print "(0)    " & CODEXT$(CIK%(II%))
                      Else
                        FIN& = ULTREG&("ARTIPRO")
                        Debug.Print "(1)    " & CODEXT$(CIK%(II%))
                        For j& = 1 To FIN&
                          X$ = REGLEIDO$("ARTIPRO", j&)
                            If CVI(Left$(X$, 2)) = CIK%(II%) Then
                              If CVI(Mid$(X$, 3, 2)) > 0 Then
                                PAD%(II%) = 0
                                GoTo 29
                              End If
                            End If
                        Next j&
                    End If
29                  Next II%
                '
                ' CONFIGURACION PARA TRABAJAR CON LA MODALIDAD LINEAL + LINEALIZADA
                NIVELMAXIMO% = XVALO(CONTROL("SIMU_ORDEFABR", "NIVMAXEX"))
                '
                For II% = 2 To CONTAMUL%
                    If MODOMIXTO% = 0 Then ' LINEALIZADA POR SEPARADO
                        If PAD%(II%) = 1 Then
                            NORDE% = NORDE% + 1
                            CILOTE%(NORDE%) = CIK%(II%)
                            CALOTE#(NORDE%) = CAN# * CANT(II%)
                        End If
                        ' ESTO APLICA CUANDO SE ENCUENTRA CONFIGURADA LA MODALIDAD LINEAL + LINEALIZADA
                        ' EL NIVEL MAXIMO O PUNTO DE QUIEBRE SE ALMACENA EN NIVELMAXIMO%. HASTA EL NIVEL ANTERIOR
                        ' AL MAXIMO SE APLICA EL MODO LINEA, A PARTIR DEL MAXIMO EN ADELANTE SE APLICA LINEALIZADA.
                        ' MODOMIXTO CORRESPONDE A LA CONFIGURACION LINEAL + LINEALIZADA
                        ' EN CASO DE ESTAR CONFIGURADA LA MODALIDAD PERO NO TIENE UN NIVEL MAXIMO, SE APLICA LINEALIZADA
                        ' POR DEFAULT, DEL SEGUNDO EN ADELANTE.
                    ElseIf MODOMIXTO% = 1 Then
                        If NIVELMAXIMO% <= 0 Then NIVELMAXIMO% = 1 ' NIVEL POR DEFAULT
                        ' PAD%() CONTIENE LA MARCA DE TODOS LOS SUBCONJUNTOS. SOLO SE AGREGAN A LOS VECTOES
                        ' CILOTE% Y CALOTE, HASTA EL NIVEL MAXIMO.
                        If PAD%(II%) = 1 And NIVE%(II%) <= NIVELMAXIMO% Then
                            NORDE% = NORDE% + 1
                            CILOTE%(NORDE%) = CIK%(II%)
                            CALOTE#(NORDE%) = CAN# * CANT(II%)
                            '
                            ' EN MODOEXPLO% SE DEJA UNA MARCA QUE INDICA QUE DE AHI EN ADELANTE COMIENZA EL MODO LINEALIZADA
                            If NIVE%(II%) = NIVELMAXIMO% Then
                                MODOEXPLO%(NORDE%) = 1
                            End If
                        End If
                    End If
                Next II%
            End If
        '
        'ProxiNumSimulaOF& ES UNA FUNCION QUE TRAE EL NUMERO PROXIMO
        'EL ID DE SIMULOR SE REPITE, EL NUMEORFA SE VA INCREMENTANDO
        NUMORF& = ProxiNumSimulaOF&
        NBRS$ = TRIM$(Str$(NUMORF&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
        NBRS$ = "SF-" + NBRS$ + "-"
        
        
        'GRABAR OPERFAB

        For II% = 1 To NORDE%
          '
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          DESCRI$ = TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT$(CILOTE%(II%)))

          With ORDEFFFABR1
            .AddNew
            !IdSimulOR = SIMULA&
            ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !NumeOrFa = NUMORF&
            !REFEREN = AJUSTI$(DESCRI$, 64)
            !idsubor = NORFA$
            !FECHGEN = Now
            !COD_INTE = CILOTE%(II%)
            !COD_EXTE = Left$(CODEXT$(CILOTE%(II%)), 16)
            !DESCRIP = Left$(DESCRIT0$(CILOTE%(II%)), 64)
            !Unidad = Left$(UNIMED$(CILOTE%(II%)), 4)
            !CANPROY = CALOTE#(II%)
                DESTI$ = TRIM$(Text3)
                If DESTI$ = "" Then DESTI$ = "STOCK"
            !DESTINO = DESTI$
            !FechEnSol = CVDAT(FENTRE%)
            !Priorid = PRIO%
            !StatuOrf = 1 ' ORIGINAL ES 0 PERO SI NO NO LA TOMA EN EL CALCULO DE CARMAQ
            !FechLan = CVDAT(HOY(HO$))
            !SecciPro = 0
            !Fechcierr = CVDAT(0)
            !CANAPRO = 0
            !CanRech = 0
            !IdenRetra = ""
            !FechProyEntre = CVDAT(0)
            !FechTempEntre = CVDAT(0)
            !FechRecal = CVDAT(0)
            !Observa = Left$(Me.TXTREFERENCIA, 64)
            !Origen = Left$(Me.TXTREFERENCIA, 64)
            !ID_ORFA = Left$(TXTCOMPRA, 24)
            !NROPED = 0
            !MARBORRA = 0
            '
          .Update
          End With
          Next II%
          
          '
          
          For II% = 1 To NORDE%
            NORFA$ = NBRS$ + TRIM$(Str$(II%))
            CIJK% = CILOTE%(II%)
            CAPROYE# = CALOTE#(II%)
            '
            SQLX$ = "SELECT * FROM SecOper WITH(NOLOCK) "
            SQLX$ = SQLX$ + "WHERE CodIConj=" & CIJK% & " "
            SQLX$ = SQLX$ + " AND StatOper = 1 "
            SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
            Set SECOTEMP = READSET(SQLX$)
            With SECOTEMP
              Do While Not .EOF
                No% = !NUMEOPER
                '
                CAREALI# = 0
                If REPIM% = 2 Then
                  CAREALI# = CANREALI(NORFA$, No%)
                End If
                CAPENDE# = CAPROYE# - CAREALI#
                If CAPENDE# < 0 Then CAPENDE# = 0
                '
                ATO% = 0
                'AQUI CORREGI POR QUE SI AltOper ERA NULL DABA ERROR
                If Not IsNull(!AltOper) Then ATO% = !AltOper
                If ATO% < 1 Then   ' SOLO OPERACIONES BASICAS
                  DOPE$ = SAFETEXT$(!DESCOPER)
                  EQ1$ = SAFETEXT$(!CodMaq0) 'PARA EVITAR ERROR 91 NULL
                  'EQAUX$ = AJUSTI$(!CodMaq1, 6) + AJUSTI$(!CodMaq2, 6) + AJUSTI$(!CodMaq1, 6)
                  EQAUX$ = ""
                  If Not IsNull(!CodMaq1) Then EQAUX$ = AJUSTI$(!CodMaq1, 6)
                  If Not IsNull(!CodMaq2) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq2, 6)
                  If Not IsNull(!CodMaq1) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq1, 6)
                  LOTESTA = !LOTESTAN: If LOTESTA < 1 Then LOTESTA = 1
                  CALOTES = 1
                  CAOPS = !CantiOps
                    If CAOPS < 0.01 Then CAOPS = 1
                  CICLOHE = !PieCiclo
                    If CICLOHE < 1 Then CICLOHE = 1
                  MINPREP = 60 * !HorsPrep * CALOTES
                    If CAREALI# > 0.05 Then MINPREP = 0
                    On Error GoTo 0
                  MINFIJO = 60 * !HorsFijo * CALOTES
                  MINVARI = CAPENDE# * !MinuStan
                  MINHERR = CAPENDE# * 60 * !HorsMHer / CICLOHE
                  MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
                  MINLIMP = 60 * !HorsLimp * CALOTES
                  MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO + MINLIMP
                  HORAPROD = MINLOTE / 60
                  '
                  BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
                  NUOROP% = 0
                  FEAL% = HOY(HOS$)
                  HOAL% = Int(HORANUM(Time$))
                  USUA% = USNBR% Mod 100
                  DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
                  INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
                  PROGBLA$ = Space$(6) + String$(26, 0)
                  '
                  If Status% <= 1 Then
                    SQLX$ = "SELECT * FROM SIMU_OPERFAB WHERE COD_INTE < 1 "
                    Dim OPERFFFAB As ADODB.Recordset
                    Set OPERFFFAB = New ADODB.Recordset
                    OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    'With OperFab
                    With OPERFFFAB
                      .AddNew   ' YA CONVERTIDO
                      On Error Resume Next
                      !CodiEmpr = CodiEmp%
                      On Error GoTo 0
                        !idsubor = Left$(NORFA$, 12)
                        !COD_INTE = CIJK%
                        !COD_EXTE = Left$(CODEXT$(CIJK%), 16)
                        !DESCRIP = Left$(DESCRIT0$(CIJK%), 64)
                        !NUMEOPER = No%
                        !DESCOPER = Left$(DOPE$, 48)
                        !MAQUINA = Left$(EQ1$, 6)
                        !DescMaq = Left$(ECOARCH("PADMAQ", EQ1$), 64)
                        !EquiAux = Left$(EQAUX$, 18)
                        !CantiOps = CAOPS
                        !CANTPROG = CAPROYE#
                        !CANTREAL = CAREALI#
                        !CANTPEND = CAPENDE#
                        !TiemPrep = MINPREP
                        !TiemFijo = MINFIJO
                        !TiemVari = MINVARI
                        !TiemDemo = MINDEMO
                        !TiemMHer = MINHERR
                        !TiemLimp = MINLIMP
                        !TiemOtro = 0
                        !TiemTota = MINLOTE
                        !TiemEjec = 0
                        !TiemPend = MINLOTE
                        !StatOper = 1
                        !ComiePro = CVDAT(0)
                        !TermiPro = CVDAT(0)
                        !TermiTemp = CVDAT(0)
                        !Observa = " "
                        !MARBORRA = 0
                        !IdSimulOP = SIMULA&
                      .Update
                    End With
                    OPERFFFAB.Close
                    Set OPERFFFAB = Nothing
                  End If
                End If
              .MoveNext
              Loop
            End With
            Set SECOTEMP = Nothing
          Next II%
          
    Next I&
    
    Screen.MousePointer = 1

ERROR_GUARDAR:
    
    Screen.MousePointer = 11
    If Err <> 0 Then
'      FLEXCONN.RollbackTrans
      ORDEFFFABR1.Close
      Set ORDEFFFABR1 = Nothing

      CONDI$ = "IdSimulOR = " & SIMULA&
      CONDI$ = CONDI$ + " AND MARBORRA = 1"
      Call ACTUREGISTRO("SIMU_ORDEFABR", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")
      
      CONDI$ = "IdSimulOP = " & SIMULA& & " AND MARBORRA = 1"
      Call ACTUREGISTRO("SIMU_OPERFAB", "MARBORRA", CONDI$, 0, REGAF&, "NOMESS")

      Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
      Exit Sub
    Else
'      FLEXCONN.CommitTrans
      ORDEFFFABR1.Close
      Set ORDEFFFABR1 = Nothing
      CONDI$ = "IdSimulOR = " & SIMULA& & " AND MARBORRA = 1"
      Call BORRAREGISTROS("SIMU_ORDEFABR", CONDI$, REGAF&, "NOMESS")
      Call CARGALISEL("LISTSIMU", "SIMU_ORDEFABR", "IdSimulOR/F6;Referen/A48", "IdSimulOR > 0 ORDER BY IdSimulOR DESC", "DISTINCT")
      
      CONDI$ = "IdSimulOP = " & SIMULA& & " AND MARBORRA = 1"
      Call BORRAREGISTROS("SIMU_OPERFAB", CONDI$, REGAF&, "NOMESS")
      '
      Call GENEOPERFAP_SIMU
      Call AGRESIMU_SECOPERPEN
      Screen.MousePointer = 1
          
      If Modo% = 0 Then Call COMUNI("Simulación Grabada")
    End If
    
99  Exit Sub

End Sub

Function NuevaSimulacion&()
        CONDI2$ = " IdSimulOR > 0"
        NuevaSimulacion& = 1 + XVALO(VALOBADA("SIMU_ORDEFABR", "MAX(IdSimulOR)", CONDI2$, "NOMESS"))
End Function

Function ProxiNumSimulaOF&()
   CONDI$ = "NumeOrFa > 0  ORDER BY NUMEORFA DESC"
   ProxiNumSimulaOF& = 1 + XVALO(VALOBADA("SIMU_ORDEFABR", "TOP 1 NumeOrFa", CONDI$, "NOMESS"))
End Function

Sub Command1_Click()
   
   Call CARGALISEL("LISTSIMU", "SIMU_ORDEFABR", "IdSimulOR/F6;Origen/A48", "IdSimulOR > 0 ORDER BY IdSimulOR DESC", "DISTINCT")
   
   Call SELECHO("LISTSIMU/LISTA DE SIMULACIONES")
   SIMULA& = XVALO(VALACT1$("LISTSIMU"))
   If SIMULA& < 1 Then Exit Sub
   '
   Text2 = SIMULA&
   
End Sub

Private Sub Command17_Click()
      
      Command17.Enabled = False

      Call GuardaSimulacion(0)
    
      Command17.Enabled = True
      
End Sub



Sub AGRESIMU_SECOPERPEN()
   '
   Call ABREORFAB
   '
   'RECORRER EL OPERFAB TOMAR EL CODIGO DE LA MAQUINA
   SQLX$ = "SELECT * FROM OPERFAB "
   SQLX$ = SQLX$ + " WHERE CANTPEND > 0 "
   
   Dim OPERTMP As ADODB.Recordset
   Set OPERTMP = New ADODB.Recordset
   OPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With OPERTMP
     Do While Not .EOF
       CODMAQ$ = SAFETEXT$(!MAQUINA)
       'TRAIGO EL SECTOR SEGUN LA MAQUINA DE LA TABLA EQUIPOS
       CONDI$ = "CODIGO_EQU = '" & CODMAQ$ & "'"
       SECT = XVALO(VALOBADA("EQUIPOS", "UBICA_EQU", CONDI$, "N0MESS"))
       '
       If SECT > 0 Then ' SI LO ENCUENTRA LO GRABA
          !SecciPro = SECT
          .Update
          GoTo 20
       Else
         SECT = CALCAR07.SECMAQ(CODMAQ$)
         !SecciPro = SECT
         .Update
       End If
       '
20    .MoveNext
     Loop
   End With
   OPERTMP.Close
   Set OPERTMP = Nothing
   
   SQLX$ = "SELECT * FROM SIMU_OPERFAB "
   SQLX$ = SQLX$ + " WHERE CANTPEND > 0 "
   
   Dim SIMU_OPERTMP As ADODB.Recordset
   Set SIMU_OPERTMP = New ADODB.Recordset
   SIMU_OPERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With SIMU_OPERTMP
     Do While Not .EOF
       CODMAQ$ = SAFETEXT$(!MAQUINA)
       'TRAIGO EL SECTOR SEGUN LA MAQUINA DE LA TABLA EQUIPOS
       CONDI$ = "CODIGO_EQU = '" & CODMAQ$ & "'"
       SECT = XVALO(VALOBADA("EQUIPOS", "UBICA_EQU", CONDI$, "N0MESS"))
       '
       If SECT > 0 Then ' SI LO ENCUENTRA LO GRABA
          !SecciPro = SECT
          .Update
          GoTo 30
       Else
         SECT = CALCAR07.SECMAQ(CODMAQ$)
         !SecciPro = SECT
         .Update
       End If
       '
30    .MoveNext
     Loop
   End With
   SIMU_OPERTMP.Close
   Set SIMU_OPERTMP = Nothing

End Sub

Sub GENEOPERFAP_SIMU()
    '
    JJ& = 0
    For U& = 1 To ULTREG&("OPERFAP")
      YY$ = REGLEIDO$("OPERFAP", U&)
      For KKI% = 103 To 131 Step 4
        TITO = 60 * CVS(Mid$(YY$, KKI%, 4))
        Call REPLA(YY$, MKS$(TITO), KKI%, 4)
      Next KKI%
      JJ& = JJ& + 1
      Call GRAREG("OPERFAH", YY$, JJ&)
    Next U&
    Call SETULTREG("OPERFAH", JJ&)
    Call CIERRARCH("OPERFAP")
    Call CIERRARCH("OPERFAH")
    '
    Call COPYSTRU("OPERFAP", "SOPERFAP")
    Call TRALOCAL("OPERFAP", "SOPERFAP")
    Call TRALOCAL("SOPERFAP", "")
    '
    Exit Sub

    Call COPYSTRU("OPERFAP", "SOPERFAP")
    
' TIENE QUE QUEDAR ASI !!! - epb - 12/01/05


'Exit Sub 'LE SAQUE EL COMENTARIO PARA GENERAR
          'OPERFAP.DAT Y OPERFAH.DAT
          'WAR 20/12/04
    '
    SCMA = Screen.MousePointer
    Call ABREORFAB
    '
    Screen.MousePointer = 11
    SQLX$ = " SELECT OPERFAB.*, ORDEFABR.Destino FROM OPERFAB "
    SQLX$ = SQLX$ + " INNER JOIN ORDEFABR "
    SQLX$ = SQLX$ + " ON OPERFAB.IdSubOr = ORDEFABR.IdSubOr "
    SQLX$ = SQLX$ + " WHERE StatOper = 0 OR StatOper = 1 "
    SQLX$ = SQLX$ + " AND CantReal < CantProg "
    SQLX$ = SQLX$ + " ORDER BY OPERFAB.IdSubOr "
    '
    '
    Dim OPERFAPTMP As ADODB.Recordset
    Set OPERFAPTMP = New ADODB.Recordset
    OPERFAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With OPERFAPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Operaciones Pendientes.")
        On Error GoTo 0
        GoTo 99
      End If
      On Error GoTo 0
      '
      JJ& = 0
      Do While (.EOF = False)
        CAPROG# = XVALO(!CANTPROG)
        CAREA# = XVALO(!CANTREAL)
        CAPEN# = CAPROG# - CAREA#
        CIXII% = XVALO(!COD_INTE)
        NOPEII% = XVALO(!NUMEOPER)
        '
        If CAPEN# >= 0.05 Then
          If CAPROG# >= 0.05 Then
            COEF = CAPEN# / CAPROG#
            If COEF > 1 Then COEF = 1
            TPREP = XVALO(!TiemPrep): If CAREA# >= 0.05 Then TPREP = 0
            TFIJO = XVALO(!TiemFijo)
            'TVARI = TIESTAN(CIXII%, NOPEII%, CAPEN#, 1) 'COEF * CVS(Mid$(O$, 111, 4))
            TPROD = TIEMAQUI(CIXII%, NOPEII%, CAPEN#, 1) ' MODO 1 PARA NO INCLUIR PREPARACION
            TVARI = TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
            TDEMO = 0: TMANT = 0
              If (XVALO(!TiemFijo) + XVALO(!TiemVari)) > 0.1 Then
                TDEMO = TVARI * XVALO(!TiemDemo) / (XVALO(!TiemFijo) + XVALO(!TiemVari))
              End If
            TMANT = COEF * CVS(Mid$(O$, 119, 4))
            TLIMP = CVS(Mid$(O$, 123, 4))
            TOTRO = COEF * CVS(Mid$(O$, 127, 4))
            TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
            '
            YY$ = REGBLAN("OPERFAP")
            On Error Resume Next
            Call REPLA(YY$, SAFETEXT$(!idsubor), 1, 12)
            Call REPLA(YY$, MKI$(XVALO(!COD_INTE)), 13, 2)
            Call REPLA(YY$, MKI$(XVALO(!NUMEOPER)), 15, 2)
            Call REPLA(YY$, SAFETEXT$(!DESCOPER), 17, 48)
            Call REPLA(YY$, SAFETEXT$(!DESTINO), 65, 8)
            Call REPLA(YY$, SAFETEXT$(!MAQUINA), 73, 6)
            Call REPLA(YY$, MKD$(XVALO(!CANTPROG)), 79, 8)
            Call REPLA(YY$, MKD$(XVALO(!CANTREAL)), 87, 8)
            Call REPLA(YY$, MKD$(XVALO(!CANTPEND)), 95, 8)
            Call REPLA(YY$, MKS$(TPREP), 103, 4)
            Call REPLA(YY$, MKS$(TFIJO), 107, 4)
            Call REPLA(YY$, MKS$(TVARI), 111, 4)
            Call REPLA(YY$, MKS$(TDEMO), 115, 4)
            Call REPLA(YY$, MKS$(TMANT), 119, 4)
            Call REPLA(YY$, MKS$(TLIMP), 123, 4)
            Call REPLA(YY$, MKS$(TOTRO), 127, 4)
            Call REPLA(YY$, MKS$(TTOTA), 131, 4)
                HORACO$ = Format(!ComiePro, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!ComiePro)), 135, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORACO$))), 137, 2)
                HORATE$ = Format(!TermiPro, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!TermiPro)), 139, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORATE$))), 141, 2)
                HORATEMP$ = Format(!TermiTemp, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!TermiTemp)), 143, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORATEMP$))), 145, 2)
            Call REPLA(YY$, SAFETEXT$(!Observa), 151, 48)
            On Error GoTo 0
            '
            JJ& = JJ& + 1
            Call GRAREG("SOPERFAP", YY$, JJ&)
            For KKI% = 103 To 131 Step 4
              TITO = 60 * CVS(Mid$(YY$, KKI%, 4))
              Call REPLA(YY$, MKS$(TITO), KKI%, 4)
            Next KKI%
            Call GRAREG("OPERFAH", YY$, JJ&)
            '
          End If
        End If
      .MoveNext
      Loop
    End With
    Set OPERFAPTMP = Nothing
    '
    CONSIMULACION% = XVALO(Text2)
    
    SQLX$ = " SELECT SIMU_OPERFAB.*, SIMU_ORDEFABR.Destino FROM SIMU_OPERFAB "
    SQLX$ = SQLX$ + " INNER JOIN SIMU_ORDEFABR "
    SQLX$ = SQLX$ + " ON SIMU_OPERFAB.IdSubOr = SIMU_ORDEFABR.IdSubOr "
    SQLX$ = SQLX$ + " WHERE StatOper = 0 OR StatOper = 1 "
    SQLX$ = SQLX$ + " AND CantReal < CantProg "
    SQLX$ = SQLX$ + " AND IDSIMULOP = " & CONSIMULACION%
    SQLX$ = SQLX$ + " ORDER BY SIMU_OPERFAB.IdSubOr "
    '
    Set OPERFAPTMP = New ADODB.Recordset
    OPERFAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With OPERFAPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Operaciones Pendientes.")
        On Error GoTo 0
        GoTo 99
      End If
      On Error GoTo 0
      '
      Do While (.EOF = False)
        CAPROG# = XVALO(!CANTPROG)
        CAREA# = XVALO(!CANTREAL)
        CAPEN# = CAPROG# - CAREA#
        CIXII% = XVALO(!COD_INTE)
        NOPEII% = XVALO(!NUMEOPER)
        '
        If CAPEN# >= 0.05 Then
          If CAPROG# >= 0.05 Then
            COEF = CAPEN# / CAPROG#
            If COEF > 1 Then COEF = 1
            TPREP = XVALO(!TiemPrep): If CAREA# >= 0.05 Then TPREP = 0
            TFIJO = XVALO(!TiemFijo)
            'TVARI = TIESTAN(CIXII%, NOPEII%, CAPEN#, 1) 'COEF * CVS(Mid$(O$, 111, 4))
            TPROD = TIEMAQUI(CIXII%, NOPEII%, CAPEN#, 1) ' MODO 1 PARA NO INCLUIR PREPARACION
            TVARI = TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
            TDEMO = 0: TMANT = 0
              If (XVALO(!TiemFijo) + XVALO(!TiemVari)) > 0.1 Then
                TDEMO = TVARI * XVALO(!TiemDemo) / (XVALO(!TiemFijo) + XVALO(!TiemVari))
              End If
            TMANT = COEF * CVS(Mid$(O$, 119, 4))
            TLIMP = CVS(Mid$(O$, 123, 4))
            TOTRO = COEF * CVS(Mid$(O$, 127, 4))
            TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
            '
            YY$ = REGBLAN("SOPERFAP")
            On Error Resume Next
            Call REPLA(YY$, SAFETEXT$(!idsubor), 1, 12)
            Call REPLA(YY$, MKI$(XVALO(!COD_INTE)), 13, 2)
            Call REPLA(YY$, MKI$(XVALO(!NUMEOPER)), 15, 2)
            Call REPLA(YY$, SAFETEXT$(!DESCOPER), 17, 48)
            Call REPLA(YY$, SAFETEXT$(!DESTINO), 65, 8)
            Call REPLA(YY$, SAFETEXT$(!MAQUINA), 73, 6)
            Call REPLA(YY$, MKD$(XVALO(!CANTPROG)), 79, 8)
            Call REPLA(YY$, MKD$(XVALO(!CANTREAL)), 87, 8)
            Call REPLA(YY$, MKD$(XVALO(!CANTPEND)), 95, 8)
            Call REPLA(YY$, MKS$(TPREP), 103, 4)
            Call REPLA(YY$, MKS$(TFIJO), 107, 4)
            Call REPLA(YY$, MKS$(TVARI), 111, 4)
            Call REPLA(YY$, MKS$(TDEMO), 115, 4)
            Call REPLA(YY$, MKS$(TMANT), 119, 4)
            Call REPLA(YY$, MKS$(TLIMP), 123, 4)
            Call REPLA(YY$, MKS$(TOTRO), 127, 4)
            Call REPLA(YY$, MKS$(TTOTA), 131, 4)
                HORACO$ = Format(!ComiePro, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!ComiePro)), 135, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORACO$))), 137, 2)
                HORATE$ = Format(!TermiPro, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!TermiPro)), 139, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORATE$))), 141, 2)
                HORATEMP$ = Format(!TermiTemp, "hh:mm:ss")
            Call REPLA(YY$, MKI$(CVINT(!TermiTemp)), 143, 2)
            Call REPLA(YY$, MKI$(Int(HORANUM(HORATEMP$))), 145, 2)
            Call REPLA(YY$, SAFETEXT$(!Observa), 151, 48)
            On Error GoTo 0
            '
            JJ& = JJ& + 1
            Call GRAREG("SOPERFAP", YY$, JJ&)
            For KKI% = 103 To 131 Step 4
              TITO = 60 * CVS(Mid$(YY$, KKI%, 4))
              Call REPLA(YY$, MKS$(TITO), KKI%, 4)
            Next KKI%
            Call GRAREG("OPERFAH", YY$, JJ&)
            '
          End If
        End If
      .MoveNext
      Loop
    End With
    Set OPERFAPTMP = Nothing
  
    Call SETULTREG("SOPERFAP", JJ&)
    Call SETULTREG("OPERFAH", JJ&)
    Call FILESORT("SOPERFAP", "", 1, 4, "ASC")
    Call CIERRARCH("SOPERFAP")
    Call FILESORT("OPERFAH", "", 1, 4, "ASC")
    Call CIERRARCH("OPERFAH")
    '
99  Screen.MousePointer = SCMA
    '
End Sub


Private Sub Command2_Click()
    'GRABA EN LA TABLA DE SIMULACION
    Command2.Enabled = False
    Screen.MousePointer = 11
    If MSF0.Rows <= 1 Then
         Call COMUNI("Falta Cargar Detalle de Simulación")
         GoTo 99
    End If

    If TRIM$(TXTREFERENCIA) = "" Then
         Call COMUNI("Falta Ingresar Descripción de la Simulación")
         GoTo 99
    End If
    
    Call BLOQUEAR_CONTROLES(Me, 2, 1)

    Call Command17_Click
    Call BLOQUEAR_CONTROLES(Me, 2, 1)
    'ACA DEBE RECALCULAR
    CALCAR07.CONSIMULACION% = XVALO(Text2)
    Call GRACONTROL("CARMAQUI", "ULTISIMU", Text2.TEXT)
    Call CALCAR07.Command2_Click
    FGANTT07.mnuincluyesimulacion.Caption = FGANTT07.mnuincluyesimulacion.Caption & "(" & XVALO(Text2) & ")"

    CALCAR07.Show 1
    'GRABO EL NUMERO DE LA ULTIMA SIMULACION
    '
99  Call BLOQUEAR_CONTROLES(Me, 1, 1)
    Command2.Enabled = True
    Screen.MousePointer = 1
End Sub
Private Sub Command28_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text17.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command3_Click()
   Call BLANFORMU
   Text2 = NuevaSimulacion&
End Sub

Private Sub Command4_Click()

    Call GuardaSimulacion(1)
    
    For II& = 1 To MSF0.Rows - 1
        CodigoInterno% = MSF0.TextMatrix(II&, 1)
        Cantidad# = XVALO(MSF0.TextMatrix(II&, 4))
        FechaSolicitada% = CVINT(MSF0.TextMatrix(II&, 5))
        Referencia$ = TRIM$(TXTCOMPRA)
        DestinoFabricacion$ = Text3
        
        ValorDevuelto% = COMALTER%("O-Fabricación para '" & DESCRIT0$(CodigoInterno%) & "'\\Generar OF\Obviar")
        If ValorDevuelto% = 1 Then
            ACONEC$ = "ORFABR07/ORFABNUE" & Chr$(124) & "CARMAQ" + Chr$(124) + CStr(CodigoInterno%) + Chr$(124) + TRIM$(FORMATNUM$(Cantidad#, "F12.1")) + Chr$(124) + CStr(FechaSolicitada%) + Chr$(124) + Referencia$ + Chr$(124) + DestinoFabricacion$ + Chr$(124)
            Call CONECRUN(ACONEC$, "Ordenes de Fabricación")
        End If
    Next II&

    ValorDevuelto% = COMALTER%("¿Desea Eliminar la Simulación?\\No, Mantener\Si, Eliminar")
    If ValorDevuelto% = 2 Then
        NroSimulacion& = XVALO(Text2)
        CondicionSql$ = "IdSimulOR = " & CStr(NroSimulacion&)
        Call BORRAREGISTROS("SIMU_ORDEFABR", CondicionSql$, REGAF&)
        Call BLANFORMU
    End If
    
End Sub

Private Sub Command5_Click()

    Call SELECHO("DESTFAB/DESTINOS DE FABRICACION")
    DestinoFabricacion$ = TRIM$(VALACT1$("DESTFAB"))
    Text3 = DestinoFabricacion$
    
End Sub


Private Sub Form_Load()
  CAMPOS$ = "CI/F0;Código/A16;Descripción/A48;Cant.Sol./F15.1;Sol.Entr./D8"
  Call CARGA_ENCABEZADO(Me.MSF0, CAMPOS$, Me, 0, 0)
  Text1 = CONTROL("CARMAQUI", "ULTISIMU")
  

  '
  Call CREATABLA_SIMU_ORDEFABR_SIMU_OPERFAB
End Sub
Sub GENORDEFDAT_SIMU(Modo%, Optional SECSEL)
  '
  SECSEL1 = XVALO(SECSEL)
  Call ABREORFAB
  '
  Call COPYSTRU("ORDEFABR", "SRDEFABR")
  Call COPYSTRU("ORDEFABP", "SRDEFABP")
  
  SQLX$ = " SELECT  NumeOrFa, REFEREN, IDSUBOR, FECHGEN, COD_INTE, Unidad, CANPROY, DESTINO, FECHENSOL, Priorid,"
  SQLX$ = SQLX$ + " StatuOrf, FechLan, SecciPro, Fechcierr, CANAPRO, CanRech, CanRech, IdenRetra, FechProyEntre,"
  SQLX$ = SQLX$ + " FechTempEntre, FechRecal, OBSERVA, ID_ORFA"
  SQLX$ = SQLX$ + " FROM ORDEFABR "
  If Modo% = 1 Then SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
  If CONSIMULACION% > 0 Then
        SQLX$ = SQLX$ + " UNION ALL"
        SQLX$ = SQLX$ + " SELECT  NumeOrFa, REFEREN, IDSUBOR, FECHGEN, COD_INTE, Unidad, CANPROY, DESTINO, FECHENSOL, Priorid,"
        SQLX$ = SQLX$ + " StatuOrf, FechLan, SecciPro, Fechcierr, CANAPRO, CanRech, CanRech, IdenRetra, FechProyEntre,"
        SQLX$ = SQLX$ + " FechTempEntre, FechRecal, OBSERVA, ID_ORFA"
        SQLX$ = SQLX$ + " FROM SIMU_ORDEFABR "
        'SQLX$ = SQLX$ + " WHERE IdSimulOR = " & CONSIMULACION%
        SQLX$ = SQLX$ + " WHERE STATUORF = 1"   ' EL ESTADO 1 SIGNIFICA QUE FUE LA SIMULACION INCLUIDA EN EL ULTIMO CALCULO
  End If
  SQLX$ = SQLX$ + " ORDER BY IdSubOr "
  '
  'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim ORDEFABTMP As ADODB.Recordset
  Set ORDEFABTMP = New ADODB.Recordset
  ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  '
  With ORDEFABTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
        On Error GoTo 0
        Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.\Imposible continuar la carga")
        Exit Sub
    End If
    On Error GoTo 0
    FIN& = .RecordCount
    JJ& = 0: J2& = 0: J3& = 0
    ASELE$ = "SRDEFABR":    ASEL2$ = "SRDEFABP":
    If SECSEL1 > 0 Then
       ASELE$ = "!" + ASELE$
       ASEL2$ = "!" + ASEL2$
    End If
    Do While (.EOF = False)
      J3& = J3& + 1
      Call TRACE(20, J3&, FIN&)
      If SECSEL1 > 0 Then
         If Abs(XVALO(!SecciPro) - SECSEL1) > 0.0005 Then
           GoTo 19
         End If
      End If

      Y$ = REGBLAN("SRDEFABR")
      '
      On Error Resume Next
      
      If InStr(1, !idsubor, "SF") > 0 Then
      A = B
      End If
      Call REPLA(Y$, MKS$(XVALO(!NumeOrFa)), 1, 4)
      Call REPLA(Y$, SAFETEXT$(!REFEREN), 5, 64)
      Call REPLA(Y$, SAFETEXT$(!idsubor), 69, 12)
      Call REPLA(Y$, MKI$(CVINT(!FECHGEN)), 81, 2)
      Call REPLA(Y$, MKI$(XVALO(!COD_INTE)), 83, 2)
      Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
      Call REPLA(Y$, MKD$(XVALO(!CANPROY)), 89, 8)
      '
      Call REPLA(Y$, SAFETEXT$(!DESTINO), 97, 8)
      '
      Call REPLA(Y$, MKI$(CVINT(!FechEnSol)), 105, 2)
      Call REPLA(Y$, Chr$(XVALO(!Priorid)), 107, 1)
      Call REPLA(Y$, Chr$(XVALO(!StatuOrf)), 108, 1)
      Call REPLA(Y$, MKI$(CVINT(!FechLan)), 109, 2)
      Call REPLA(Y$, MKS$(XVALO(!SecciPro)), 113, 4)
      Call REPLA(Y$, MKI$(CVINT(!Fechcierr)), 117, 2)
      Call REPLA(Y$, MKD$(XVALO(!CANAPRO)), 119, 8)
      Call REPLA(Y$, MKD$(XVALO(!CanRech)), 127, 8)
      Call REPLA(Y$, SAFETEXT$(!IdenRetra), 135, 8)
      Call REPLA(Y$, MKI$(CVINT(!FechProyEntre)), 147, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechTempEntre)), 149, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechRecal)), 151, 2)
      Call REPLA(Y$, SAFETEXT$(!Observa), 155, 64)
        FechaSolicitada% = CVINT(!FechEnSol)
        FechaCalculo% = CVINT(!FechProyEntre)
        If FechaCalculo% >= FechaSolicitada% Then
            DifeDias& = DIENTRE%(FechaSolicitada%, FechaCalculo%)
        Else
            DifeDias& = (-1) * DIENTRE%(FechaCalculo%, FechaSolicitada%)
        End If
      Call REPLA(Y$, MKS$(DifeDias&), 189, 4)
      Call REPLA(Y$, SAFETEXT$(!ID_ORFA), 193, 48)
      On Error GoTo 0
      '
      JJ& = JJ& + 1
      Call GRAREG(ASELE$, Y$, JJ&)
      Call TRACE(20, JJ&, FIN&)
      '
      If XVALO(!StatuOrf) < 2 Then
        J2& = J2& + 1
        Call GRAREG(ASEL2$, Y$, J2&)
      End If
19  .MoveNext
    Loop
  End With
  Set ORDEFABTMP = Nothing
  Call SETULTREG(ASELE$, JJ&)
  Call CIERRARCH(ASELE$)
  Call SETULTREG(ASEL2$, J2&)
  Call CIERRARCH(ASEL2$)
  
  Call FILESORT(ASELE$, "", 10, 10)
  Call FILESORT(ASEL2$, "", 10, 10)
End Sub


Sub CREATABLA_SIMU_ORDEFABR_SIMU_OPERFAB()

   Call ABRECONEXION
   
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "NumeOper", 3, 2)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "EquiAux", 10, 18)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "CantiOps", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "CantProg", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "CantPend", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemPrep", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemFijo", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemVari", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemDemo", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemMHer", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemLimp", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemOtro", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemTota", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TiemPend", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "StatOper", 3, 2)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "ComiePro", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TermiPro", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "TermiTemp", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "Observa", 10, 48)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "HoBatchRec", 11, 0)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "SecciPro", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "IdSimulOP", 4, 4, 1)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "DesSimulOP", 10, 48)
   Call CREACAMPO("OFABRIC", "SIMU_OPERFAB", "MARBORRA", 3, 0)
   '
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "NumeOrFa", 4, 4, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechGen", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Unidad", 10, 32)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "CanProy", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Destino", 10, 32)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Priorid", 3, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "StatuOrf", 3, 1, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "NuOrProd", 10, 16)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechLan", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "SecciPro", 6, 4)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechCierr", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "CanApro", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "CanRech", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "IdenRetra", 10, 12)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechProyEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechTempEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "FechRecal", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Observa", 12, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Origen", 10, 64)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "AnotaOrf", 12, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "NumRem_Prov", 10, 16)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Nume_Ocom", 10, 16)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Nume_Cli", 4, 4)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Id_Orfa", 10, 32)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "NroPed", 4, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Armador ", 3, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "LoteIndivis", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "inicio_informado", 8, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Canti_fin_proceso", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Status_Proceso", 3, 0)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "Canti_Iniciada", 7, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "NumeroChasis", 10, 24)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "TipoChasis", 10, 24)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "LineaInterno", 10, 8)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "IdSimulOR", 4, 4, 1)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "DesSimulOR", 10, 48)
   Call CREACAMPO("OFABRIC", "SIMU_ORDEFABR", "MARBORRA", 3, 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call SETULTREG("!CARDETPE", 0)
   Call CIERRARCH("*.*")
End Sub


Private Sub MSF0_DblClick()
   Call BLANARCH("!CARDETPE")
   For U& = 1 To MSF0.Rows - 1
     CI1% = XVALO(MSF0.TextMatrix(U&, 1))
     CANTPEDIDA = XVALO(MSF0.TextMatrix(U&, 4))
     FEENT% = FECHANUM(MSF0.TextMatrix(U&, 5))
     Call REPLA(Y$, MKI$(CI1%), 1, 2)
     Call REPLA(Y$, MKS$(CANTPEDIDA), 47, 4)
     Call REPLA(Y$, MKI$(FEENT%), 63, 2)
     Call GRAREG("!CARDETPE", Y$, U&)
     '
   Next U&

10 XXX% = VENTABU%("COCAFECH")
   '
   If XXX% < 0 Then
      Call CIERRARCH("!CARDETPE")
      Call BLANARCH("!CARDETPE")
      On Error Resume Next
      On Error GoTo 0
      Exit Sub
   End If
   '

20 If ULTREG&("!CARDETPE") < 1 Then
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   Screen.MousePointer = 11   ' Registrando Pedido ...
   FIN& = ULTREG&("!CARDETPE")
   '
   FENTREGEN% = FECHANUM(Text17)
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI1% = CVI(Left$(Y$, 2))
     If CI1% < 1 Then
       Call MENSERR(24, "Código no Válido")
       Call REPLA(Y$, MKI$(0), 1, 2)
       Call GRAREG("!CARDETPE", Y$, U&)
       GoTo 10
     End If
     '
     FENTRESOL% = CVI(Mid$(Y$, 63, 2))
     If FENTRESOL% < FENTREGEN% Then
        FENTRESOL% = FENTREGEN%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     If FENTRESOL% < HOII% Then
        FENTRESOL% = HOII%
        Call REPLA(Y$, MKI$(FENTRESOL%), 63, 2)
        Call GRAREG("!CARDETPE", Y$, U&)
     End If
     '
     
     '
     X$ = REGLEIDO$("!CARDETPE", U&)
     CI1% = CVI(Left$(X$, 2))
     TTXX$ = Space$(96)
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     '
     CAN = CVS(Mid$(X$, 47, 4))
     '
     If CAN < 0.5 Then
          Call MENSERR(24, "Falta Cantidad")
          Call REPLA(X$, MKS$(0), 47, 4)
          Call GRAREG("!CARDETP", X$, U&)
          GoTo 10
     End If
     '
     CANTPEDIDA = CVS(Mid$(X$, 47, 4))
     '
     Call GRAREG("!CARDETPE", X$, U&)
     '
   Next U&
   '
   MSF0.Rows = 1
   For U& = 1 To FIN&
     Y$ = REGLEIDO$("!CARDETPE", U&)
     CI1% = CVI(Left$(Y$, 2))
     '
     X$ = REGLEIDO$("!CARDETPE", U&)
     CI1% = CVI(Left$(X$, 2))
     FENTRESOL% = CVI(Mid$(X$, 63, 2))
     CANTPEDIDA = CVS(Mid$(X$, 47, 4))
     '
     MSF0.Rows = U& + 1
     MSF0.TextMatrix(U&, 1) = CI1%
     MSF0.TextMatrix(U&, 2) = CODEXT$(CI1%)
     MSF0.TextMatrix(U&, 3) = DESCRIT$(CI1%)
     MSF0.TextMatrix(U&, 4) = FORMATNUM$(CANTPEDIDA, "F10.1")
     MSF0.TextMatrix(U&, 5) = FECHATEX$(FENTRESOL%)
   Next U&

   Screen.MousePointer = 1
   '
End Sub

Private Sub Text17_DblClick()
   Text17 = ""
End Sub

Private Sub Text2_Change()
    SIMULA& = XVALO(Text2)
    
    If SIMULA& = 0 Then
        TXTREFERENCIA = "": MSF0.Rows = 1
        Exit Sub
    End If
    
    CONDI2$ = " IdSimulOR =" & SIMULA&
    TXTREFERENCIA = VALOBADA("SIMU_ORDEFABR", "ORIGEN", CONDI2$, "NOMESS")
    TXTCOMPRA = VALOBADA("SIMU_ORDEFABR", "ID_ORFA", CONDI2$, "NOMESS")
    '
    MSF0.Rows = 1: j& = 0
    SQLX$ = "SELECT * FROM SIMU_ORDEFABR WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IdSimulOR = " & SIMULA&
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
          CI1% = XVALO(!COD_INTE)
          NORSIMU$ = SAFETEXT$(!idsubor)
          FEGEN% = CVINT(!FECHGEN)
          CODI$ = SAFETEXT$(!COD_EXTE)
          DESCRIX$ = SAFETEXT$(!DESCRIP)
          FEENT% = CVINT(!FechEnSol)
          CANTI = XVALO(!CANPROY)
          DESTI$ = SAFETEXT$(!DESTINO)
          j& = j& + 1
          MSF0.Rows = j& + 1
          MSF0.TextMatrix(j&, 1) = CI1%
          MSF0.TextMatrix(j&, 2) = CODI$
          MSF0.TextMatrix(j&, 3) = DESCRIX$
          MSF0.TextMatrix(j&, 4) = FORMATNUM$(CANTI, "F10.1")
          MSF0.TextMatrix(j&, 5) = FECHATEX$(FEENT%)
          Text3 = DESTI$
          '
          .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    '
End Sub


Private Sub Text2_DblClick()
    Call Command1_Click
End Sub

Private Sub Text3_Change()

    Text4 = TRIM$(ECOARCH$("DESTFAB", Text3))
    
End Sub

Private Sub TXTREFERENCIA_GotFocus()
    
    If TRIM$(Text2) = "" Then Command1.SetFocus
    
End Sub
