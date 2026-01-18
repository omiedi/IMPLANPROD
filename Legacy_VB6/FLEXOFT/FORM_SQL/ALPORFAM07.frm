VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ALPORFAM07 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Código de Artículo por Familia"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTrazable 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Trazable"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtCodigoSecundario 
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
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   2730
   End
   Begin MSFlexGridLib.MSFlexGrid mfgAltaPorFamilia 
      Height          =   855
      Left            =   240
      TabIndex        =   21
      Top             =   5640
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1508
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Armado de Código por Grupos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4680
      TabIndex        =   20
      Top             =   1150
      Width           =   3135
   End
   Begin VB.TextBox txtCodigo 
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
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1080
      Width           =   2730
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   -360
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   8040
      ScaleHeight     =   5685
      ScaleWidth      =   1050
      TabIndex        =   13
      Top             =   0
      Width           =   1080
      Begin VB.CommandButton BOTOK 
         Caption         =   "Ok"
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
         Picture         =   "ALPORFAM07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Alta del Nuevo Código"
         Top             =   840
         Width           =   640
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "ALPORFAM07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   100
         Width           =   640
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1800
      MaxLength       =   64
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   6090
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CODIGOS EXISTENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   7680
      Begin MSFlexGridLib.MSFlexGrid mfgCodigosFamilia 
         Height          =   2295
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4048
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "."
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
      Left            =   3360
      Picture         =   "ALPORFAM07.frx":0454
      TabIndex        =   1
      Top             =   210
      Width           =   175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ALPORFAM07.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   -120
      OleObjectBlob   =   "ALPORFAM07.frx":0992
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Código secundario"
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
      Left            =   -720
      TabIndex        =   24
      Top             =   1540
      Width           =   2430
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Código primario"
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
      Left            =   -720
      TabIndex        =   19
      Top             =   1180
      Width           =   2430
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Denominación"
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
      Left            =   -720
      TabIndex        =   12
      Top             =   1800
      Width           =   2430
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   3885
      TabIndex        =   9
      Top             =   670
      Width           =   540
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   4515
      TabIndex        =   8
      Top             =   585
      Width           =   435
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   4515
      TabIndex        =   7
      Top             =   240
      Width           =   435
   End
   Begin VB.Label Label5 
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
      Left            =   5040
      TabIndex        =   6
      Top             =   585
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "U.Medida"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   315
      Width           =   705
   End
   Begin VB.Label Label3 
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
      Left            =   5040
      TabIndex        =   4
      Top             =   210
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEDD&
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Familia"
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
      Left            =   315
      TabIndex        =   3
      Top             =   315
      Width           =   2430
   End
   Begin VB.Label Label1 
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
      Left            =   210
      TabIndex        =   2
      Top             =   585
      Width           =   3225
   End
End
Attribute VB_Name = "ALPORFAM07"
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
    ' VERIFICA ALTA POR TABLA DE FAMILIAS
    If TRIM$(TXTCODIGO.TEXT <> "") Then
        If Len(TXTCODIGO.TEXT) <= 15 Then
            CODINUE$ = TXTCODIGO.TEXT
            GoTo 50
        End If
    End If
    
    Screen.MousePointer = 11
    PRONULI& = 1
    ABUS$ = UCase$(TRIM$(Text1)) + "-"
    FIN& = NUMAS%
    Call ABREMASTER
    SQLX$ = "SELECT * FROM MASTER WHERE CODIGO LIKE '" & ABUS$ & "%'"
    'Set ARTLITMP = Articulos.OpenRecordset(SQLX$, 4)
    Dim ARTLITMP As ADODB.Recordset
    Set ARTLITMP = New ADODB.Recordset
    ARTLITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ARTLITMP
        Do Until .EOF = True
            CODILEIDO$ = UCase$(SAFETEXT$(!Codigo))

            NUITN& = XVALO(Mid$(CODILEIDO$, 4))
            If NUITN& >= PRONULI& Then PRONULI& = 1 + NUITN&
            
19         .MoveNext
        Loop
    End With
    ARTLITMP.Close
    Set ARTLITMP = Nothing
    '
    PRONUX$ = TRIM$(Str$(PRONULI&))
    While Len(PRONUX$) < 5
      PRONUX$ = "0" + PRONUX$
    Wend
    CODINUE$ = ABUS$ + PRONUX$
    '
50  DESCRINUE$ = Text2
    TIMA% = XVALO(Label7)
    UNID$ = TRIM$(Label6)
    NPLANO$ = txtCodigoSecundario.TEXT
    TRAZABLE% = chkTrazable.Value
    Call AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%, TRAZABLE%, NPLANO$)
    '/////////***********/////////////////
    If OKX% < 1 Then
      Call MENSERR(24, "El Alta de Codigo\no pudo Completarse.\  \Ejecute Control de Maestro\y Consulte a su Soporte Técnico.")
    Else
      AMASTOK07.Text1 = CODINUE$
      Call Text1_Change
      Text2 = ""
      Unload Me
    End If
    '
99  Screen.MousePointer = 1
    BOTOK.Enabled = True
    '
End Sub

Private Sub Command1_Click()

'    Dim ta As New DatosAuxiliares
'    Dim rst As ADODB.Recordset
'    Dim Campos(1 To 14) As String
'
'    If TRIM$(Text1) = "" Then
'        Call COMUNI("Falta Elegir una Familia.")
'        Exit Sub
'    End If
'
'    Unload VENTABNEW
'
'    VENTANA.Inicializar = 0
'    VENTANA.AgregaRegistro = 0
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaBorrar = 0
'    VENTANA.HabilitaInsertar = 0
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.TITULO = "ALTA POR FAMILIA"
'
'    idTabla% = ta.ObtenerIdTablaAuxiliar("TAFORALFA")
'    ConsultaSql$ = "select * from datasql with(nolock) where id_lista=" & CStr(idTabla%) & " and campo3='" & Text1 & "'"
'    Set rst = READSET(ConsultaSql$)
'    With rst
'        If .EOF And .BOF Then
'            Call COMUNI("Falta Configurar Tabla de Formato de Familia.")
'            Exit Sub
'        Else
'            .MoveFirst
'            Campos$(4) = SAFETEXT$(!CAMPO4)
'            Campos$(5) = SAFETEXT$(!CAMPO5)
'            Campos$(6) = SAFETEXT$(!CAMPO6)
'            Campos$(7) = SAFETEXT$(!CAMPO7)
'            Campos$(8) = SAFETEXT$(!CAMPO8)
'            Campos$(9) = SAFETEXT$(!CAMPO9)
'            Campos$(10) = SAFETEXT$(!CAMPO10)
'            Campos$(11) = SAFETEXT$(!CAMPO11)
'            Campos$(12) = SAFETEXT$(!CAMPO12)
'            Campos$(13) = SAFETEXT$(!CAMPO13)
'            Campos$(14) = SAFETEXT$(!CAMPO14)
'        End If
'    End With
'    On Error Resume Next
'    rst.Close
'    Set rst = Nothing
'    On Error GoTo 0
'
'    mfgAltaPorFamilia.Rows = 2
'    mfgAltaPorFamilia.Cols = 12
'    VENTANA.CampEditables = ""
'    For i = 4 To 14
'        nroColumna% = i - 3
'        idTabla% = ta.ObtenerIdTablaAuxiliar(Campos$(i))
'        If idTabla% > 0 Then                                                            ' DATO
'            VENTANA.CampEditables = VENTANA.CampEditables & CStr(nroColumna%) & ";"
'            TituloGrupo$ = VALOBADA("encaprfun", "descripcion", "id_lista=" & CStr(idTabla%))
'            VENTANA.ColEcoCampo(nroColumna%) = "campo2;campo3|NO_REPLICA_ECO;datasql;id_lista=" & CStr(idTabla%) & ";order by campo2"
'        Else
'            VENTANA.ColEcoCampo(nroColumna%) = ""
'            mfgAltaPorFamilia.TextMatrix(1, nroColumna%) = Campos$(i)                   ' SEPARADOR
'            TituloGrupo$ = Campos$(i)
'        End If
'        Encabezado$ = Encabezado$ & TituloGrupo$ & "/A12;"
'    Next i
'
'    VENTANA.Campox = Encabezado$
'    VENTANA.Inicializar = 1
'
'    Call EDITARFILA(mfgAltaPorFamilia, VENTABNEW.MSF, 1)
'
'10  With VENTABNEW
'        .Show 1
'        If .APROBADO% > 0 Then
'            txtCodigo.TEXT = ""
'            For i = 1 To 11
'                mfgAltaPorFamilia.TextMatrix(1, i) = .MSF.TextMatrix(1, i)
'                If InStr(1, Campos$(i + 3), "SPACE") > 0 Then
'                    CantEspacios% = XVALO(Mid$(Campos$(i + 3), 7))
'                    txtCodigo.TEXT = txtCodigo.TEXT & Space(CantEspacios%)
'                Else
'                    txtCodigo.TEXT = txtCodigo.TEXT & TRIM$(.MSF.TextMatrix(1, i))
'                End If
'                idTabla% = ta.ObtenerIdTablaAuxiliar(Campos$(i + 3))
'                If idTabla% > 0 Then
'                    DESCRIP$ = TRIM$(VALOBADA("datasql", "campo3", "id_lista=" & CStr(idTabla) & " and campo2='" & .MSF.TextMatrix(1, i) & "'"))
'                    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then
'                        If TRIM$(Label7) <> "5" Then
'                            If InStr(1, DESCRIP$, "S/C") < 1 And InStr(1, DESCRIP$, "N/A") < 1 Then
'                                Text2 = Text2 & DESCRIP$ & " "
'                            End If
'                        Else
'                            If i = 1 Then
'                                Text2 = Text2 & DESCRIP$ & " "
'                            End If
'                        End If
'                    Else
'                        Text2 = Text2 & DESCRIP$ & " "
'                    End If
'                End If
'            Next i
'
'            txtCodigoSecundario.TEXT = ""
'            If Len(txtCodigo.TEXT) > 20 Then
'                txtCodigoSecundario.TEXT = txtCodigo.TEXT
'                txtCodigo.TEXT = ""
'            End If
'        End If
'    End With
    
End Sub


Private Sub Command3_Click()
    Call CIERRARCH("*.*")
    Unload Me
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
    Call APLICACIONSKINS(Me)
   
    Call CARGA_ENCABEZADO(mfgCodigosFamilia, "Codigo/A16;Descripción/A40", FRMZELECHO)
    
    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then
        chkTrazable.Value = 1
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



Private Sub Text1_LostFocus()
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
    TXTCODIGO.TEXT = ""
    txtCodigoSecundario.TEXT = ""
    Call CARGA_ENCABEZADO(mfgCodigosFamilia, "Codigo/A16;Descripción/A40", FRMZELECHO)
    mfgAltaPorFamilia.Rows = 1
    mfgAltaPorFamilia.Rows = 2
    mfgAltaPorFamilia.Cols = 12
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
    SQLX$ = "SELECT * FROM MASTER WHERE CODIGO LIKE '" & ABUS$ & "%'"
    Dim ARTLITMP As ADODB.Recordset
    Set ARTLITMP = New ADODB.Recordset
    ARTLITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

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
            mfgCodigosFamilia.Rows = mfgCodigosFamilia.Rows + 1
          
            mfgCodigosFamilia.TextMatrix(mfgCodigosFamilia.Rows - 1, 1) = SAFETEXT$(!Codigo)
            mfgCodigosFamilia.TextMatrix(mfgCodigosFamilia.Rows - 1, 2) = SAFETEXT$(!Descripcion)
          
19          .MoveNext
        Loop
    End With
    ARTLITMP.Close
    Set ARTLITMP = Nothing
    '
End Sub

Private Sub Text1_DblClick()
    '
    Command9_Click
    '
End Sub
