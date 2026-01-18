VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRMREGARMADO 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Registro de Armado"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   3345
      TabIndex        =   18
      Top             =   50
      Width           =   3375
      Begin VB.Image IMMM 
         Height          =   1620
         Left            =   0
         Picture         =   "FRMREGARMADO.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Height          =   1815
      Left            =   3480
      TabIndex        =   0
      Top             =   -30
      Width           =   7815
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
         Left            =   3000
         TabIndex        =   21
         Top             =   960
         Width           =   180
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO"
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "REVISION"
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
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE N°"
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1005
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
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
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1365
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   9
         Top             =   1320
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "OF N°"
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
         Index           =   5
         Left            =   3360
         TabIndex        =   8
         Top             =   285
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N° DE D.T."
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
         Index           =   6
         Left            =   3360
         TabIndex        =   6
         Top             =   645
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   4560
         TabIndex        =   5
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARO"
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
         Index           =   7
         Left            =   3000
         TabIndex        =   4
         Top             =   1005
         Width           =   1485
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   4560
         TabIndex        =   3
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FIRMA"
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
         Index           =   8
         Left            =   3360
         TabIndex        =   2
         Top             =   1365
         Width           =   1125
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   4560
         TabIndex        =   1
         Top             =   1320
         Width           =   1545
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Bindings        =   "FRMREGARMADO.frx":1693E
      Height          =   6855
      Left            =   0
      TabIndex        =   17
      Top             =   1800
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12091
      _Version        =   393216
      BackColorFixed  =   14865344
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo "
      Begin VB.Menu mnuPrintRegArmado 
         Caption         =   "Imprimir Registro de Armado"
      End
   End
End
Attribute VB_Name = "FRMREGARMADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub PresentarDatosEncabezadoImportadoDeExcel(NSERIE$)
    SQLX$ = "SELECT * FROM ENCAREGARMA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
    Set rs = READSET(SQLX$)
    With rs
      Label13(0) = SAFETEXT$(!MODELO)
      Label13(1) = SAFETEXT$(!Revision)
      Label13(3) = FECHATEX$(CVINT(!FECHA))
      Label13(4) = SAFETEXT$(!IDSUBOR)
      Label13(5) = SAFETEXT$(!DESV_TRANS)
      Label13(6) = SAFETEXT$(!PREPARO)
      Label13(7) = ""
   End With
   rs.Close
   Set rs = Nothing
End Sub

Private Sub Command4_Click()
   Call SELECHO("INDINUSE")
   TTXX$ = TRIM$(VALACT1$("INDINUSE"))
   If TTXX$ <> "" Then Label13(2) = TRIM$(TTXX$)
End Sub

Private Sub Form_Load()
'   Call CREAR_TABLA_IMPOREGDATOEYM
   Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0;TABLA/A0"
   Call CARGA_ENCABEZADO(MSF4, Campos$, Me)

End Sub

Private Sub Label13_Change(Index As Integer)
    If Index <> 2 Then Exit Sub
    NSERIE$ = Label13(2)
    Label2(1).Visible = False
    If ExisteNSerieEnTanumse%(NSERIE$) < 1 Then Exit Sub
    'CAPTURO DATOS DE TANUMSE
    SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
    Set rs = READSET(SQLX$)
    With rs
      Codigo = SAFETEXT$(!COD_EXTE)
      CIFAB% = XVALO(!COD_INTE)
      NORFA$ = SAFETEXT$(!IDSUBOR) 'OF
      FEALTA% = CVINT(!FECHALTA)
    End With
    rs.Close: Set rs = Nothing
    '
    
    If NumeroRegistroSegunSerieEym&(NSERIE$) > 0 Then
       'PERTENECE A CARGA POR IMPORTACION REALIZADA DESDE EXCEL
       Label2(1) = "Registro Importado desde Excel"
       Label2(1).Visible = True
       Label2(1).BackColor = vbRed
       Call PresentarDatosEncabezadoImportadoDeExcel(NSERIE$)
       Call PresentarRegArmadoImportado(NSERIE$)
    Else
       'PERTENECE A CARGA NUEVA POR ORDEN DE FABRICACION
       Label2(1).Visible = False
       Call RegistroArmadoEym(NORFA$, NSERIE$, FEALTA%, CIFAB%, Me.MSF4)
    End If
   
End Sub
Sub PresentarRegArmadoImportado(NSERIE$)

  MSF4.Rows = 1
  'PRESENTO DetaRegArma EL DETALLE
  SQLX$ = "SELECT * FROM DetaRegArma WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE NumeroSerie = '" & NSERIE$ & "'"
  SQLX$ = SQLX$ + " order by orden "
  Set rs = READSET(SQLX$)
  With rs
     Do While Not .EOF
       i& = i& + 1
       MSF4.Rows = i& + 1
       MSF4.TextMatrix(i&, 1) = SAFETEXT(!Codigo)
       MSF4.TextMatrix(i&, 2) = SAFETEXT(!Descripcion)
       MSF4.TextMatrix(i&, 3) = SAFETEXT(!IDENTIFICACION)
       MSF4.TextMatrix(i&, 4) = SAFETEXT(!NROdeDT)
       MSF4.TextMatrix(i&, 5) = SAFETEXT(!DetaRegArma_id)
       MSF4.TextMatrix(i&, 6) = ""
       MSF4.TextMatrix(i&, 7) = "DetaRegArma"
       
       .MoveNext
     Loop
  End With
  rs.Close
  Set rs = Nothing
  '


End Sub
Sub RegistroArmadoEym(NORFA$, NSERIE$, FEALTA%, CIFAB%, MSF As MSFlexGrid)

        'uso el ventabnew para mantener el codigo copiado
        'luego paso la inforamcion a pantalla (encabezado a las cajas de texto y detalle a la grilla)
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0;TABLA/A0"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        'VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        If CAMPOSEDITABLES$ <> "" Then
           VENTANA.CampEditables = STRINGCAMPOSEDITABLES$
        Else
           VENTANA.CampEditables = "1;2;3;4"
        End If
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Impresión F054"
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 7500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        FALTACPU% = 0
        'CARGO PRIMERO LOS DATOS DEL CPU
        SQLX$ = "SELECT CODIGO,DESCRIPCION,IDENTIFICACION,DESV_TRANS,DATOSCPU_ID,COD_INTE FROM DATOSCPU WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
        'AND IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CIFAB%
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!Codigo)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = SAFETEXT$(!IDENTIFICACION)
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!DATOSCPU_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!IDENTIFICACION)) = "" Then
               FALTACPU% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If
            VENTABNEW.MSF.TextMatrix(k&, 7) = "DATOSCPU"

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        
        If k& < 1 Then
          Call COMUNI("NO SE ENCONTRARON DATOS CPU")
        End If
        
        If k& < 3 Then k& = 2
        FALTALOTE% = 0
        CONDI$ = "COD_EXTE = '" & Label44 & "' AND  IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
        OF_NIVEL_SUPERIOR$ = VALOBADA("TRAZAXSERIE", "OF_ARRASTRA", CONDI$, "NOMESS")
        'TRAE TODOS LOS COMPONENTES CRITICOS CON EXCEPCION DE LOS DEL PRIMER NIVEL DE LA OF SELECCIONADA
        'SELECT * FROM TRAZAXSERIE WITH(NOLOCK)  WHERE NUMEROSERIE = '185-1473'  AND IDSUBOR <> 'OF-002221-1'
        SQLX$ = "SELECT COD_EXTE,IDENLOTE,TRAZAXSERIE_ID,COD_INTE,DESCRIPCION,DESV_TRANS FROM TRAZAXSERIE WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
        SQLX$ = SQLX$ + " AND  COD_EXTE <> '" & Label44 & "' AND  IDSUBOR <> '" & OF_NIVEL_SUPERIOR$ & "'"
        SQLX$ = SQLX$ + " ORDER BY  FECHALTA DESC"
        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            k& = k& + 1
            VENTABNEW.MSF.Rows = k& + 1
            LOTEX$ = SAFETEXT$(!idenlote)
            LOTEX$ = REPLACAR(LOTEX$, "BR", "PR")
            VENTABNEW.MSF.TextMatrix(k&, 1) = SAFETEXT$(!COD_EXTE)
            VENTABNEW.MSF.TextMatrix(k&, 2) = SAFETEXT$(!Descripcion)
            VENTABNEW.MSF.TextMatrix(k&, 3) = LOTEX$
            VENTABNEW.MSF.TextMatrix(k&, 4) = SAFETEXT$(!DESV_TRANS)
            VENTABNEW.MSF.TextMatrix(k&, 5) = XVALO(!TRAZAXSERIE_ID)
            VENTABNEW.MSF.TextMatrix(k&, 6) = XVALO(!COD_INTE)
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(VENTABNEW.MSF, &H808080, k&)
            If TRIM$(SAFETEXT$(!idenlote)) = "" Then
               FALTALOTE% = 1 'PARA IDENTIFICAR SI FALTAN DATOS
               Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, k&, 2)
            End If
            VENTABNEW.MSF.TextMatrix(k&, 7) = "TRAZAXSERIE"

            .MoveNext
          Loop
        End With
        rs.Close
        Set rs = Nothing
        ' +
        'COMPLETA LOS DATOS FIJOS DESDE EL EXCEL Y LOS IMPRIME EN MSF4
        Call completarDatosFijosDetalleRegDatos(MSF4)
        '
        PREPARADOPORX$ = VENTABNEW.MSF.TextMatrix(1, 3)

        'ESCRIBO LOS DATOS EN OTROS CASILLEROS QUE ESTAN REFERENCIADOS
        For i& = 2 To VENTABNEW.MSF.Rows - 1
            CODI$ = UCase$(TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)))
            IDENTIF$ = VENTABNEW.MSF.TextMatrix(i&, 3)
            DV$ = CEROIGUALVACIO(VENTABNEW.MSF.TextMatrix(i&, 4))
            For j& = 1 To MSF4.Rows - 1
               CODIEXCEL$ = MSF4.TextMatrix(j&, 1)
               If CODIEXCEL$ = CODI$ Then
                  MSF4.TextMatrix(j&, 3) = IDENTIF$
                  MSF4.TextMatrix(j&, 4) = CEROIGUALVACIO(DV$)
                  MSF4.TextMatrix(j&, 5) = VENTABNEW.MSF.TextMatrix(i&, 5) 'id
                  MSF4.TextMatrix(j&, 6) = VENTABNEW.MSF.TextMatrix(i&, 6) 'ci1%
                  MSF4.TextMatrix(j&, 7) = VENTABNEW.MSF.TextMatrix(i&, 7) 'TABLA%
                  Exit For
               End If
            Next j&
        Next i&
        '
        'ESTOS DATOS SON DE ENCABEZADO LOS ESCRIBO DIRECTAMENTE EN LABEL CORRESPONDIENTE CORRESPONDIENTE
        For i& = 1 To 7
         Select Case i&
           Case 1: DATO$ = CODEXT$(CIFAB%): Label13(0) = DATO$
           Case 2: DATO$ = REVICOD$(CIFAB%): Label13(1) = DATO$
           Case 3: DATO$ = NSERIE$: Label13(2) = DATO$
           Case 4: DATO$ = FECHATEX$(FEALTA%): Label13(3) = DATO$
           Case 5: DATO$ = NORFA$:  Label13(4) = DATO$
           Case 6: DATO$ = DESVIOTRANSITORIO$:  Label13(5) = DATO$
           Case 7: DATO$ = PREPARADOPORX$:  Label13(6) = DATO$
         End Select
        Next i&
End Sub

Sub completarDatosFijosDetalleRegDatos(MSF As MSFlexGrid)
        Dim ObjetoExcel As Excel.Application
        Dim Libro As Excel.Workbook
        Dim Hoja As Excel.Worksheet
        MSF.Rows = 1
        PDES$ = CONTROL$(IDENTER$, "PRTQEYM3")
        If PDES$ = "" Then PDES$ = CONTROL$("*", "PRTQEYM3")
        RUTAEXCEL$ = CONTROL("TRAZABLE", "EXCEARMA")
        
        MSF.Rows = 1
        Set ObjetoExcel = CreateObject("Excel.Application")
        Set Libro = ObjetoExcel.Workbooks.Open(RUTAEXCEL$)
        With ObjetoExcel
            '
             For R& = 16 To 37
                REG& = REG& + 1
                MSF.Rows = REG& + 1
                CODIEXCEL$ = UCase$(TRIM$(Libro.Worksheets("Hoja1").Cells(R&, 2)))
                DESCRIEXCEL$ = UCase$(TRIM$(Libro.Worksheets("Hoja1").Cells(R&, 3)))
                MSF.TextMatrix(REG&, 1) = CODIEXCEL$
                MSF.TextMatrix(REG&, 2) = DESCRIEXCEL$
                'DATOS POR DEFECTO
                MSF.TextMatrix(REG&, 3) = "N/A"
                MSF.TextMatrix(REG&, 4) = ""
             Next R&

             'MUESTRO E IMPRIMO
             ObjetoExcel.ActiveWorkbook.Save
             Set Hoja = Libro.Sheets("Hoja1")
             '
        End With
        
   
        'CIERRA HOJA, LIBRO Y EXCEL
99      On Error Resume Next
        Set Hoja = Nothing
        Libro.Close: Set Libro = Nothing
        ObjetoExcel.Close: Set ObjetoExcel = Nothing
        On Error GoTo 0

End Sub

Private Sub mnuPrintRegArmado_Click()
   Call COMUNI("En Proceso")
   Exit Sub

        'SELECCION DE OF DE LOS PRODUCTOS TERMINADOS
        NORFA$ = TRIM$(Label13(4))
        If NORFA$ = "" Then
           Call COMUNI("No Hay una Orden de Fabricación Activa")
           Exit Sub
        End If
        '
        NSERIE$ = TRIM$(Label13(2))
        If NSERIE$ = "" Then
           Call COMUNI("Falta Seleccionar Un Número de Serie")
           Exit Sub
        End If
        '
        'VALIDA SI YA SEW GRABARON LOS DATOS DEL CPU
        'CIFAB% = XVALO(Me.TXTCI)
        
        '
        Dim ObjetoExcel As Excel.Application
        Dim Libro As Excel.Workbook
        Dim Hoja As Excel.Worksheet
        
        PDES$ = CONTROL$(IDENTER$, "PRTQEYM3")
        If PDES$ = "" Then PDES$ = CONTROL$("*", "PRTQEYM3")
        RUTAEXCEL$ = CONTROL("TRAZABLE", "EXCEARMA")
    
        Set ObjetoExcel = CreateObject("Excel.Application")
        Set Libro = ObjetoExcel.Workbooks.Open(RUTAEXCEL$)
        With ObjetoExcel
            '
             'LIMPIO PRIMERO LOS CAMOS EN EL EXCEL
             For R& = 16 To 37
                Libro.Worksheets("Hoja1").Cells(R&, 6) = "N/A"
                Libro.Worksheets("Hoja1").Cells(R&, 8) = " "
             Next R&
            
             'ESCRIBO LOS DATOS EN OTROS CASILLEROS QUE ESTAN REFERENCIADOS
             For i& = 2 To MSF4.Rows - 1
                CODI$ = UCase$(TRIM$(MSF4.TextMatrix(i&, 1)))
                IDENTIF$ = MSF4.TextMatrix(i&, 3)
                DV$ = CEROIGUALVACIO(MSF4.TextMatrix(i&, 4))
                For R& = 16 To 37
                   CODIEXCEL$ = UCase$(TRIM$(Libro.Worksheets("Hoja1").Cells(R&, 2)))
                   If CODIEXCEL$ = CODI$ Then
                      Libro.Worksheets("Hoja1").Cells(R&, 6) = IDENTIF$
                      Libro.Worksheets("Hoja1").Cells(R&, 8) = CEROIGUALVACIO(DV$)
                      Exit For
                   End If
                Next R&
             Next i&

             'ESTOS DATOS SON DE ENCABEZADO LOS ESCRIBO DIRECTAMENTE EN LA CELDA CORRESPONDIENTE
             'EN LA POSICION A PARITR DEL REGISTRO 4 EN COLUMNA 6
             For i& = 0 To Label13.UBound 'LABEL 13 ES LABEL VECTOR DE 0 A 7
                DATO$ = Label13(i&): Libro.Worksheets("Hoja1").Cells(i& + 4, 6) = DATO$
             Next i&
             '
             'MUESTRO E IMPRIMO
             ObjetoExcel.ActiveWorkbook.Save
             Set Hoja = Libro.Sheets("Hoja1")
             Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=PDES$
             '
        End With
        
    
        'CIERRA HOJA, LIBRO Y EXCEL
99      On Error Resume Next
        Set Hoja = Nothing
        Libro.Close: Set Libro = Nothing
        ObjetoExcel.Close: Set ObjetoExcel = Nothing
        On Error GoTo 0
    
   
End Sub




Private Sub MSF4_DblClick()

    If DERACCE("MOREGARM", "MODIFICAR REGISTRO DE ARMADO") < 2 Then Exit Sub
    If MSF4.Rows - 1 = 0 Then Exit Sub
    Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0;TABLA/A22"

    Unload VENTABNEW
    REG& = MSF4.MouseRow
    If REG& < 1 Then Exit Sub
    ID& = XVALO(MSF4.TextMatrix(REG&, 5))
    VENTANA.Inicializar = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.AgregaRegistro = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.Campox = Campos$
    VENTANA.TITULO = "Modficación de Datos de Registro de Armado"
    VENTANA.CampEditables = "3;4"
    VENTANA.Inicializar = 1
'    Campos$ = "Código:/A16;Descripción:/A32;Identificación Nº.:/A16;Nº de D.T.:;ID/A0;CI/A0"
   
    Call EDITARFILA(MSF4, VENTABNEW.MSF, REG&)
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
         IDENTIFICACION$ = VENTABNEW.MSF.TextMatrix(1, 3)
         DESVIOTRANSITORIO$ = VENTABNEW.MSF.TextMatrix(1, 4)
         OPX% = COMALTER("Desea Grabar Los Cambios Realizados en la Fila Seleccionada\\Cancelar\Sí,Grabar")
         If OPX% = 2 Then
           
           If Label2(1).Visible = True Then 'REGISTRO DE ARMADO DE IMPORTACION
            CONDI$ = "DetaRegArma_id = " & ID&
            Call ACTUREGISTRO("DetaRegArma", "IDENTIFICACION", CONDI$, IDENTIFICACION$, REGAF&, "NOMESS")
            Call ACTUREGISTRO("DetaRegArma", "NROdeDT", CONDI$, DESVIOTRANSITORIO$, REGAF&, "NOMESS")
           Else
            TABLA$ = VENTABNEW.MSF.TextMatrix(1, 7)
            CONDI$ = TABLA$ & "_id = " & ID&
            CAMPO$ = "IDENTIFICACION"
            If TABLA$ = "TRAZAXSERIE" Then
               CAMPO$ = "IDENLOTE"
            End If

            Call ACTUREGISTRO(TABLA$, CAMPO$, CONDI$, IDENTIFICACION$, REGAF&, "NOMESS")
            Call ACTUREGISTRO(TABLA$, "DESV_TRANS", CONDI$, DESVIOTRANSITORIO$, REGAF&, "NOMESS")
           End If
           Call PEGARFILAEDITADA(VENTABNEW.MSF, MSF4, REG&)
         End If
    End If


End Sub


