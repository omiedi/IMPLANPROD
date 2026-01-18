VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TABLAUXI 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   15105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command25 
      Caption         =   "Guardar y Genera Código"
      Height          =   615
      Left            =   11640
      TabIndex        =   24
      Top             =   6840
      Width           =   3015
   End
   Begin VB.CommandButton Command22 
      Caption         =   "TEST"
      Height          =   375
      Left            =   11760
      TabIndex        =   23
      Top             =   7635
      Width           =   2895
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   1635
      IntegralHeight  =   0   'False
      ItemData        =   "TABLAUXI.frx":0000
      Left            =   1800
      List            =   "TABLAUXI.frx":0002
      TabIndex        =   21
      Top             =   4200
      Width           =   2625
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Top             =   3840
      Width           =   2625
   End
   Begin VB.CheckBox Check18 
      Caption         =   "Levanta de Recordset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   0
      TabIndex        =   19
      Top             =   4920
      Width           =   1560
   End
   Begin VB.CheckBox Check14 
      Caption         =   "UtilizaFormula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CheckBox Check15 
      Caption         =   "HabilitaBorrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CheckBox Check16 
      Caption         =   "HabilitaInsertar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CheckBox Check17 
      Caption         =   "NoMostrarBotonTilde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   4680
      Width           =   2680
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Agrega Registos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   1305
      Left            =   4695
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3960
      Width           =   10200
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      Height          =   1305
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5430
      Width           =   10200
   End
   Begin VB.TextBox TXTIDLISTA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   1305
   End
   Begin VB.CommandButton Command23 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4990
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command26 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5400
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox TXTMOVIL2 
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8280
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "   "
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Eliminar estructura Activa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   0
      ToolTipText     =   "Eliminar Estructura Activa"
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Height          =   2805
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Doble Click Para Editar"
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   4948
      _Version        =   393216
      Cols            =   4
      FixedRows       =   0
      BackColor       =   16777215
      BackColorFixed  =   16501405
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      GridLinesFixed  =   1
      SelectionMode   =   1
      MergeCells      =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label20 
      Caption         =   "Título de Recordset"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1920
      TabIndex        =   22
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label20 
      Caption         =   "Consulta para Recordset  (Ejemplo: Select campo1, campo2 from datasql with(nolock))"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4680
      TabIndex        =   14
      Top             =   3720
      Width           =   10200
   End
   Begin VB.Label Label20 
      Caption         =   "Código Generado para que se cree la estrucutura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4680
      TabIndex        =   13
      Top             =   5235
      Width           =   10200
   End
   Begin VB.Label Label24 
      Caption         =   "Cant.Campos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "TABLAUXI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KLIK%
Dim MODIFENCA%, ILIS1&
Dim ESTADO$
    ' variable para la clase
    Dim tTip As CLSTOOLTIP


Sub BLANFORMU()
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
        Check8.Value = 0
        Check9.Value = 0
        Check10.Value = 0
        Check11.Value = 0
        Text10 = ""
        Text11 = ""
        Text9 = ""
        Text8 = 4
        TXTCONSU_ESPEJO = ""
        TitVentana = ""
        Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(VENTADISE.MSFESPEJO)
        Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(VENTADISE.MSF)
        List3.Clear

End Sub

Sub CARGADATOSUNICOS(IDLISTA&)
  Dim obj As New RECORXET
  Set RS1 = obj.RS_TAB_COMP_ENCAPRFUN(IDLISTA&)
  Text11 = XVALO(RS1!CANTICAMPOS)
  Text12 = SAFETEXT$(RS1!TITURECORD)
  Text13 = SAFETEXT$(RS1!ConsultaSql)
  Text14 = SAFETEXT$(RS1!Descripcion)
  Check13.Value = XVALO(RS1!ADDREGISTROS)
  Check14.Value = XVALO(RS1!USAFORMULA)
  Check15.Value = XVALO(RS1!BORRAR)
  Check16.Value = XVALO(RS1!INSERTAR)
  Check17.Value = XVALO(RS1!NOTILDE)
  Check18.Value = XVALO(RS1!USARECORD)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing

End Sub

Sub CARGAMSF4()
  MSF4.Rows = 8
  MSF4.ColWidth(0) = 4000
  MSF4.TextMatrix(0, 0) = "Encabezado + Formato (Canti/F8.1)"
  MSF4.TextMatrix(1, 0) = "Formula (4;*;6)"
  MSF4.TextMatrix(2, 0) = "CampEditables (S/N)"
  MSF4.TextMatrix(3, 0) = "ColEcoCampo (CODIGO;Descripcion;Master)"
  MSF4.TextMatrix(4, 0) = "TxtUbicadoGrillaTotal(Texto)"
  MSF4.TextMatrix(5, 0) = "Total (S/N)"
  MSF4.TextMatrix(6, 0) = "Dejar Libre (S/N)"
  MSF4.TextMatrix(7, 0) = "NUMERICO=0 - TEXTO=1"

End Sub

Sub EscribeEstrucSql(CODICONSULTA$, TEXTOAESCRIBIR$)
'Sub GENERA_ESTRUCTURAS()
   ILET$ = "D"
10 RUT1$ = ILET$ + ":\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT\BIB_EST_SQL\"
   RUT1$ = RUT1$ + "EscribeEstrucSql.bas"
   '
   On Error Resume Next
   If EXISTE%(RUT1$) < 1 Then
     If ILET$ = "D" Then ILET$ = "C": GoTo 10
     MsgBox "Imposible Abrir Archivo: " + RUT1$
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   CARPETOK$ = ILET$ & ":\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT\BIB_EST_SQL\"

   NARCHI1% = FILEOPEN%(CARPETOK$ + "EscribeEstrucSql.bas", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
   NARCHI2% = FILEOPEN%(CARPETOK$ + "EscribeEstrucSql2.bas", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(XX$))
       If XX1$ = "'|INICIO " & CODICONSULTA$ Then NOESCRIBIR% = 1
       If NOESCRIBIR% < 1 Then
          If UCase$(TRIM$(XX1$)) <> "END SUB" Then
            Print #NARCHI2%, XX$
          End If
       End If
    Loop
    Print #NARCHI2%, TEXTOAESCRIBIR$
    Print #NARCHI2%, "End Sub"
    Close #NARCHI1%
    Close #NARCHI2%
    
    On Error Resume Next
    Kill (CARPETOK$ + "EscribeEstrucSql.bas")
    On Error GoTo 0
    '
    'COPIO EL ARCHIVO ESCRITO SOBRE EL ORIGINAL
    XY1% = COPYFILE(CARPETOK$ + "EscribeEstrucSql2.bas", CARPETOK$ + "EscribeEstrucSql.bas")
    '

    '
99
End Sub


Sub GENERA_AGREGA_RECORDSET(CAMPOS_GRI$, Optional VENTANAX$)
    If TRIM$(VENTANAX$) = "" Then
      VENTANAX$ = "VENTABNEW.MSF"
    End If
    
    TX1$ = "  Dim obj As New RECORXET" & vbCrLf
    CAMPOS_BD$ = ""
    For i& = 0 To List3.ListCount - 1
       List3.ListIndex = i&
       CAMPOS_BD$ = CAMPOS_BD$ & List3.TEXT & ","
    Next i&
    CAMPOS_BD$ = TRIM$(CAMPOS_BD$)
    If Right(CAMPOS_BD$, 1) = "," Then CAMPOS_BD$ = Mid$(CAMPOS_BD$, 1, Len(CAMPOS_BD$) - 1)
    TX1$ = TX1$ + "  Set RS1 = obj.RS_" & TRIM$(Text10) & "()" & vbCrLf ' FUNCION QUE DEVUELVE UN RECORDSET" & vbCrLf
    
    TX1$ = TX1$ + "  Call Carga_MSF_Recordset(RS1," & TRIM$(CAMPOS_GRI$) & "," & TRIM$(VENTANAX$) & ")" & vbCrLf
    
    TX1$ = TX1$ + "  Set obj = Nothing" & vbCrLf
    TX1$ = TX1$ + "  RS1.Close" & vbCrLf
    TX1$ = TX1$ + "  Set RS1 = Nothing" & vbCrLf
    
    TX1$ = TX1$ + "'*****************************************************'" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    TX1$ = TX1$ + "ESTO VA EN LA CLASE RECORXET" & vbCrLf
    NOMB_CLASE$ = "RS_" & TRIM$(Text10)
    TX1$ = TX1$ + "Public Function " & NOMB_CLASE$ & "() As ADODB.Recordset " & vbCrLf
    TX1$ = TX1$ + "  Dim RS_FUN As ADODB.Recordset" & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = New ADODB.Recordset" & vbCrLf
    'TX1$ = TX1$ + "  If TRIM$(CAMPOS_BD$) = " & Chr$(34) & Chr$(34) & " Then CAMPOS_BD$ = " & Chr$(34) & " * " & Chr$(34) & vbCrLf
    
    TX1$ = TX1$ + "  SQLX$ = " & Chr$(34) & TXTCONSU_ESPEJO & Chr$(34) & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = READSET(SQLX$)" & vbCrLf
    TX1$ = TX1$ + "  Set " & NOMB_CLASE$ & " = RS_FUN" & vbCrLf
    TX1$ = TX1$ + "  Set RS_FUN = Nothing" & vbCrLf
    TX1$ = TX1$ + "End Function" & vbCrLf
    TX1$ = TX1$ + "'*****************************************************'" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    Text9 = Text9 + TX1$
    
    
 End Sub

 Function GENERA_CODI_MENUGRILLA$()
 
  TXTMG$ = "Dim GLOBALFRM_strContenidoCelda$" & vbCrLf
  TXTMG$ = TXTMG$ + "Dim GLOBALFRM_FILA&" & vbCrLf
  TXTMG$ = TXTMG$ + "Public NORFAEXTERNA$" & vbCrLf
  TXTMG$ = TXTMG$ + "Private Sub mnucopiar_Click()" & vbCrLf
  TXTMG$ = TXTMG$ + "       Clipboard.Clear" & vbCrLf
  TXTMG$ = TXTMG$ + "       Clipboard.SetText GLOBALFRM_strContenidoCelda$" & vbCrLf
  TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf
TXTMG$ = TXTMG$ + "  If Button = 2 Then" & vbCrLf
TXTMG$ = TXTMG$ + "     REG& = MSF1.MouseRow" & vbCrLf
TXTMG$ = TXTMG$ + "     COL% = MSF1.MouseCol" & vbCrLf
TXTMG$ = TXTMG$ + "     Call MENU_CELDA(MSF1, REG&, COL%)" & vbCrLf
TXTMG$ = TXTMG$ + "  End If" & vbCrLf
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
 TXTMG$ = TXTMG$ + " Sub INVICTRL()" & vbCrLf
'  '
TXTMG$ = TXTMG$ + "   mnucopiar.Visible = Not (mnucopiar.Visible)" & vbCrLf
TXTMG$ = TXTMG$ + "  mnudesap.Visible = Not (mnudesap.Visible)" & vbCrLf
TXTMG$ = TXTMG$ + "'  InformaCantidadEnProceso.Visible = Not InformaCantidadEnProceso.Visible" & vbCrLf
'  '
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)" & vbCrLf
TXTMG$ = TXTMG$ + "        If MSF.Rows <= 1 Then Exit Sub" & vbCrLf
'
TXTMG$ = TXTMG$ + "        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA" & vbCrLf
TXTMG$ = TXTMG$ + "        Call INVICTRL"
TXTMG$ = TXTMG$ + "        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA" & vbCrLf
TXTMG$ = TXTMG$ + "        GLOBALFRM_FILA& = REG&" & vbCrLf
TXTMG$ = TXTMG$ + "        PopupMenu MNUARCHIVO" & vbCrLf
TXTMG$ = TXTMG$ + "        Call INVICTRL" & vbCrLf
TXTMG$ = TXTMG$ + "        On Error Resume Next" & vbCrLf
TXTMG$ = TXTMG$ + "        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA" & vbCrLf
TXTMG$ = TXTMG$ + "        On Error GoTo 0" & vbCrLf
TXTMG$ = TXTMG$ + "End Sub" & vbCrLf
ENCAB$ = "SE DEBERA GENEAR EN EL MENU" & vbCrLf
ENCAB$ = ENCAB$ + "'AGREGAR EN MENU'" & vbCrLf
ENCAB$ = ENCAB$ + "'CAPTION: Archivo - NO VISIBLE" & vbCrLf
ENCAB$ = ENCAB$ + "'....CAPTION -      Copiar Celda Seleccionada" & vbCrLf
ENCAB$ = ENCAB$ + "'.....NAME:  S   - VISIBLE" & vbCrLf
ENCAB$ = ENCAB$ + "'SI SE NECESITA PARA OTRA GRILLA SE DEBERA DUPLICAR LA INFORMACION RENOMBRANDO CADA METODO AGREGANDO UN 2"
GENERA_CODI_MENUGRILLA$ = ENCAB$ & TXTMG$
End Function

Sub GENERAR_CODIGO(OK%)
    If Check10.Value > 0 Then
      If TRIM$(Text10) = "" Then
        Call COMUNI("Falta Título de Recordset")
        OK% = -1
        Exit Sub
      End If
    End If
    '
    If Check10.Value > 0 Then
        POS1% = InStr(TXTCONSU_ESPEJO, "(NOLOCK)")
        If POS1% < 1 Then
           Call COMUNI("Falta With(NOLOCK) en consulta")
           OK% = -1
           Exit Sub
        End If
    End If
    '
    
    T1$ = "VENTANA.Inicializar = 0" + vbCrLf
    Campos$ = ""
    For i& = 1 To MSF.Cols - 1
       If TRIM$(MSF.TextMatrix(0, i&)) <> "" Then Campos$ = Campos$ & MSF.TextMatrix(0, i&) + ";"
    Next i&
    If Right$(Campos$, 1) = ";" Then Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    Campos$ = Chr$(34) & Campos$ & Chr$(34) & vbCrLf
    For i& = 1 To MSF.Rows - 1
       PROPIEDADX$ = MSF.TextMatrix(i&, 0)
       Select Case PROPIEDADX$
          Case "Formula (4;*;6)"
            For J% = 1 To MSF.Cols - 1
              If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 T0$ = "VENTANA.formula(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                 If TRIM$(T0$) <> "" Then If Right$(T0$, 1) = ";" Then T0$ = Mid$(T0$, Len(T0$) - 1)
                 T1$ = T1$ + T0$
              End If
            Next J%
            
          Case "CampEditables (S/N)"
              For J% = 1 To MSF.Cols - 1
               If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 'VENTANA.CampEditables(" & J% & ") = " & MSF.TextMatrix(i&, J%) + vbCrLf"
                 If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then T2$ = T2$ & J% & ";"
               End If
              Next J%
              If TRIM$(T2$) <> "" Then
                 If Right$(T2$, 1) = ";" Then T2$ = Mid$(T2$, 1, Len(T2$) - 1)
'                 T2$ = Chr$(34) & T2$ & Chr$(34)
                 T1$ = T1$ + "VENTANA.CampEditables = " & Chr$(34) & T2$ & Chr$(34) & vbCrLf
              End If
              
          Case "TxtUbicadoGrillaTotal(Texto)"
              For J% = 1 To MSF.Cols - 1
               If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                 H% = J%
                 'VENTANA.TxtUbicadoGrillaTotal(J%) = MSF.TextMatrix(i&, J%)
                 T3$ = "VENTANA.TxtUbicadoGrillaTotal(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) + vbCrLf
                 If TRIM$(T3$) <> "" Then If Right$(T3$, 1) = ";" Then T3$ = Mid$(T3$, Len(T3$) - 1) + vbCrLf
                 T1$ = T1$ + T3$
               End If
              Next J%
          Case "ColEcoCampo (CODIGO;Descripcion;Master)"
              For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.ColEcoCampo(J%) = MSF.TextMatrix(i&, J%)
                  T4$ = "VENTANA.ColEcoCampo(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T4$
''                  If TRIM$(MSFESPEJO.TextMatrix(i&, J%)) <> "" Then
''                    T4$ = T4$ & ";" & TRIM$(MSFESPEJO.TextMatrix(i&, J%)) & Chr$(34) & vbCrLf
''                    T1$ = T1$ + T4$
''                  End If
                End If
              Next J%
           Case "Total (S/N)"
              For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T5$ = "VENTANA.Totales(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T5$
                End If
              Next J%
           Case "Dejar Libre (S/N)"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T6$ = "VENTANA.DejarLibre(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T6$
                End If
              Next J%
           Case "CONDICION 1"
               For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.Condicion1(J%) = MSF.TextMatrix(i&, J%)
                  T7$ = "VENTANA.Condicion1(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T7$
                End If
              Next J%
           Case "CONDICION 2"
               For J% = 1 To MSF.Cols - 1
                If TRIM$(MSF.TextMatrix(i&, J%)) <> "" Then
                  VENTANA.Condicion2(J%) = MSF.TextMatrix(i&, J%)
                  T8$ = "VENTANA.Condicion2(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T8$
                End If
              Next J%
           Case "ENUMERAR"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.Enumerar(J%) = MSF.TextMatrix(i&, J%)
                  T9$ = "VENTANA.ENUMERAR(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T9$
                End If
              Next J%
         
       End Select
    Next i&
    

    T1$ = T1$ + "VENTANA.AgregaRegistro = " & Check5.Value & vbCrLf
    T1$ = T1$ + "VENTANA.UtilizaFormula = " & Check6.Value & vbCrLf
    T1$ = T1$ + "VENTANA.HabilitaBorrar = " & Check7.Value & vbCrLf
    T1$ = T1$ + "VENTANA.HabilitaInsertar = " & Check8.Value & vbCrLf
    T1$ = T1$ + "VENTANA.NoMostrarBotonTilde = " & Check9.Value & vbCrLf
    T1$ = T1$ + "VENTANA.Titulo = " & Chr$(34) & TXTTITULO.TEXT & Chr$(34) & vbCrLf
    T1$ = T1$ + "VENTANA.Inicializar = 1" & vbCrLf
    T1$ = T1$ + "VENTANA.Campox = " & Campos$ & vbCrLf
    '
    If Check11.Value > 0 Then
         T1$ = T1$ + "*** SE DEBERA CARGAR PARA EL PARAMETRO  VENTANA.Campox = LOS CAMPOS DEL MSF BASE  ****" + vbCrLf
         T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
         T1$ = T1$ + "*** SE DEBERA CARGAR PARA LOS PARAMETROS CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)   ****" + vbCrLf
         T1$ = T1$ + "*** REALIZAR LA LLAMADA " + vbCrLf
         T1$ = T1$ + "Call CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)"
    End If
    Text9 = T1$
    T1$ = T1$ & vbCrLf
    T1$ = T1$ & "'"
    T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
    Text9 = T1$
    
    If Check10.Value > 0 Then
       Call GENERA_AGREGA_RECORDSET(Campos$)
    End If
    If Check11.Value > 0 Then T1$ = T1$ & "Call T_Espera(1 / 10)" & vbCrLf
    
    Text9 = Text9 & "VENTABNEW.SHOW 1" + vbCrLf
    
    If Check11.Value > 0 Then
      Text9 = Text9 & "If VENTABNEW.APROBADO% > 0 Then" + vbCrLf
      Text9 = Text9 + "   Call CopiarMsfAOtroMsf(VENTADISE.MSF,GRILLABASE )" + vbCrLf
      Text9 = Text9 & "End If" + vbCrLf
    End If
    '
    GECOMENUGRI$ = ""
    If Check12.Value = 1 Then
       GECOMENUGRI$ = GENERA_CODI_MENUGRILLA$
       Text9 = Text9 & GECOMENUGRI$
    End If

    Text9 = Text9 & "Set VENTANA = Nothing"
    '
End Sub

Sub GRABASQLDAT(SQLX$, RET_CAMPO$, RET_TABLA$, RET_CONDI$)
   CONSU$ = UCase$(SQLX$)
   POS1% = InStr(CONSU$, " ")
   POS2% = InStr(CONSU$, " FROM ")
   POS3% = POS2% + 6 ' DONDE COMIENZA FROM + LONGITUD DE FROM Y 2 ESPACIOS
   POS4% = InStr(CONSU$, "WITH(NOLOCK)")
   POS5% = InStr(CONSU$, "WHERE ")
   POS6% = InStr(CONSU$, "ORDER ")
   POS5% = POS5% + 6
   RET_CAMPO$ = Mid$(CONSU$, POS1%, POS2% - POS1%)
   RET_TABLA$ = Mid$(CONSU$, POS3%, POS4% - POS3%)
   If POS6% > 0 Then
     RET_CONDI$ = Mid$(CONSU$, POS5%, POS6%)
   Else
     RET_CONDI$ = Mid$(CONSU$, POS5%)
   End If
End Sub

Sub MOSTRARCONFI(IDLIS&)
'  Dim obj As New RECORXET
'  Set RS1 = obj.RS_MUESTRAENCAPRF(IDLIS&)
'  CAMPOS$ = "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10"
'
'  Call Carga_MSF_Recordset(RS1, CAMPOS$, VENTADISE.MSF4)
'  Set obj = Nothing
'  RS1.Close
'  Set RS1 = Nothing

End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub CMDAPPCAMPOS_Click()
   CMDAPPCAMPOS.Enabled = False
   MSFUPD.Rows = 5
   Campos$ = TRIM$(Me.TXTCAMPOSUPDATE)
   Call TEXTOLOTES$(Campos$, ";")
   On Error Resume Next
   CANTICAMP = UBound(CADENATEX$)
   On Error GoTo 0
   If XVALO(CANTICAMP) > 0 Then
     MSFUPD.Cols = CANTICAMP + 1
     For i& = 1 To UBound(CADENATEX$)
        MSFUPD.ColWidth(i&) = 3500
        MSFUPD.TextMatrix(0, i&) = CADENATEX$(i&)
     Next i&
   
   End If
   
   CMDAPPCAMPOS.Enabled = True
End Sub

'Private Sub CMBNOMBRECAMPO_KeyPress(KeyAscii As Integer)



Private Sub Command16_Click()
VENTANA.Inicializar = 0
VENTANA.formula(5) = "3;*;4"
VENTANA.CampEditables = "1;3;4"
VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
VENTANA.TOTALES(5) = "S"
VENTANA.AgregaRegistro = 1
VENTANA.UtilizaFormula = 1
VENTANA.HabilitaBorrar = 1
VENTANA.HabilitaInsertar = 1
VENTANA.NoMostrarBotonTilde = 0
VENTANA.TITULO = "PEDIDOS MUESTRA"
VENTANA.Inicializar = 1
VENTANA.Campox = "CODIGO/A18;DESCRIPCION/A32;CANTIDAD/F12.1;PRECIO/F12.2;IMPORTE/F12.2"




  Dim obj As New RECORXET
  Set RS1 = obj.RS_PEDIDOS1()
  Call Carga_MSF_Recordset(RS1, "CODIGO/A18;DESCRIPCION/A32;CANTIDAD/F12.1;PRECIO/F12.2;IMPORTE/F12.2", VENTABNEW.MSF)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing




'*****************************************************'




VENTABNEW.Show 1
If VENTABNEW.APROBADO% > 0 Then

End If
Set VENTANA = Nothing
End Sub
Private Sub Command17_Click()
    Command17.Enabled = False
    CamposRecordset$ = ""
    'lista de campos de recordset separados por el palito
    For i& = 0 To List3.ListCount - 1
       List3.ListIndex = i&
       CamposRecordset$ = CamposRecordset$ & List3.TEXT & "|"
    Next i&
    CamposRecordset$ = TRIM$(CamposRecordset$)
    If Right(CamposRecordset$, 1) = "|" Then CamposRecordset$ = Mid$(CamposRecordset$, 1, Len(CamposRecordset$) - 1)

    
    Condi$ = "ID = " & XVALO(TXTID)
'    Call ACTUREGISTRO("GENECODI", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
'    Call ACTUREGISTRO("DatGrilla", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    '
    IDX& = XVALO(TXTID)
    Dim RS As New ADODB.Recordset
    SQLX$ = "SELECT *  FROM GeneCodi WHERE ID = " & XVALO(TXTID)
    RS.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS
      If XVALO(TXTID) < 1 Then
         IDX& = PROXIMO_ID("GeneCodi", "ID", "ID > 0") ' FUNCION
        .AddNew
        !ID = IDX&
      End If
      
      !TitVentana = TXTTITULO
      !ADDREGISTROS = Check5.Value
      !UtilFormula = Check6.Value
      !Borra = Check7.Value
      !Inserta = Check8.Value
      !NOTILDE = Check9.Value
      !NameRecord = TRIM$(Text10.TEXT)
      !CamposRecord = XVALO(CamposRecordset$)
      !USARECORD = Check10.Value
      !AotraGrilla = Check11.Value
      !ConsultaRec = XVALO(TXTCONSU_ESPEJO)
      !CodiGen = Text9
      !CanCampos = XVALO(Text8)
      .Update
    End With
    RS.Close: Set RS = Nothing
    '
    'ACTUALIZO EL NUEVO ID
    TXTID = IDX&

    Condi$ = "ID = " & XVALO(TXTID)
    Call BORRAREGISTROS("DATGRILLA", Condi$, REGAF&, "NOMESS")
    Dim RS1 As New ADODB.Recordset
    SQLX$ = "SELECT * FROM DatGrilla WHERE ID =" & XVALO(TXTID)
    RS1.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS1
        For i& = 0 To MSF.Rows - 1
          For J& = 1 To MSF.Cols - 1
             DATOX$ = TRIM$(MSF.TextMatrix(i&, J&))
             If DATOX$ <> "" Then
                .AddNew
                !ID = IDX&
                !FilaCol = SAFETEXT$(i&) & "|" & SAFETEXT$(J&)
                !DATO = DATOX$
                 .Update
              End If
          Next J&
        Next i&
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    '
99  Command17.Enabled = True
End Sub

Private Sub Command18_Click()


    Command18.Enabled = False
    If Check10.Value > 0 Then
      If TRIM$(Text10) = "" Then
        Call COMUNI("Falta Título de Recordset")
        OK% = -1
        GoTo 99
      End If
    End If
    '
    If Check10.Value > 0 Then
        POS1% = InStr(TXTCONSU_ESPEJO, "(NOLOCK)")
        If POS1% < 1 Then
           Call COMUNI("Falta With(NOLOCK) en consulta")
           OK% = -1
           GoTo 99
        End If
    End If
    '
    Campos$ = ""
    For i& = 1 To MSF.Cols - 1
       If TRIM$(MSF.TextMatrix(0, i&)) <> "" Then Campos$ = Campos$ & MSF.TextMatrix(0, i&) + ";"
    Next i&
    If Right$(Campos$, 1) = ";" Then Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    Campos$ = Chr$(34) & Campos$ & Chr$(34) & vbCrLf
    T1$ = "Call CARGA_ENCABEZADO(FRMZELECHO.msf2," & TRIM$(Campos$) & ", FRMZELECHO)" + vbCrLf
    T1$ = T1$ + "Call CARGA_ENCABEZADO(FRMZELECHO.msf1," & TRIM$(Campos$) & ", FRMZELECHO)" + vbCrLf

    For i& = 0 To MSF.Rows - 1
       PROPIEDADX$ = MSF.TextMatrix(i&, 0)
       Select Case PROPIEDADX$
           Case "Dejar Libre (S/N)"
               For J% = 1 To MSF.Cols - 1
                If UCase$(TRIM$(MSF.TextMatrix(i&, J%))) = "S" Then
                  VENTANA.TOTALES(J%) = MSF.TextMatrix(i&, J%)
                  T6$ = "VENTANA.DejarLibre(" & J% & ") =" & Chr$(34) & MSF.TextMatrix(i&, J%) & Chr$(34) & vbCrLf
                  T1$ = T1$ + T6$
                End If
              Next J%
       End Select
    Next i&


    T1$ = T1$ + "FRMZELECHO.CAPTION = " & Chr$(34) & TXTTITULO.TEXT & Chr$(34) & vbCrLf
    '
    Text9 = T1$
    T1$ = T1$ & String$(64, "*") & vbCrLf & vbCrLf
    Text9 = T1$
    If Check10.Value > 0 Then
       Call GENERA_AGREGA_RECORDSET(Campos$, "FRMZELECHO.MSF2")
    End If
    Text9 = Text9 & "FRMZELECHO.WIDTH = 8500 " + vbCrLf
    Text9 = Text9 & "FRMZELECHO.SHOW 1" + vbCrLf
    '
    Text9 = Text9 & "NRO = XVALO(RESP_ZELE_VECT(1))" + vbCrLf

    NRO = XVALO(RESP_ZELE_VECT(1))

99  Command18.Enabled = True
End Sub

Private Sub Command19_Click()
  TXTID = ""
End Sub

Sub LIMPIA_EDICION()
    Me.TXTID = ""
    Me.TXTIDLISTA.TEXT = ""
'    Me.TXTNOMBRELISTA.TEXT = ""
'    Me.TXTORDEN = ""
'    Me.TXTORDPRES.TEXT = ""
'    Me.TXTTIPOID = ""
    Me.CMBCAMPOS = ""
'    Me.TXTFORMATO.TEXT = ""
'    Me.CMBNOMBRECAMPO = ""
End Sub

Private Sub Command22_Click()
    Command22.Enabled = False
    '
    LISTAX$ = TRIM$(Text14)
    IDLIST& = Me.TXTIDLISTA
    'CARGO LOS DATOS UNICOS DE ENCABEZADOS
    Dim obj As New RECORXET
    Set RS1 = obj.RS_TAB_COMP_ENCAPRFUN(IDLIST&)
    MAXIMOCAMPOS = XVALO(RS1!CANTICAMPOS)
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = XVALO(RS1!ADDREGISTROS)
    VENTANA.UtilizaFormula = XVALO(RS1!USAFORMULA)
    VENTANA.HabilitaBorrar = XVALO(RS1!BORRAR)
    VENTANA.HabilitaInsertar = XVALO(RS1!INSERTAR)
    VENTANA.NoMostrarBotonTilde = XVALO(RS1!NOTILDE)
    VENTANA.TITULO = SAFETEXT$(RS1!TITURECORD)
    
    CONSU$ = SAFETEXT$(RS1!ConsultaSql)
    Check18.Value = XVALO(RS1!USARECORD)
    
    
    SQLX$ = "SELECT * FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA = " & IDLIST&
    SQLX$ = SQLX$ + " ORDER BY ENCAPRF_ID ASC"
    Set RS = READSET(SQLX$)
    With RS
    J% = 0
      LONGIT% = 0
      Do While Not .EOF
        J% = J% + 1
        ENCAX1$ = TRIM$(SAFETEXT$(!Encabezado))
        POS1% = InStr(ENCAX1$, "/")
        LONGIT% = LONGIT% + XVALO(Mid$(ENCAX1$, POS1% + 2))
        
        FORMUX1 = TRIM$(SAFETEXT$(!formula))
        CAMPEDI1$ = TRIM$(SAFETEXT$(!CAMPEDITAB))
        COLECO1$ = TRIM$(SAFETEXT$(!COLCAMECO))
        TXTTOTX1$ = TRIM$(SAFETEXT$(!TXTTOTAL))
        TOTAX1$ = TRIM$(SAFETEXT$(!Total))
        DEJLIB1$ = TRIM$(SAFETEXT$(!DejarLibre))
        If ENCAX1$ <> "" Then ENCAX$ = ENCAX$ + ENCAX1$ + ";" 'ENCABEZADO
        If FORMUX1 <> "" Then VENTANA.formula(J%) = FORMUX1
        If TRIM(UCase(CAMPEDI1$)) = "S" Then CAMPEDI$ = CAMPEDI$ & J% & ";" 'CAMPOS EDITABLES
        If COLECO1$ <> "" Then VENTANA.ColEcoCampo(J%) = COLECO1$
        If TXTTOTX1$ <> "" Then VENTANA.TxtUbicadoGrillaTotal(J%) = TXTTOTX1$
        If TOTAX1$ <> "" Then VENTANA.TOTALES(J%) = TOTAX1$
        If DEJLIB1$ <> "" Then VENTANA.DejarLibre(J%) = DEJLIB1$
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If ENCAX$ <> "" Then
      If Right(ENCAX$, 1) = ";" Then ENCAX$ = Left$(ENCAX$, Len(ENCAX$) - 1)
      VENTANA.Campox = ENCAX$
    End If
    '
    If CAMPEDI$ <> "" Then
      If Right(CAMPEDI$, 1) = ";" Then CAMPEDI$ = Left$(CAMPEDI$, Len(CAMPEDI$) - 1)
      VENTANA.CampEditables = CAMPEDI$
    End If
    '
    
    VENTANA.Inicializar = 1
    '
    RS1.Close
    Set RS1 = Nothing
    
    SQLX$ = CONSU$ + " WHERE ID_LISTA = " & IDLIST&
    Set RS1 = READSET(SQLX$)
    Call Carga_MSF_Recordset(RS1, VENTANA.Campox, VENTABNEW.MSF)
    RS1.Close
    Set RS1 = Nothing
    
'    If VENTABNEW.MSF.Rows < 2 Then VENTABNEW.MSF.Rows = 2
'    VENTABNEW.TEPRUEBA.Caption = String$(LONGIT%, "A")
'    VENTABNEW.MSF.Width = VENTABNEW.TEPRUEBA.Width
'    VENTABNEW.Width = VENTABNEW.TEPRUEBA.Width + 2500
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        Call GRABASQLDAT(SQLX$, RET_CAMPOS$, RET_TABLA$, RET_CONDI$)
        'MARCO CON MARBORRA
        Call ACTUREGISTRO(RET_TABLA$, "MARBORRA", RET_CONDI$, 1, REGAF&)
        Call TEXTOLOTES$(RET_CAMPOS$, ",")
        On Error Resume Next
        CANTICAMP = UBound(CADENATEX$)
        
        On Error GoTo 0
        If CANTICAMP > 0 Then
          For i& = 1 To VENTABNEW.MSF.Rows - 1
             If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)) = "" Then GoTo 60
             VALOX$ = IDLIST& & ","
             VALOX$ = VALOX$ & Chr(39) & LISTAX$ & Chr(39) & ","
             For k& = 1 To VENTABNEW.MSF.Cols - 1
                DATOX$ = Chr(39) & VENTABNEW.MSF.TextMatrix(i&, k&) & Chr(39)
                VALOX$ = VALOX$ + DATOX$ + ","
             Next k&
             If Right$(VALOX$, 1) = "," Then VALOX$ = Left$(VALOX$, Len(VALOX$) - 1)
             SQLX$ = "INSERT INTO " + RET_TABLA$
               PRICAMP$ = "ID_LISTA, LISTA, "
             SQLX$ = SQLX$ + " (" + PRICAMP$ + RET_CAMPOS$ + ") "
             SQLX$ = SQLX$ + " VALUES (" & VALOX$ & ")"
             FLEXCONN.Execute SQLX$
60        Next i&
        End If
        Call BORRAREGISTROS(RET_TABLA$, RET_CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    End If
    Set VENTANA = Nothing
  
99  Command22.Enabled = True

End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   IDMAX& = 1 + MAXNUMCOD&("ENCAPRFUN", "ID_LISTA")
   Me.Text20 = ""
   Me.TXTIDLISTA = IDMAX&
   ESTADO$ = "NEW"
   Command23.Enabled = True
End Sub

Private Sub Command24_Click()
    Command24.Enabled = False
    Text9 = ""
    If TRIM$(TXTCONSU_ESPEJO) = "" Then
    Call COMUNI("FALTA CONDICION PARA  MARBORRA\LA MISMA VA A SER USADA PARA EL ADDNEW")
      GoTo 99
    End If
    If TRIM$(TXTNAMETABLA) = "" Then
      Call COMUNI("FALTA SELECCIONAR TABLA")
      GoTo 99
    End If
    '
'    For i& = 1 To MSFUPD.Cols - 1
'        CAMPO1$ = TRIM$(MSFUPD.TextMatrix(1, i&))
'        For J& = 0 To LSTMOSTARCMP2.ListCount
'            LSTMOSTARCMP2.ListIndex = J&
'            CAMPOLEIDO$ = LSTMOSTARCMP2.TEXT
'            If CAMPO1$ = CAMPOLEIDO$ Then
'             LSTMOSTARCMP2.Selected(J&) = True
'             Exit For
'            End If
'        Next J&
'    Next i&
    '
T1$ = "DIM MSF AS MSFLEXGRID" & vbCrLf
T1$ = T1$ + "SET MSF = MSF11" & vbCrLf & vbCrLf
T1$ = T1$ + "CONDI$ = " & Chr(34) & TXTCONSU_ESPEJO & Chr(34) & vbCrLf
T1$ = T1$ + "Call ACTUREGISTRO(" & Chr(34) & TXTNAMETABLA.TEXT & Chr(34)

T1$ = T1$ + "," & Chr(34) & "MARBORRA" & Chr(34)
T1$ = T1$ + ", CONDI$ "
T1$ = T1$ + ", 1"
T1$ = T1$ + ", REGAF "
T1$ = T1$ + "," & Chr(34) & "NOMESS" & Chr(34) & ")" & vbCrLf & vbCrLf

T1$ = T1$ + "SQLX$ = " & Chr(34) & "SELECT * FROM " & TXTNAMETABLA & Chr(34) & " " & vbCrLf
T1$ = T1$ + "SQLX$ = SQLX$ + " & Chr(34) & " WHERE " & Chr(34) & vbCrLf
T1$ = T1$ + "SQLX$ = SQLX$ + " & Chr(34) & " " & TXTCONSU_ESPEJO & Chr(34) & vbCrLf
T1$ = T1$ + "Dim RS As ADODB.Recordset" & vbCrLf
T1$ = T1$ + "Set RS = New ADODB.Recordset" & vbCrLf
T1$ = T1$ + "25      On Error Resume Next" & vbCrLf
T1$ = T1$ + "RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText" & vbCrLf
T1$ = T1$ + "If Err Then" & vbCrLf
T1$ = T1$ + " On Error GoTo 0" & vbCrLf
T1$ = T1$ + " Call CapturaErrorOpen" & vbCrLf
T1$ = T1$ + " GoTo 25" & vbCrLf
T1$ = T1$ + "End If" & vbCrLf

T1$ = T1$ + "  With RS" & vbCrLf
T1$ = T1$ + "      FOR REG& = 1 TO MSF.ROWS - 1" & vbCrLf
T1$ = T1$ + "       .AddNew" & vbCrLf
For i& = 1 To MSFUPD.Cols - 1
       FORMAT1$ = MSFUPD.TextMatrix(0, i&)
       Call TEXTOLOTES$(FORMAT1$, "/")
       On Error Resume Next
       FORMAT1$ = UCase(TRIM$(CADENATEX$(2)))
       If Left(FORMAT1$, 1) = "F" Then
          T1$ = T1$ & "          !" & MSFUPD.TextMatrix(1, i&) & "=" & "XVALO(MSF.TextMatrix(REG&, " & i& & "))" & vbCrLf
       Else
          T1$ = T1$ & "          !" & MSFUPD.TextMatrix(1, i&) & "=" & "MSF.TextMatrix(REG&, " & i& & ")" & vbCrLf
       End If
    Next i&
T1$ = T1$ + "       .UPDATE" & vbCrLf
T1$ = T1$ + "      NEXT REG&" & vbCrLf
T1$ = T1$ + "  END WITH" & vbCrLf & vbCrLf
T1$ = T1$ + "  RS.CLOSE" & vbCrLf
T1$ = T1$ + "  SET RS= NOTHING" & vbCrLf & vbCrLf

T1$ = T1$ + "CONDI$ =  CONDI$ + " & Chr(34) & " AND  MARBORRA = 1  " & Chr(34) & vbCrLf
T1$ = T1$ + "Call BORRAREGISTROS(" & Chr(34) & TXTNAMETABLA.TEXT & Chr(34) & ", CONDI$," & "REGAF&," & Chr(34) & "NOMESS" & Chr(34) & ")"


Text9 = T1$
            
99   Command24.Enabled = True
End Sub

Private Sub Command24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 1 Then
     If Button = 2 Then
       Call AYUDA("GEN.COD.Update")
     End If
   End If

End Sub

Sub AYUDA(DATO$)
   Select Case DATO$
     Case "GEN.COD.Update"
       TXTO$ = "GENERA CODIGO PARA GRABAR DESDE UNA GRILLA " & vbCrLf
       TXTO$ = TXTO$ + "A LA TABLA CORRESPONDIENTE  " + vbCrLf
       TXTO$ = TXTO$ + "1.-ESCRIIR EN EL CASILLERO CAMPOS (TXTCAMPOSUPDATE)" + vbCrLf
       TXTO$ = TXTO$ + "LOS CAMPOS QUE ESTAN EN LAS PROPIEDADES DE LA GRILLA PARA CARGAR ENCBEZADO" + vbCrLf

       TXTO$ = TXTO$ + "2.-CLICK EN EL BOTON ADD CAMPOS" + vbCrLf
       TXTO$ = TXTO$ + "3.-CLICK EN CADA COLUMNA DE LA GRILLA PARA SELECCIONAR LOS CAMPOS DE LA TABLA A ENLAZAR" + vbCrLf
       TXTO$ = TXTO$ + "4.-CLICK EN BOTON GEN. COD. UPDATE" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       MsgBox TXTO$
     Case "Crear Archivo Estrucutura"
       TXTO$ = "GENERA ESTRUCTURA DE ARCHIVO TIPO PRF Y RFS" & vbCrLf
       TXTO$ = TXTO$ + "INGRESAR EN LA SOLAPA TAB2  " + vbCrLf
       TXTO$ = TXTO$ + "0.-CLIC EN NEW - PONERLE UN NOMBRE Y UNA DESCRIPCION A LA ESTRUCTURA" + vbCrLf
       TXTO$ = TXTO$ + "1.-INGRESAR EL ENCABEZADO Y FORMATO DE LOS CAMPOS)" + vbCrLf
       TXTO$ = TXTO$ + "2.-AGREGAR LOS CAMPOS A MOSTRAR  Y LA TABLA DATASQL " + vbCrLf
       TXTO$ = TXTO$ + "3.-EJ.: SELECT CAMPO1, CAMPO2,.... FROM DATASQL WITH(NOLOCK)" + vbCrLf
       TXTO$ = TXTO$ + " -ESTO EN EL CASILLERO ROTULADO CONSULTA PARA RECORDSET" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + "4.-CAMPOS ECO" + vbCrLf
       TXTO$ = TXTO$ + " -Los campos Eco deben estar editables." + vbCrLf
       TXTO$ = TXTO$ + " -SE PUEDE INVOCAR A UNA TABLA SQL O A UNA ESTRUCTURA" + vbCrLf
       TXTO$ = TXTO$ + " -La referencia del eco debe ir donde va a ir el codigo en esa " + vbCrLf
       TXTO$ = TXTO$ + " -columna se va agregar un boton automatico, y en la columna " + vbCrLf
       TXTO$ = TXTO$ + " -siguiente va a ir la descripcion" + vbCrLf
       TXTO$ = TXTO$ + Space$(10) + String$(40, "-") + vbCrLf
       TXTO$ = TXTO$ + " -en el caso de invocar como eco a una estructura se debe" + vbCrLf
       TXTO$ = TXTO$ + " -poner el siguiente codigo " + vbCrLf
       TXTO$ = TXTO$ + " -CAMPO1;CAMPO2;DATASQL;LISTA = 'CODIMOTI'" + vbCrLf
       TXTO$ = TXTO$ + " -con el mismo criterio que en el paso anterior." + vbCrLf
       TXTO$ = TXTO$ + " -poner el siguiente codigo " + vbCrLf
       
       MsgBox TXTO$
     Case "ATAJOS EN VENTABNEW (VENTANA DE CARGA)"
       TXTO$ = "ATAJOS EN VENTABNEW (VENTANA DE CARGA)" & vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI SE OPRIME F2 EN UN CAJA DE TEXTO (CELDA EDITABLE CON ECO) " + vbCrLf
       TXTO$ = TXTO$ + "     DESPLEGA LA LISTA DE SELECCION " + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI SE OPRIME F3 SOBRE UNA CELDA REPITE EL VALOR INMEDIATO SUPERIOR" + vbCrLf + vbCrLf
       TXTO$ = TXTO$ + "SI LA CELDA ES EDITABLE Y TIENE FORMATO D8 " + vbCrLf
       TXTO$ = TXTO$ + "   CUANDO ESTA LA CAJA DE TEXTO VISIBLE SI SE REALIZA DOBLE CLIC " + vbCrLf
       TXTO$ = TXTO$ + "   MUESTRA LA LISTA DE SELECCION DE FECHA" + vbCrLf
       MsgBox TXTO$
     
   End Select
End Sub
Private Sub Command25_Click()
   Command25.Enabled = False
    
   IDLISTAX& = XVALO(TXTIDLISTA)
   If IDLISTAX& < 1 Then
     Call COMUNI("Falta Seleccionar Id Lista")
     GoTo 99
   End If
   '
   If XVALO(Text11) < 1 Then
      Call COMUNI("Falta Ingresar Cantidad de Campos")
      Call PINTATEXT(Text11)
      GoTo 99
   End If
   
   If ESTADO$ = "NEW" Then
      IDLISTAX& = 1 + MAXNUMCOD&("ENCAPRFUN", "ID_LISTA")
   End If
   '
   CODICON$ = TRIM$(Text20)
   If CODICON$ = "" Then
      Call COMUNI("Falta Ingresar Codigo Alfa de Consulta")
      Text20.Locked = False
      Text20.BackColor = vbWhite
      Call PINTATEXT(Text20)
      GoTo 99
   End If
   
   REFECON$ = TRIM$(Text14)
   If REFECON$ = "" Then
      Call COMUNI("Falta Ingresar Referencia de Consulta")
      Text14.Locked = False
      Text14.BackColor = vbWhite
      Call PINTATEXT(Text14)
      GoTo 99
   End If
   '
   For i& = 0 To MSF4.Cols - 1
       ENK$ = TRIM$(MSF4.TextMatrix(0, i&))
       If ENK$ = "" Then
          CANTIENCA% = i&
          Text11 = i&
          MSF4.Cols = i&
          Call COMUNI("Atención !!! Se ha Ajustado la Cantidad de Columnas Por Que se Detectaron Columnas Vacias")
          Exit For
       End If
   Next i&
   '
   
   If TRIM(Text13) = "" Then
     Call COMUNI("FALTA CONDICION DE TABLA")
     Call PINTATEXT(Text13)
     GoTo 99
   End If
   
   'Left(CAMPEDI$, 1) < 2
   If MSF4.TextMatrix(2, 1) <> "S" Then
       Call COMUNI("ATENCIÓN CUANDO EL CAMPO 1 NO ESTA EDITABLE EL SISTEMA AUTONUMERA\ESTE CAMPO PONERLO COMO NUMERICO F...\CONFIGURAR QUE NO SE PUEDA BORRAR NI INSERTAR\PARA QUE NO SE AUTONMERE Y CAMBIE LOS CODIGOS ORIGINALES")
   End If
   SQLX$ = "SELECT * FROM  ENCAPRFUN  "
   SQLX$ = SQLX$ + " WHERE ID_LISTA = " & IDLISTAX&
   SQLX$ = SQLX$ + " AND (MARBORRA = 0 OR MARBORRA IS NULL)"
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
        If .EOF Then
          .AddNew
        End If
        !ID_LISTA = IDLISTAX&
        !LISTA = Left$(Text20, 64)
        !Descripcion = Left$(Text14, 64)
        !ADDREGISTROS = Check13.Value
        !USAFORMULA = Check14.Value
        !BORRAR = Check15.Value
        !INSERTAR = Check16.Value
        !NOTILDE = Check17.Value
        !CANTICAMPOS = XVALO(Text11)
        !CodiEmpr = CodiEmp%
        !USARECORD = Check18.Value
        !FECHEMISI = FETEXCOR1$(HOY(HO$))
        !TITURECORD = Text12
        !ConsultaSql = Text13
        .Update
   End With
   RS.Close
   Set RS = Nothing
   Call ACTUREGISTRO("ENCAPRF", "MARBORRA", "ID_LISTA = " & IDLISTAX&, 1, REGAF&, "NOMESS")
   '
   SQLX$ = "SELECT * FROM  ENCAPRF  "
   SQLX$ = SQLX$ + " WHERE ENCAPRF_ID = " & IDLISTAX&
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS1
    For i& = 1 To MSF4.Cols - 1
       .AddNew
       !ID_LISTA = IDLISTAX&
       !CodiEmpr = CodiEmp%
       !MARBORRA = 0
       !Encabezado = MSF4.TextMatrix(0, i&)
       !formula = MSF4.TextMatrix(1, i&)
       !CAMPEDITAB = MSF4.TextMatrix(2, i&)
       !COLCAMECO = MSF4.TextMatrix(3, i&)
       !TXTTOTAL = MSF4.TextMatrix(4, i&)
       !Total = MSF4.TextMatrix(5, i&)
       !DejarLibre = MSF4.TextMatrix(6, i&)
    Next i&
    .Update
             
   End With
   RS1.Close
   Set RS1 = Nothing
   Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX& & " AND MARBORRA = 1", REGAF&, "NOMESS")
   ESTADO$ = "EDIT"
   '

    Call GENERACODIGO_ARCHIVO_ESTRUCTURA(Text20)

99 Command25.Enabled = True

End Sub
Sub GENERACODIGO_ARCHIVO_ESTRUCTURA(CODICONSULTA$)
   '
   'VALIDO SI EXISTE EN ALGUNA TABLA DE CUALQUIER SISTEMA
   EXISTE_EN_LISTA% = 0
   If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
     ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + "CONTROLESTSQL.txt"
     NARCHI2% = FILEOPEN%(ASALI$, 1, 128) ' ABRE EL ARCHIVO DE ESCRITURA
     Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(Mid$(XX$, 1, 24)))
       
       If XX1$ = TRIM$(UCase$(Text20)) Then
          Descripcion$ = TRIM$(UCase$(Mid$(XX$, 26, 32)))
          EMPRESA$ = TRIM$(UCase$(Mid$(XX$, 60, 24)))
          txt1$ = "Estructura: " & Text20 & " Referencia: " & Descripcion$ & " Empresa: " & EMPRESA$ & "\Anteriormente Creada"
          OPX% = COMALTER%(txt1$ & "\\Cancelar\Generar Código Para Escribir en Módulo EscribeEstrucSql.bas")
          EXISTE_EN_LISTA% = 1
          If OPX% < 2 Then Exit Sub
       End If
    Loop
    '
    If EXISTE_EN_LISTA% < 1 Then
          Call REPLA(TX$, TRIM$(UCase$(Text20)), 1, 24)
          Call REPLA(TX$, TRIM$(UCase$(Text14)), 26, 32)
          Call REPLA(TX$, TRIM$(UCase$(Left$(EMPREAC$, 24))), 60, 24)
          Call REPLA(TX$, TRIM$(UCase$(Left$(USERTER$, 24))), 86, 24)
          
          Close #NARCHI2%
          NARCHI3% = FILEOPEN%(ASALI$, 3, 128) 'ABRE PARA ESCRIBIR AL FINAL
          Print #NARCHI3%, , TX$
          Close #NARCHI3%
    End If
    '
    On Error Resume Next
    Close #NARCHI2%
    On Error GoTo 0
   End If
   
   Text15 = "'|INICIO " & CODICONSULTA$ & vbCrLf
   Text15 = Text15 + "'" & vbCrLf
   Text15 = Text15 + "CONDI$ = " & Chr$(34) & " LISTA =   '" & CODICONSULTA$ & "'" & vbCrLf
   Text15 = Text15 + "EXILISTA% = CantRegistros&("
   Text15 = Text15 & Chr$(34) & "ENCAPRFUN" & Chr$(34)
   Text15 = Text15 + ", CONDI$ ," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "If EXILISTA% > 0 Then" & vbCrLf
   Text15 = Text15 + "  FE% = " & HOY(HO$) & vbCrLf
   Text15 = Text15 + "  FEBD% = CVINT(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "FECHEMISI" & Chr$(34) & ", CONDI$, " & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "  GENERAREST% = 0" & vbCrLf
   Text15 = Text15 + "  If FEBD% < FE% Then" & vbCrLf
   Text15 = Text15 + "    IDLISTAX& = XVALO(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "ID_LISTA" & Chr$(34) & ", CONDI$," & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "    Call BORRAREGISTROS(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "ID_LISTA = " & Chr$(34) & " & IDLISTAX&,REGAF&," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "    Call BORRAREGISTROS(" & Chr$(34) & "ENCAPRF" & Chr$(34) & "," & Chr$(34) & "ID_LISTA = " & Chr$(34) & " & IDLISTAX&,REGAF&," & Chr$(34) & "NOMESS" & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "    GENERAREST% = 1" & vbCrLf
   Text15 = Text15 + "  END IF  " & vbCrLf
   Text15 = Text15 + "ELSE" & vbCrLf
   Text15 = Text15 + "  GENERAREST% = 1" & vbCrLf
   Text15 = Text15 + "  CONDI$ =  " & Chr$(34) & " ID_LISTA  > 0 " & Chr$(34) & vbCrLf
   Text15 = Text15 + "  IDLISTAX& = 1 + XVALO(VALOBADA(" & Chr$(34) & "ENCAPRFUN" & Chr$(34) & "," & Chr$(34) & "MAX(ID_LISTA)" & Chr$(34) & ",CONDI$," & Chr$(34) & "NOMESS" & Chr$(34) & "))" & vbCrLf
   Text15 = Text15 + "END IF  " & vbCrLf
   Text15 = Text15 + "If GENERAREST% > 0 Then  " & vbCrLf
   
   '***********************************************************************************************************************************

   Text15 = Text15 + "  SQLX$ = " & Chr$(34) & "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE " & Chr$(34) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)" & Chr$(34) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "VALUES ( " & Chr$(34) & " & IDLISTAX&" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ",'" & Left$(Text20, 64) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ",'" & Left$(Text14, 64) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "," & Check13.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & "," & Check14.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check15.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check16.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check17.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & XVALO(Text11) & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & CodiEmp% & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ," & Check18.Value & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & CVDAT(HOY(HO$)) & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & Text12 & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & " ,'" & Text13 & "'" & vbCrLf
   Text15 = Text15 + "  SQLX$ = SQLX$ + " & Chr$(34) & ")" & vbCrLf
   Text15 = Text15 + "  FLEXCONN.Execute (SQLX$)" & vbCrLf
      '
   For i& = 1 To MSF4.Cols - 1
        Text15 = Text15 + "    SQLX$ = " & Chr$(34) & "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) " & Chr$(34) & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " VALUES ( " & Chr$(34) & " & IDLISTAX&" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " ," & CodiEmp% & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " ,0" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(0, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(1, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(2, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(3, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(4, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(5, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " , '" & MSF4.TextMatrix(6, i&) & "'" & vbCrLf
        Text15 = Text15 + "    SQLX$ = SQLX$ + " & Chr$(34) & " )" & vbCrLf
        Text15 = Text15 + "    FLEXCONN.Execute (SQLX$)" & vbCrLf & vbCrLf
        Text15 = Text15 + Chr$(39) & vbCrLf
   Next i&
   Text15 = Text15 + "End If " & vbCrLf
   Text15 = Text15 + "'|FIN " & CODICONSULTA$ & vbCrLf
   '

   OPX% = COMALTER%("Desea Escribir el Modulo 'EstrucSQL'\\No, No Escribir\Si, Escribir (Cierre Los Fuentes)")
   If OPX% = 2 Then
      Call EscribeEstrucSql(CODICONSULTA$, Text15)
   End If
   
   
End Sub
   '
Private Sub Command26_Click()
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  
  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  NRO = XVALO(RESP_ZELE_VECT(1))
  
  If NRO > 0 Then
    ESTADO$ = "EDIT"
    TXTIDLISTA = NRO
    Text20 = RESP_ZELE_VECT(2)
  End If
End Sub

Private Sub Command27_Click()
   IDLISTAX& = XVALO(Me.TXTIDLISTA)
   If IDLISTAX& > 0 Then
     OPX% = COMALTER%("Confirma Eliminar Definitivamente esta Estructura\\Cancelar\Eliminar")
     If OPX% = 2 Then
        Condi$ = " ID_LISTA = " & IDLISTAX&
        Call BORRAREGISTROS("ENCAPRFUN", Condi$, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", Condi$, REGAF&, "NOMESS")
     End If
   End If
End Sub

Private Sub LSTMOSTARCMP2_Click()
     Call TEXTOLOTES$(LSTMOSTARCMP2.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     Me.MSFUPD.TextMatrix(1, COLU1&) = LSTMOSTARCMP2.TEXT
     LSTMOSTARCMP2.Visible = False

End Sub

Private Sub LSTMOSTARCMP2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       LSTMOSTARCMP2.Visible = False
    End If
End Sub


Private Sub LSTTABLAS2_Click()
    TAPRIN$ = LSTTABLAS2.TEXT
    PREFPRIN$ = ""
    If TAPRIN$ = "" Then Exit Sub
    LSTMOSTARCMP2.Clear
    LSTMOSTARCMP2.AddItem ""
    Me.TXTNAMETABLA = TAPRIN$
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           'CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           LSTMOSTARCMP2.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           'Combo4.AddItem AUX.Fields(i& - 1).Name
        End If
    Next i&

End Sub

Private Sub mnuArmaSql_Click()
  ARMASQLX.Show 1
  
End Sub

Private Sub MnuAutonumTabAux_Click()
       Call COMUNI("CUANDO EL CAMPO 1 NO ESTA EDITABLE EL SISTEMA AUTONUMERA\ESTE CAMPO PONERLO COMO NUMERICO F...\CONFIGURAR QUE NO SE PUEDA BORRAR NI INSERTAR\PARA QUE NO SE AUTONMERE Y CAMBIE LOS CODIGOS ORIGINALES")

End Sub

Private Sub mnuAyudaVentabnew_Click()
   Call AYUDA("ATAJOS EN VENTABNEW (VENTANA DE CARGA)")
End Sub

Private Sub mnuCreaEstructura_Click()
    Call AYUDA("Crear Archivo Estrucutura")
   
End Sub

Private Sub mnuReverUpd_Click()
    FRMREVUPD.Show 1
End Sub

Private Sub mnuUpdate_Click()
   Call AYUDA("GEN.COD.Update")
End Sub

Private Sub MSF_RowColChange()
    MSFESPEJO.Rows = MSF.Rows
    MSFESPEJO.Cols = MSF.Cols
End Sub

Private Sub MSF_Scroll()
   If TXTMOVIL.Visible = True Then
     Call TXTMOVIL_LostFocus
     TXTMOVIL.Visible = False
   End If
End Sub

Private Sub CAMAUXI_Click()
  Call ARMASELFIELD
End Sub

Private Sub chkUnionAll_Click()

    LSTTABLASEC.ListIndex = -1
    UNIONALL$ = ""
    CMBJUNTAR.TEXT = ""
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
    
    Call ARMA_CONSULTX
    
End Sub

Private Sub CMBJUNTAR_Click()
   Call ARMA_CONSULTX
End Sub


'
Private Sub CMDBORRAR_Click()
    TXTCONSULTA.TEXT = ""
End Sub

Private Sub Combo2_Click()
    Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & Combo2.TEXT))
    For i& = 1 To AUX.Fields.Count
        Combo3.AddItem AUX.Fields(i& - 1).Name
    Next i&
End Sub

Private Sub Combo6_Click()

    Text5 = Combo6.TEXT
    
End Sub

Private Sub Command10_Click()
'    Static CTX%
'    If CTX% < 1 Then
'       Call ARMALISTARCH(LUDAT$, "*.RFS")
'       CTX% = 1
'    End If
'    Call SELECHO("INDIRFS")   ' , "DIRECTORIO RFS")
'    VDEVU$ = TRIM$(VALACT1$("INDIRFS"))
'    If VDEVU$ <> "" Then
'       Text5 = VDEVU$
'    End If

End Sub

Private Sub Command11_Click()
    If Label1 <> "" Then
       If COMALTER%("¿ Confirma Eliminar la Consulta '" & TRIM$(Label1) & "' ?\  \(Esta Transaccion no Podrá Revertirse)\\No, Conservarla\Si, Eliminarla") = 2 Then
          Condi$ = "COD_QUERY = '" & Label1.Caption & "'"
          Call BORRAREGISTROS("SQL_QUERIES", Condi$, REGAF&, "NOMESS")
          
          Call ABRETACON
          STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
          Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
          If Not (MIRS.EOF And MIRS.BOF) Then
             MIRS.Delete
             DEFTREP07.Label4 = ""
             DEFTREP07.Label9 = ""
             Call DEFTREP07.VerConsul
             MIRS.Close
             Unload Me
          End If
       End If
    End If
End Sub

Private Sub Command12_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 3) <> "OR " Then
          List2.List(i%) = "OR " & List2.TEXT
         End If
         Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call GENERAR_CODIGO(OK%)
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
    'INTERVENCION OR
    i% = List2.ListIndex
    If i% >= 0 Then
        If Right(TRIM$(List2.TEXT), 1) <> ")" Then
          List2.List(i%) = List2.TEXT & ")"
        End If
        Call ARMA_CONSULTX
    End If
End Sub

Private Sub Command15_Click()
    'INTERVENCION OR

    i% = List2.ListIndex
    If i% >= 0 Then
        If Left(TRIM$(List2.TEXT), 1) <> "(" Then
          List2.List(i%) = "(" & List2.TEXT
        End If
        Call ARMA_CONSULTX
    End If
End Sub



Private Sub Command7_Click()
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
      If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
        TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
        CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
        CONDIVALO$ = AJUSTI$(Text7, 512)
        'If TRIM$(TAVALO$) <> "" Then
        '  If TRIM$(CAVALO$) <> "" Then
            If TRIM$(CONDIVALO$) <> "" Then
              CAMPOSE$ = AJUSTI$(List1.TEXT, 168) + TAVALO$ + CAVALO$ + CONDIVALO$
              List1.RemoveItem (i&)
              List1.AddItem CAMPOSE$, i&
              List1.Selected(i&) = True
            End If
        '  End If
        'End If
      End If
    End If
    '
    For i& = 1 To LSTMOSTARCMP.ListCount
       If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          LSTMOSTARCMP.RemoveItem (i& - 1)
          LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
          LSTMOSTARCMP.Selected(i& - 1) = True
       End If
    Next i&
    '
    For i& = 1 To CAMAUXI.ListCount
       If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          CAMAUXI.RemoveItem (i& - 1)
          CAMAUXI.AddItem CAMPOSE$, i& - 1
          CAMAUXI.Selected(i& - 1) = True
       End If
    Next i&
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command9_Click()
    '
    LSTTABLASEC.ListIndex = (-1)
    Combo4.Clear
    Combo5.Clear
    TAPRIN$ = LSTTABLAS.TEXT
    If TAPRIN$ = "" Then Exit Sub
    '
10  For U& = 1 To LSTMOSTARCMP.ListCount
       If Left$(LSTMOSTARCMP.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          LSTMOSTARCMP.RemoveItem (U& - 1)
          GoTo 10
       End If
    Next U&
    '
20  For U& = 1 To List1.ListCount
       If Left$(List1.List(U& - 1), Len(TAPRIN$)) = TAPRIN$ Then
          List1.RemoveItem (U& - 1)
          GoTo 20
       End If
    Next U&
    '
    Call LSTTABLAS_Click     ' PARA REFRESCAR EL COMBO4
    '
End Sub

Private Sub label1_Change()
    Dim FISELE$(128)
    List2.Clear
    
    CANFI% = 0
    ESTAENSQL% = 0
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK) where COD_QUERY = '" & Label1.Caption & "'"
    Set RS = READSET(SQLX$)
    If Not (RS.EOF And RS.BOF) Then
       Label9 = SAFETEXT$(RS!DES_QUERY)
       TTXX$ = SAFETEXT$(RS!PAR_QUERY)
       ESTAENSQL% = 1
    End If
    If ESTAENSQL% = 1 Then
      RS.Close
      Set RS = Nothing
      GoTo 30
    End If
    
    STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Label1.Caption & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_QUERY
       TTXX$ = MIRS!PAR_QUERY
       '
30     While Len(TTXX$) > 0
          PPXX% = InStr(TTXX$, Chr$(124))
          If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
          COMAI$ = Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 1)
          '
          If Left$(COMAI$, 10) = "DATABASE =" Then
              Call CMBBASE_CLICK
              CMBBASE = UCase$(TRIM$(Mid$(COMAI$, 11)))
              GoTo 19
            ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
              TASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLAS.ListCount
                 If UCase$(LSTTABLAS.List(KKI% - 1)) = TASELE Then
                   LSTTABLAS.Selected(KKI% - 1) = True
                   Call LSTTABLAS_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "SELTASEC =" Then
              TASELSEC$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLASEC.ListCount
                 If UCase$(LSTTABLASEC.List(KKI% - 1)) = TASELSEC Then
                   LSTTABLASEC.Selected(KKI% - 1) = True
                   Call LSTTABLASEC_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "MATCHFI1 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo4.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "MATCHFI2 =" Then
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              On Error Resume Next
              Combo5.TEXT = TRIM$(Left$(CASELE$, 64))
              On Error GoTo 0
            ElseIf Left$(COMAI$, 10) = "SELFIELD =" Then
              CANFI% = CANFI% + 1
              FISELE$(CANFI%) = TRIM$(Mid$(COMAI$, 11))
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
              For KKI% = 1 To LSTMOSTARCMP.ListCount
                 If UCase$(TRIM$(Left$(LSTMOSTARCMP.List(KKI% - 1), 64))) = CASELE$ Then
                   LSTMOSTARCMP.RemoveItem KKI% - 1
                   LSTMOSTARCMP.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   LSTMOSTARCMP.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
              For KKI% = 1 To CAMAUXI.ListCount
                 If UCase$(TRIM$(Left$(CAMAUXI.List(KKI% - 1), 64))) = CASELE$ Then
                   CAMAUXI.RemoveItem KKI% - 1
                   CAMAUXI.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   CAMAUXI.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
              List2.AddItem TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "TIPOJOIN =" Then
              CMBJUNTAR.TEXT = TRIM$(Mid$(COMAI$, 11))
            ElseIf Left$(COMAI$, 10) = "UNIONALL =" Then
              chkUnionAll.Value = 1
              SegundaTabla$ = TRIM$(Mid$(COMAI$, 11))
              For idTbSec% = 0 To LSTTABLASEC.ListCount - 1
                    If TRIM$(LSTTABLASEC.List(idTbSec%)) = SegundaTabla$ Then
                        LSTTABLASEC.ListIndex = idTbSec
                        UNIONALL$ = SegundaTabla$
                        TablaUnionOk% = 1
                        Exit For
                    End If
              Next idTbSec%
              If TablaUnionOk% = 0 Then
                chkUnionAll.Value = 0
                UNIONALL$ = ""
              End If
            Else
          End If
19     Wend
       '
       List1.Clear
       For KKI% = 1 To CANFI%
          List1.AddItem FISELE$(KKI%)
       Next KKI%
       '
       Call ARMA_CONSULTX
       Exit Sub
       '
    End If
    '
    'SI NO ESTABA EN SQL ENTRO EN MDB Y CIERRA EL RECORDSET
    If ESTAENSQL% = 0 Then
      MIRS.Close
      Set MIRS = Nothing
    End If

    Call LSTTABLAS_Click
    '
End Sub

Private Sub Command1_Click()
   If CMBCAMPOS.TEXT = "" Or CMBOPERADOR = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   If Len(TRIM$(CMBCAMPOS.TEXT)) > 24 Then
      MsgBox "Nombre de Campo Demasiado Largo"
      Exit Sub
   End If
   SQLY$ = AJUSTI$(CMBCAMPOS.TEXT, 24) + " " + AJUSTI$(CMBOPERADOR.TEXT, 8)
   VALOCO$ = TRIM$(Combo1.TEXT)
   '
   TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
   If (TPC = 52 Or TPC = 62) And VALOCO$ = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   '
   ' ESTO ES PARA QUE PERMITA COMPARAR VALORES DEL RS ENTRE SÍ
   For KKU& = 1 To CMBCAMPOS.ListCount
      If UCase$(CMBCAMPOS.List(KKU& - 1)) = UCase$(VALOCO$) Then GoTo 10
   Next KKU&
   '
   If TPC <> 52 And TPC <> 62 Then VALOCO$ = "'" + VALOCO$ + "'"
   '
10 SQLY$ = SQLY$ + " " + VALOCO$

   List2.AddItem SQLY$
   Call ARMA_CONSULTX
   Exit Sub
   
      If KLIK% > 0 Then GoTo 25
      SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " Where " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      KLIK% = KLIK% + 1
      TXTCONSULTA.TEXT = SQLX$
      Exit Sub
      
25    SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " AND " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      TXTCONSULTA.TEXT = SQLX$
      
End Sub

Function TIPODECAMPO(NTABLA$, NCAMPO$)
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    STRSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO = MIRS!xtype
    End If

End Function

Private Sub CMDCOPIAR_Click()
   Clipboard.SetText Mid$(TXTCONSULTA.TEXT, TXTCONSULTA.SelStart + 1, Text1.SelLength)
End Sub

Private Sub CORTAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

Private Sub Command2_Click()
   If CMBCAMPOS.TEXT <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CMBCAMPOS.TEXT & " AS VALOR FROM " & LSTTABLAS.TEXT & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CMBCAMPOS.TEXT & " ASC"
      '
      Combo1.Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      On Error Resume Next
      Set VAPOSI = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Ejecutar Consulta - Guarde y Re-Intente")
         Exit Sub
      End If
      On Error GoTo 0
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1.AddItem !valor
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
      '
   End If
   
End Sub

Private Sub Command3_Click()
10  COMO$ = InputBox$("Leyenda (Prompt)", "Configuración Consultas SQL", CMBCAMPOS.TEXT)
    If COMO$ = "" Then GoTo 10
    Combo1.Clear
    Combo1.AddItem ("¿" + COMO$ + "?")
    Combo1.TEXT = "¿" + COMO$ + "?"
End Sub

Private Sub Command4_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& > 1 Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i& - 2
         List1.Selected(i& - 2) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
End Sub

Private Sub Command5_Click()
   For i& = 1 To List1.ListCount
     If List1.Selected(i& - 1) = True Then
       If i& < List1.ListCount Then
         CAMPOSE$ = List1.List(i& - 1)
         List1.RemoveItem (i& - 1)
         List1.AddItem CAMPOSE$, i&
         List1.Selected(i&) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next i&
   Call ARMA_CONSULTX
End Sub

Private Sub Command6_Click()
    '
    MODIFENCA% = 0
    i& = List1.ListIndex
    If i& >= 0 Then
       TOTALI$ = "0": If Check1.Value = 1 Then TOTALI$ = "1"
       AGRUPA$ = "0": If Check2.Value = 1 Then AGRUPA$ = "1"
       OCULTA$ = "0": If Check4.Value = 1 Then OCULTA$ = "1"
       MODOSOR$ = "    "
       If Option2.Value = True Then MODOSOR$ = "ASC "
       If Option3.Value = True Then MODOSOR$ = "DESC"
       CAMPOSE$ = AJUSTI$(List1.List(i&), 168)
       While Len(CAMPOSE$) < 168: CAMPOSE$ = CAMPOSE$ + " ": Wend
       Mid$(CAMPOSE$, 65, 8) = AJUSTI$(Text1, 8)
       Mid$(CAMPOSE$, 73, 1) = OCULTA$
       Mid$(CAMPOSE$, 75, 1) = AGRUPA$
       Mid$(CAMPOSE$, 76, 1) = TOTALI$
       Mid$(CAMPOSE$, 77, 4) = MODOSOR$
       Mid$(CAMPOSE$, 81, 64) = AJUSTI$(Text2, 64)
       Mid$(CAMPOSE$, 145, 16) = AJUSTI$(Text5, 16)
       Mid$(CAMPOSE$, 161, 8) = AJUSTI$(Text6, 8)
       '
       If UCase$(Left$(List1.TEXT, 8)) = "VALOBADA" Or TRIM$(Text7) <> "" Then
         TAVALO$ = AJUSTI$(Combo2.TEXT, 64)
         CAVALO$ = AJUSTI$(Combo3.TEXT, 64)
         CONDIVALO$ = AJUSTI$(Text7, 512)
         If TRIM$(CONDIVALO$) <> "" Then
            While Len(CAMPOSE$) < (168 + 64 + 64 + 512): CAMPOSE$ = CAMPOSE$ + " ": Wend
            Mid$(CAMPOSE$, 169, 64) = TAVALO$
            Mid$(CAMPOSE$, 233, 64) = CAVALO$
            Mid$(CAMPOSE$, 297, 512) = CONDIVALO$
         End If
       End If
       List1.RemoveItem (i&)
       List1.AddItem CAMPOSE$, i&
       List1.Selected(i&) = True
       '
       For i& = 1 To LSTMOSTARCMP.ListCount
          If TRIM$(Left$(LSTMOSTARCMP.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             LSTMOSTARCMP.RemoveItem (i& - 1)
             LSTMOSTARCMP.AddItem CAMPOSE$, i& - 1
             LSTMOSTARCMP.Selected(i& - 1) = True
          End If
       Next i&
       '
       For i& = 1 To CAMAUXI.ListCount
          If TRIM$(Left$(CAMAUXI.List(i& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
             CAMAUXI.RemoveItem (i& - 1)
             CAMAUXI.AddItem CAMPOSE$, i& - 1
             CAMAUXI.Selected(i& - 1) = True
          End If
       Next i&
       '
    End If
    '
    Call ARMA_CONSULTX
End Sub

Private Sub Command8_Click()
   'ADAT$ = Combo2.TEXT
   ADAT$ = Label1.Caption
   Call CARGADEBASE(ADAT$, TEBUSCA$, OKXX%)
   
End Sub

Private Sub EJECUTAR_Click()
    SHOWREP07.Label2 = ""
    If TRIM$(Label9) <> "" Then
          SHOWREP07.Label3 = Label9        ' Descripcion de la consulta
       Else
          SHOWREP07.Label3 = "Consulta en Desarrollo"
    End If
    TEXMO$ = ARMA_CONSULTB
    Call DECOD_CONSULTA(TEXMO$)
    'If TRIM$(Combo2.TEXT) <> "" Then
    If TRIM$(Label1.Caption) <> "" Then
         SHOWREP07.Label2 = ""
         'SHOWREP07.Label2 = Combo2.TEXT
         SHOWREP07.Label2 = Label1.Caption
      Else
         SHOWREP07.Label2 = "_TEXT_"
    End If
    SHOWREP07.Show 1
    MODICONSU% = 0
End Sub

Private Sub GUARDACON_Click()
    '
    FGUAREP07.Text1 = TRIM$(Label1.Caption)
    CODMO$ = TRIM$(Label1.Caption)
    FGUAREP07.Text2 = Label9
    FGUAREP07.Text3 = ""
    FGUAREP07.Label11 = ""
    FGUAREP07.Text4 = ""
    FGUAREP07.Label18 = ""
    FGUAREP07.Show 1
    If FGUAREP07.Label1 = "OK" Then
10     CODMO$ = TRIM$(FGUAREP07.Text1)
       DESMO$ = TRIM$(FGUAREP07.Text2)
       TEXMO$ = ARMA_CONSULTB
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If
       
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_QUERY = CODMO$
          !DES_QUERY = DESMO$
          !TIP_QUERY = 0        ' REPORTE TABULADO
          !PAR_QUERY = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       '
30     If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
           ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + CODMO$ + ".SQF"
           Call SAVEFILE(ASALI$, TEXMO$)
       End If
       '
       If FGUAREP07.Text3 <> "" Then
          If FGUAREP07.Label11 <> "" Then
             ASALI$ = FGUAREP07.Label11 + "\" + FGUAREP07.Text3 + ".SQF"
             Call SAVEFILE(ASALI$, TEXMO$)
          End If
       End If
       '
    End If
    '
    Label1.Caption = CODMO$
    '
End Sub




Private Sub Label8_DblClick()
    Unload Me
End Sub

Private Sub List1_Click()
   If MODIFENCA% = 2 Then Exit Sub
   If MODIFENCA% = 1 Then
      If COMALTER%("¿ Descarta Cambios ?\\No, Conservar\Si, Descartar") <> 2 Then
         MODIFENCA% = 2
         On Error Resume Next
         List1.ListIndex = ILIS1&
         On Error GoTo 0
         MODIFENCA% = 1
         Exit Sub
      End If
   End If
   '
   CAMPOSE$ = List1.TEXT
   Text1 = TRIM$(Mid$(CAMPOSE$, 65, 8))
   Check3.Value = 0: If Mid$(CAMPOSE$, 74, 1) = "1" Then Check3.Value = 1
   Check4.Value = 0: If Mid$(CAMPOSE$, 73, 1) = "1" Then Check4.Value = 1    ' OCULTA
   Check1.Value = 0: If Mid$(CAMPOSE$, 76, 1) = "1" Then Check1.Value = 1
   Check2.Value = 0: If Mid$(CAMPOSE$, 75, 1) = "1" Then Check2.Value = 1
   Text2 = TRIM$(Mid$(CAMPOSE$, 81, 64))
   Text5 = TRIM$(Mid$(CAMPOSE$, 145, 16))
   Text6 = TRIM$(Mid$(CAMPOSE$, 161, 8))
   '
   TAUXI$ = Mid$(CAMPOSE$, 169, 640)
   Combo2.TEXT = TRIM$(Mid$(TAUXI$, 1, 64))
   Combo3.TEXT = TRIM$(Mid$(TAUXI$, 65, 64))
   Text7.TEXT = TRIM$(Mid$(TAUXI$, 129, 512))
   '
   ORDEX$ = TRIM$(UCase$(Mid$(CAMPOSE$, 77, 4)))
   Option1.Value = True
   If ORDEX$ = "ASC" Then Option2.Value = True
   If ORDEX$ = "DESC" Then Option3.Value = True
   MODIFENCA% = 0
   ILIS1& = List1.ListIndex
   
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   i& = List2.ListIndex
   If i& >= 0 Then
      CMBCAMPOS.TEXT = TRIM$(Left$(TTXX$, 24))
      CMBOPERADOR.TEXT = TRIM$(Mid$(TTXX$, 26, 8))
        VALORX$ = TRIM$(REPLACAR$(Mid$(TTXX$, 35), "'", ""))
      Combo1.TEXT = VALORX$
      List2.RemoveItem i&
   End If
   Call ARMA_CONSULTX
End Sub

Private Sub LSTMOSTARCMP_Click()
  Call ARMASELFIELD
End Sub

Sub ARMASELFIELD()
    '
    For i& = 0 To LSTMOSTARCMP.ListCount - 1
      ESTADO = LSTMOSTARCMP.Selected(i&)
      CAMPO$ = TRIM$(LSTMOSTARCMP.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 20
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
10         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 10
              End If
           Next J&
      End If
20  Next i&
    '
    For i& = 0 To CAMAUXI.ListCount - 1
      ESTADO = CAMAUXI.Selected(i&)
      CAMPO$ = TRIM$(CAMAUXI.List(i&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 120
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
110         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 110
              End If
           Next J&
      End If
120 Next i&
    '
   List3.Clear
   For i& = 0 To List1.ListCount - 1
      List1.ListIndex = i&
      List3.AddItem List1.TEXT
   Next i&

    Call ARMA_CONSULTX
    '
End Sub
'
Private Sub LSTTABLAS_Click()
    '
    Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Sub ARMALISTAS()

    TXTCONSULTA.TEXT = ""
    CMBCAMPOS.Clear
    CMBOPERADOR.TEXT = ""
    Combo1.Clear
    Combo4.Clear
    Combo5.Clear
    LSTMOSTARCMP.Clear
    CAMAUXI.Clear
    List1.Clear
    '
    TAPRIN$ = LSTTABLAS.TEXT
    PREFPRIN$ = ""
    TASEC$ = LSTTABLASEC.TEXT
    If TASEC$ <> "" Then PREFPRIN$ = TAPRIN$ + "."
    If TAPRIN$ = "" Then Exit Sub
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           CMBCAMPOS.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           LSTMOSTARCMP.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
           Combo4.AddItem AUX.Fields(i& - 1).Name
        End If
    Next i&
    '
    If TASEC$ <> "" Then
       PREFSEC$ = TASEC$ + "."
       Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TASEC$))
       For i& = 1 To AUX.Fields.Count
          If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
            CMBCAMPOS.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            LSTMOSTARCMP.AddItem PREFSEC$ + AUX.Fields(i& - 1).Name
            Combo5.AddItem AUX.Fields(i& - 1).Name
          End If
       Next i&
    End If
    '
    For i& = 1 To 9
        CAMAUXI.AddItem "ValoBada_" & i&
    Next i&
    '
End Sub

Private Sub CMBBASE_CLICK()
    LSTTABLAS.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Private Sub Form_Load()
   If App.PrevInstance = True Then
      Close: End
   End If
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   'Call CargaFuentesAnchoFijo
   '
   If APLINVO$ <> "" Then Height = 3570
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call ABRECONEXION
   KLIK% = 0
   '
   Call ABRETACON     ' crea la tabla de consultas
'   Call VerConsul
   Call CREATABLAS_SQLCTRL

  Call ABRECONEXION
  KLIK% = 0
  CMBOPERADOR.Clear
  
  CMBOPERADOR.AddItem " =  "
  CMBOPERADOR.AddItem " <  "
  CMBOPERADOR.AddItem " <= "
  CMBOPERADOR.AddItem " >  "
  CMBOPERADOR.AddItem " >= "
  CMBOPERADOR.AddItem " <> "
  CMBOPERADOR.AddItem "LIKE"
  '
  CMBJUNTAR.AddItem "INNER"
  CMBJUNTAR.AddItem "LEFT"
  CMBJUNTAR.AddItem "RIGHT"
  '
  UNIONALL$ = ""
  '
  Combo6.AddItem "DEUDOR"
  Combo6.AddItem "PROVED"
  Combo6.AddItem "MASTER"
  Combo6.AddItem "USERS"
  Combo6.AddItem "SUTOT"
 
  Call MUESTRATABLAS

  TEPRUEBA.Font = MSF.Font
  TEPRUEBA.FontSize = MSF.Font.Size
  TEPRUEBA.FontBold = MSF.Font.Bold
  MSF.Rows = 10
  MSF.ColWidth(0) = 4000
  MSF.TextMatrix(0, 0) = "Encabezado + Formato (Canti/F8.1)"
  MSF.TextMatrix(1, 0) = "Formula (4;*;6)"
  MSF.TextMatrix(2, 0) = "CampEditables (S/N)"
  MSF.TextMatrix(3, 0) = "ColEcoCampo (CODIGO;Descripcion;Master)"
  MSF.TextMatrix(4, 0) = "TxtUbicadoGrillaTotal(Texto)"
  MSF.TextMatrix(5, 0) = "Total (S/N)"
  MSF.TextMatrix(6, 0) = "Dejar Libre (S/N)"
  MSF.TextMatrix(7, 0) = "CONDICION 1"
  MSF.TextMatrix(8, 0) = "CONDICION 2"
  MSF.TextMatrix(9, 0) = "ENUMERAR"
  '
  Call CARGAMSF4
  Call ABRECONEXION
  Call CREACAMPO("", "GeneCodi", "id", 4, 0, 1)
  Call CREACAMPO("", "GeneCodi", "AddRegistros", 3, 0)
  Call CREACAMPO("", "GeneCodi", "UtilFormula", 3, 0)
  Call CREACAMPO("", "GeneCodi", "Borra", 3, 0)
  Call CREACAMPO("", "GeneCodi", "Inserta", 3, 0)
  Call CREACAMPO("", "GeneCodi", "NoTilde", 3, 0)
  Call CREACAMPO("", "GeneCodi", "NameRecord", 10, 256)
  Call CREACAMPO("", "GeneCodi", "CamposRecord", 12, 0)
  Call CREACAMPO("", "GeneCodi", "UsaRecord", 3, 0)
  Call CREACAMPO("", "GeneCodi", "AotraGrilla", 3, 0)
  Call CREACAMPO("", "GeneCodi", "ConsultaRec", 10, 1056)
  Call CREACAMPO("", "GeneCodi", "UsaRecord", 3, 0)
  Call CREACAMPO("", "GeneCodi", "CodiGen", 12, 0) 'el codigo generado
  Call CREACAMPO("", "GeneCodi", "TitVentana", 10, 256) 'titulo de la ventana
  Call CREACAMPO("", "GeneCodi", "CanCampos", 3, 1) 'Cantidad de Campos que tiene la grilla
  Call CREACAMPO("", "GeneCodi", "MARBORRA", 3, 1)
  '
  Call CREACAMPO("", "DatGrilla", "id", 4, 0, 1)
  Call CREACAMPO("", "DatGrilla", "FilaCol", 10, 32)
  Call CREACAMPO("", "DatGrilla", "Dato", 10, 512)
  Call CREACAMPO("", "DatGrilla", "MARBORRA", 3, 0)
  '

'  Call CARGA_ENCABEZADO(MSF4, "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10", VENTADISE)
'  MSF4.Rows = 2

  '
  '
  Call CREATABLA_ENCAPRF
  SSTab1.Tab = 0

    

    


End Sub

Sub MUESTRATABLAS()
    LSTTABLAS.Clear: LSTTABLASEC.Clear
    Combo2.Clear: Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
       LSTTABLASEC.AddItem UCase$(AUX!Name)
       Combo2.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Sub ARMA_CONSULTX()
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = TRIM$(Left$(List1.List(i&), 64))
       '
       If UCase$(Left$(CAMPO$, 8)) = "VALOBADA" Then GoTo 19
       '
       ALIASCAMPO$ = CAMPO$
       If Mid$(List1.List(i&), 75, 1) = "1" Then
          CAMPO1$ = CAMPO$
          If InStr(CAMPO1$, ".") > 0 Then
            PPXYZ% = InStr(CAMPO1$, ".")
            CAMPO1$ = Mid$(CAMPO1$, PPXYZ% + 1)
          End If
          ALIASCAMPO$ = "Sx" + CAMPO1$
          CAMPO$ = "SUM(" + CAMPO$ + ") AS " + ALIASCAMPO$
          AGRUPA% = 1
       End If
       If VUELTA& = 0 Then
           SQLX$ = SQLX$ + CAMPO$
          Else
           SQLX$ = SQLX$ + ", " + CAMPO$
       End If
       VUELTA& = VUELTA& + 1
       '
       ORDEX$ = UCase$(TRIM$(Mid$(List1.List(i&), 77, 4)))
       If ORDEX$ <> "" Then
          If VUELTB& > O Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = Mid$(List1.List(i&), 75, 1)
       If AGRUX$ <> "1" Then
          If VUELTC& > O Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If

19  Next i&
    '
    If VUELTA = 0 Then SQLX$ = "SELECT *"
    '
    TAPRIN$ = LSTTABLAS.TEXT
    SQLY$ = " FROM " + TAPRIN$
    '
    TASEC$ = LSTTABLASEC.TEXT
    
    If TASEC$ <> "" Then
       CAMPRIN$ = TRIM$(Combo4.TEXT)
       If CAMPRIN$ <> "" Then
          CAMSEC$ = TRIM$(Combo5.TEXT)
          If CAMSEC$ <> "" Then
             TIPOJOIN$ = TRIM$(CMBJUNTAR.TEXT)
             If TIPOJOIN$ = "" Then TIPOJOIN$ = "INNER"
             SQLY$ = SQLY$ + " " & TIPOJOIN$ & " JOIN"
             SQLY$ = SQLY$ + " " + TASEC$
             
             SQLY$ = SQLY$ + " ON " + TAPRIN$ + "." + CAMPRIN$
             SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
          End If
       End If
    End If
    '
    If List2.ListCount > 0 Then
       JJ& = 0
       For i& = 0 To List2.ListCount - 1
          SQLY$ = SQLY$ '+ vbCrLf
          'INTERVENCION OR
          If JJ& < 1 Then
              SQLY$ = SQLY$ + " WHERE "
          ElseIf UCase$((Left$(TRIM$(List2.List(i&)), 3))) = "OR " Then
              SQLY$ = SQLY$ + " "
          Else
               SQLY$ = SQLY$ + " AND "
          End If
          SQLY$ = SQLY$ + TRIM$(Left$(List2.List(i&), 24)) + " " + TRIM$(Mid$(List2.List(i&), 26, 8)) + " " + TRIM$(Mid$(List2.List(i&), 35))
          JJ& = JJ& + 1
       Next i&
    End If
    '
    TXTCONSULTA.TEXT = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " ORDER BY " + SQLZ$
    '
    '
    If chkUnionAll.Value = 1 Then
        If LSTTABLASEC.ListIndex > 0 Then
            PrimeraTabla$ = TRIM$(LSTTABLAS.List(LSTTABLAS.ListIndex))
            SegundaTabla$ = TRIM$(LSTTABLASEC.List(LSTTABLASEC.ListIndex))
            ConsultaSql$ = TXTCONSULTA.TEXT
            POSI% = InStr(1, ConsultaSql$, PrimeraTabla$)
            If POSI% > 0 Then
                ConsultaSql$ = Mid$(ConsultaSql$, 1, POSI% - 1) & SegundaTabla$ & Mid$(ConsultaSql$, POSI% + Len(PrimeraTabla$))
                TXTCONSULTA.TEXT = SQLX$ + SQLY$ & " UNION ALL " & ConsultaSql$
            End If
            
            UNIONALL$ = SegundaTabla$
        End If
    End If
       
    TXTCANREGI.TEXT = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
End Sub

Function ARMA_CONSULTB()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(FGUAREP07.Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(FGUAREP07.Text2) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   0" + Chr$(124)
    CONSUB$ = CONSUB$ + "SELTABLE = " + LSTTABLAS.TEXT + Chr$(124)
    '
      TASECU$ = TRIM$(LSTTABLASEC.TEXT)
    If TASECU$ <> "" And chkUnionAll.Value = 0 Then CONSUB$ = CONSUB$ + "SELTASEC = " + TASECU$ + Chr$(124)
      CAMPOMATCH1$ = TRIM$(Combo4.TEXT)
    If CAMPOMATCH1$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI1 = " + CAMPOMATCH1$ + Chr$(124)
      CAMPOMATCH2$ = TRIM$(Combo5.TEXT)
    If CAMPOMATCH2$ <> "" Then CONSUB$ = CONSUB$ + "MATCHFI2 = " + CAMPOMATCH2$ + Chr$(124)
    '
    For i& = 0 To List1.ListCount - 1
       CAMPO$ = List1.List(i&)
       CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    Next i&
    '
    For i& = 0 To List2.ListCount - 1
       Condi$ = List2.List(i&)
       CONSUB$ = CONSUB$ + "CONDITIO = " + Condi$ + Chr$(124)
    Next i&
    '
    TIPODEJOIN$ = TRIM$(CMBJUNTAR.TEXT)
    CONSUB$ = CONSUB$ + "TIPOJOIN = " + TIPODEJOIN$ + Chr$(124)
    '
    If chkUnionAll.Value = 1 And LSTTABLASEC.ListIndex > 0 Then
        SegundaTabla$ = LSTTABLASEC.List(LSTTABLASEC.ListIndex)
        CONSUB$ = CONSUB$ + "UNIONALL = " + SegundaTabla$ + Chr$(124)
    End If
    
    ARMA_CONSULTB = CONSUB$
End Function

Private Sub CMDPEGAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Clipboard.GetText + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

'QUEDA PENIDIENTE SI ES UN TEXTO  EL QUE VA A COMPARAR VAYA ENTRE _
COMILLAS SIMPLES (HACERLO CON ISNUMERIC ) _
SI EL OPERADOR ES LIKE DEBERIA LLEVAR EL SIMBOLO DEL PORCENTAJE _
FALTA TAMBIEN EL TEMA DEL EDITOR DE TEXTO _
VALIDAR LOS CAMPOS DE CONDICION _
CAMBIAR LAS PROPIEDADES DE LOS COMBO Y TEXT DE LA CONDICION PAARA QUE _
NO SEAN EDITABLES _
Y EJECUTAR LA CONSULTA

Private Sub LSTTABLASEC_Click()
    '
    If chkUnionAll.Value = 0 And UNIONALL$ = "" Then Call ARMALISTAS
99  Call ARMA_CONSULTX
    '
End Sub

Private Sub MSF3_Click()
'    i& = MSF3.MouseRow
'    J& = MSF3.MouseCol
'    If i& = 0 Or i& > MSF3.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    MSF3.Row = i&
'    MSF3.COL = J&
'    On Error Resume Next
'
'    REG& = i&
'    If REG& < 1 Then
'        Exit Sub
'    End If
'    '
'    TXTORDEN.TEXT = MSF3.TextMatrix(REG&, 1)
'    Me.TXTIDLISTA = MSF3.TextMatrix(ORD%, 2)
'    Me.TXTTIPOID = MSF3.TextMatrix(ORD%, 3)
'    Me.TXTNOMBRELISTA = MSF3.TextMatrix(ORD%, 4)
'    Me.CMBCAMPOS.TEXT = MSF3.TextMatrix(ORD%, 5)
'    Me.CMBPROPVENTANA = MSF3.TextMatrix(ORD%, 6)
'    Me.TXTORDPRES = MSF3.TextMatrix(ORD%, 7)
'    Me.TXTFORMATO.TEXT = MSF3.TextMatrix(ORD%, 8)
'    Me.LBLIDENTITY = MSF3.TextMatrix(REG&, 9)
'    '
'    Call SELECCIONARFILA(MSF3, REG&)
 
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFUPD_Click()
    COL% = MSFUPD.MouseCol
    REG& = MSFUPD.MouseRow
    If REG& > 0 Then Exit Sub
    If COL% < 1 Then Exit Sub
    LSTMOSTARCMP2.ToolTipText = REG& & "," & COL%
    LSTMOSTARCMP2.TEXT = MSFUPD.TextMatrix(REG&, COL%)

    MSFUPD.COL = COL%

    With MSFUPD
        .COL = COL%
        Me.LSTMOSTARCMP2.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8
        Me.LSTMOSTARCMP2.Visible = True
        Me.LSTMOSTARCMP2.ZOrder 0
        FilaActiva = .Row
    End With
End Sub
Private Sub MSFUPD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       Me.TXTFORMUEXCEL = Me.TXTFORMUEXCEL + UBICACIONX1$(MSFUPD.MouseRow, MSFUPD.MouseCol)
    End If
    
End Sub


'Private Sub MSF3_DblClick()
'    VENTANA.Inicializar = 0
'    VENTANA.AgregaRegistro = 1
'    VENTANA.UtilizaFormula = 1
'    VENTANA.HabilitaBorrar = 0
'    VENTANA.HabilitaInsertar = 1
'    VENTANA.NoMostrarBotonTilde = 0
'    VENTANA.Titulo = "EDIT"
'    VENTANA.CampEditables = "1;2;3;4;5;6"
'    VENTANA.Campox = "TIPO_ID/F6;Nombre Campo/a32;Prop.Ventana/a24;Encabezado/A24;Ord.Pres/f7;Formato/a10"
'    VENTANA.Inicializar = 1
'    Call CopiarMsfAOtroMsf(MSF3, VENTABNEW.MSF, 1)
'    '****************************************************************
'
'    VENTABNEW.Show 1
'    If VENTABNEW.APROBADO% > 0 Then
'       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF3, 1)
'    End If
'    Set VENTANA = Nothing
' End Sub
    

'Private Sub MSF_RowColChange()
''    For i& = 1 To MSF.Cols - 1
''       If msf2.Cols < MSF.Cols Then msf2.Cols = MSF.Cols
''       msf2.TextMatrix(0, i&) = MSF.TextMatrix(0, i&)
''       msf2.COL = i&
''       TEPRUEBA.Caption = MSF.TextMatrix(0, i&)
''       msf2.ColWidth(i&) = TEPRUEBA.Width + 300
''
''    Next i&
'
'End Sub

Private Sub Option1_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option2_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option3_Click()
   MODIFENCA% = 1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   A = PreviousTab
   B = SSTab1.Tab
   If LSTTABLAS2.ListCount > 0 Then Exit Sub
   If B = 1 Then
        LSTTABLAS2.Clear
        Me.TXTNAMETABLA = ""
        Dim AUX As ADODB.Recordset
        Set AUX = New ADODB.Recordset
22      SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
        Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
        Do While Not AUX.EOF
           If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
           LSTTABLAS2.AddItem UCase$(AUX!Name)
29      AUX.MoveNext
        Loop

   End If
End Sub

Private Sub Text1_Change()
   MODIFENCA% = 1
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       MSF4.Cols = XVALO(Text11)
       Call POSICIONAR_TXTMOVIL(VENTADISE.TXTMOVIL2, VENTADISE.MSF4, 0, 1)
    End If

End Sub


Private Sub Text15_DblClick()
   If Text15.Left <> MSF5.Left Then
        Text15.Left = MSF5.Left
        Text15.Top = MSF5.Top
        Text15.Height = 4000
        Text15.Width = MSF5.Width
   Else
        Text15.Left = 5020
        Text15.Top = 6750
        Text15.Height = 1305
        Text15.Width = 10200
   
   End If
End Sub


Private Sub Text2_Change()
   MODIFENCA% = 1
End Sub

'==============================================================================================================
Private Sub MSF_Click()
          COLUM = MSF.MouseCol
          REG& = MSF.MouseRow
          If COLUM < 1 Then Exit Sub
          TXTMOVIL.ToolTipText = REG& & "," & COLUM
          TXTMOVIL.TEXT = MSF.TextMatrix(REG&, COLUM)

          MSF.COL = COLUM
          
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            .Visible = True
            .ToolTipText = REG& & "," & COLUM
            .SetFocus
          End With

End Sub
Private Sub MSF4_Click()
          COLUM = MSF4.MouseCol
          REG& = MSF4.MouseRow
          If COLUM < 1 Then Exit Sub
          TXTMOVIL2.ToolTipText = REG& & "," & COLUM
          TXTMOVIL2.TEXT = MSF4.TextMatrix(REG&, COLUM)

          MSF4.COL = COLUM
          
          With TXTMOVIL2
            .ZOrder 0
            .Width = MSF4.CellWidth
            .Left = MSF4.CellLeft + MSF4.Left
            .Top = MSF4.CellTop + MSF4.Top
            .Visible = True
            .ToolTipText = REG& & "," & COLUM
            .SetFocus
          End With

End Sub

Private Sub Text21_Change()

End Sub

Private Sub Text20_DblClick()
   If Text20.Locked = True Then
      Text20.Locked = False
      Text20.BackColor = vbWhite
      Call PINTATEXT(Text20)
   Else
      Text20.BackColor = &HE0E0E0
      Text20.Locked = True
   End If

End Sub
Private Sub Text14_DblClick()
   If Text14.Locked = True Then
      Text14.Locked = False
      Text14.BackColor = vbWhite
      Call PINTATEXT(Text14)
   Else
      Text14.BackColor = &HE0E0E0
      Text14.Locked = True
   End If

End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       MSF.Cols = XVALO(Text8)
       Call POSICIONAR_TXTMOVIL(VENTADISE.TXTMOVIL, VENTADISE.MSF, 0, 1)
    End If
End Sub


Private Sub Timer1_Timer()
   Call PINTATEXT(TXTMOVIL)
   Timer1.Enabled = False
End Sub

 
Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      GRID1.Visible = False
      GRID2.Visible = True
      Exit Sub
    End If
    '
    GRID1.Rows = 2
    For i& = 1 To GRID1.Cols - 1
       GRID1.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    GRID1.Visible = True
    GRID1.ZOrder 0
    GRID1.Cols = GRID2.Cols
    GRID1.Redraw = False
    GRID2.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID2.Rows - 1
      For J& = 1 To GRID2.Cols - 1
        If (InStr(UCase$(GRID2.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID2.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            GRID1.Rows = REG& + 1
            For H& = 1 To GRID1.Cols - 1
              GRID1.TextMatrix(REG&, H&) = GRID2.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    GRID1.Redraw = True
    
    '
End Sub


Private Sub TXTCONSULTA_Change()
   TXTCONSU_ESPEJO = TXTCONSULTA
End Sub

Private Sub TXTFORMUEXCEL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
'    Set tTip = New CLSTOOLTIP
'             ' almacena la fila y columna actual
'
'
'   ' Establece El tipo ( balloon o normal )
'   tTip.Estilo = TTBalloon
'   ' Indica el icono a utilizar ( info, Warning , error etc..)
'   tTip.Icono = TTIconInfo
'   tTip.Delay = 50 ' Tiempo de duración
'  Static f As Long
'  Static c As Long
'  With TXTFORMUEXCEL
'   If f <> .Left Or c <> .Top Then
'             ' almacena la fila y columna actual
'             f = .Left
'             c = .Top
'             ' texto y titulo
''             tTip.Titulo = "AYUDA   "
'             TT1$ = "Texto de la celda : " & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'             TT1$ = TT1$ & vbCrLf & " - -------------"
'              tTip.Texto = TT1$
'             tTip.Crear TXTFORMUEXCEL.hwnd ' crea el Tips
'    Else
'        ' Lo destruye
'        tTip.Destroy
'    End If
'End With
'
''  Static f As Long
''  Static c As Long
'
''  With MSFlexGrid1
''    ' verifica que el mouse está en una columna o fila
''    If .MouseRow > 0 And .MouseCol >= 0 Then
''
''          If f <> .MouseRow Or c <> .MouseCol Then
''             ' almacena la fila y columna actual
''             f = .MouseRow
''             c = .MouseCol
''
''             ' texto y titulo
''             tTip.Titulo = "Ejemplo"
''             tTip.Texto = "Texto de la celda :" & _
''                         .TextMatrix(.MouseRow, .MouseCol)
''
''             tTip.Crear .hwnd ' crea el Tips
''           End If
''    Else
''        ' Lo destruye
''        tTip.Destroy
''    End If
''  End With
  
End Sub

Private Sub TXTID_Change()
    Call BLANFORMU
    IDX& = XVALO(TXTID)
    SQLX$ = "SELECT *  FROM GeneCodi WITH(NOLOCK) WHERE ID = " & IDX&
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
            TXTTITULO = SAFETEXT$(!TitVentana)
            Check5.Value = XVALO(!ADDREGISTROS)
            Check6.Value = XVALO(!UtilFormula)
            Check7.Value = XVALO(!Borra)
            Check8.Value = XVALO(!Inserta)
            Check9.Value = XVALO(!NOTILDE)
            Text10.TEXT = SAFETEXT$(!NameRecord)
            Check10.Value = XVALO(!USARECORD)
            Check11.Value = XVALO(!AotraGrilla)
            TXTCONSU_ESPEJO = SAFETEXT$(!ConsultaRec)
            CamposRecordset$ = SAFETEXT$(!CamposRecord)
            Text9 = SAFETEXT$(!CodiGen)
            Text8 = XVALO(!CanCampos)
            .MoveNext
        Loop
    End With
    Call Text8_KeyPress(13)
    If TRIM$(CamposRecordset$) <> "" Then
       Call TEXTOLOTES$(CamposRecordset$, "|")
       On Error Resume Next
       For i& = 1 To UBound(CADENATEX$)
          List3.AddItem CADENATEX$(i&)
       Next i&
       On Error GoTo 0
    End If
    '
    RS.Close: Set RS = Nothing
    '
    '
    
    SQLX$ = "SELECT * FROM DatGrilla WITH(NOLOCK) WHERE ID = " & IDX&
    Set RS1 = READSET(SQLX$)
    With RS1
       Do While Not .EOF
         FICO$ = TRIM$(SAFETEXT$(!FilaCol))
         If FICO$ <> "" Then
           On Error Resume Next
           Call TEXTOLOTES$(FICO$, "|")
           REG1& = XVALO(CADENATEX$(1))
           FIL1& = XVALO(CADENATEX$(2))
           MSF.TextMatrix(REG1&, FIL1&) = SAFETEXT$(!DATO)
           On Error GoTo 0
         End If
         .MoveNext
        Loop
    End With
    RS1.Close: Set RS1 = Nothing
    TXTMOVIL.Visible = False ' SI NO CUANDO SE POSICIONABA BORRABA LA CELDA
    Call MSF_RowColChange
End Sub

Private Sub TXTID_DblClick()
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID/F8;TITULO VENTANA/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID/F8;TITULO VENTANA/A64", FRMZELECHO)
    FRMZELECHO.Caption = "LISTA DE CODIGOS GENERADOS"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_CODIGENE()
    Call Carga_MSF_Recordset(RS1, "ID/F8;TITULO VENTANA/A64", FRMZELECHO.msf2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    
    IDX& = XVALO(RESP_ZELE_VECT(1))
    
    TXTID = IDX&
 
End Sub
Private Sub TXTIDLISTA_Change()
  IDLIST& = XVALO(TXTIDLISTA)
  MSF4.Rows = 0
  Text11 = IDLIST&
  Call CARGAMSF4
  
  MAXIMOCAMPOS = XVALO(Text11)
  MSF4.Cols = MAXIMOCAMPOS + 1
  Call CARGADATOSUNICOS(IDLIST&)
  SQLX$ = "SELECT * FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA = " & IDLIST&
  SQLX$ = SQLX$ + " ORDER BY ENCAPRF_ID ASC"
  Set RS = READSET(SQLX$)
  With RS
  J& = 0
    Do While Not .EOF
      J& = J& + 1
      MSF4.TextMatrix(0, J&) = SAFETEXT$(!Encabezado)
      MSF4.TextMatrix(1, J&) = SAFETEXT$(!formula)
      MSF4.TextMatrix(2, J&) = SAFETEXT$(!CAMPEDITAB)
      MSF4.TextMatrix(3, J&) = SAFETEXT$(!COLCAMECO)
      MSF4.TextMatrix(4, J&) = SAFETEXT$(!TXTTOTAL)
      MSF4.TextMatrix(5, J&) = SAFETEXT$(!Total)
      MSF4.TextMatrix(6, J&) = SAFETEXT$(!DejarLibre)
      'VER SEGUN LA CANTIDAD DE CAMPOS VALIDARLOS SEGUN EL CAMPO DE CANTIDAD DE CAMPOS QUE SE GUARDA EN EL ENCABEZADO
      'FALTA PRESENTAR LO GUARDADO EN EL ENCABEZADO Y HACER LA RUTINA QUE LEVANTE DESDE LA BASE DE DATOS.
      .MoveNext
    Loop
  End With
  RS.Close
  Set RS = Nothing
  MSF4.Cols = J& + 1
End Sub

Private Sub TXTMOVIL_Change()
'     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
End Sub

Private Sub TXTMOVIL_DblClick()
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COL1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If REG1& = 3 Or REG1& = 7 Or REG1& = 8 Then ' CAMPO ECO O DE CONDICION 1 O CONDICION 2
        ARMACONDI.APROBO% = 0
        'LEVANTO LA CONDICION SI YA ESTABA GUARDADA.
        ARMACONDI.Text7 = MSFESPEJO.TextMatrix(REG1&, COL1&)
        ARMACONDI.Show 1
'        MSFESPEJO.TextMatrix(REG1&, COL1&) = ""
        If ARMACONDI.APROBO% > 0 Then
          'COPIA LA CONDICION EN UNA GRILLA OCULTA
           If REG1& = 3 Then MSFESPEJO.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
           If REG1& = 7 Or REG1& = 8 Then MSF.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
           
        End If
     End If

End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     COLU1& = XVALO(CADENATEX$(2))
     REG1& = XVALO(CADENATEX$(1))
     
     On Error GoTo 0
'     T_Espera ((0.2))
End Sub


'
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
'     If REG& = 1 Then REG& = 0
     MSF.TextMatrix(REG&, COLU1&) = TXTMOVIL.TEXT
     MSF.COL = COLU1&
     TEPRUEBA.Caption = TXTMOVIL.TEXT
     If TEPRUEBA.Width > MSF.ColWidth(COLU1&) Then
       MSF.ColWidth(COLU1&) = TEPRUEBA.Width + 300
     End If
     On Error GoTo 0
     '
     
     If MSF.Cols - 1 > COLU1& Then
       COLU1& = COLU1& + 1
     Else
       COLU1& = 1
     End If
50   MSF.COL = COLU1&
     With TXTMOVIL
        .ZOrder 0
        .Width = MSF.CellWidth
        .Left = MSF.CellLeft + MSF.Left
        .Top = MSF.CellTop + MSF.Top
        .Visible = True
        .ToolTipText = MSF.Row & "," & COLU1&
        
        'ESTO SE PARA QUE CUANDO VAYA MOVIENDOSE TOME EL TEXTO INGRESADO EN EL ENCABEZADO
         Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
         On Error Resume Next
         REG& = XVALO(CADENATEX$(1))
         COLU1& = XVALO(CADENATEX$(2))
         TXTMOVIL.TEXT = MSF.TextMatrix(REG&, COLU1&)
        .SetFocus
     End With
 
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
     Dim TXTMOVILX As TextBox
     Set TXTMOVILX = TXTMOVIL
     If SSTab1.Tab = 1 Then
        Set TXTMOVILX = TXTMOVIL
        Set MSFX = MSF
     ElseIf SSTab1.Tab = 2 Then
        Set TXTMOVILX = TXTMOVIL2
        Set MSFX = MSF4
     End If
    If TXTMOVILX.Visible = False Then Exit Sub
    
    Call TEXTOLOTES$(TXTMOVILX.ToolTipText, ",")
    On Error Resume Next
    REG1& = XVALO(CADENATEX$(1))
    COLU1& = XVALO(CADENATEX$(2))
    
    Select Case KeyCode
        Case 38 '(flecha arriba)
            On Error Resume Next
            MSF.TEXT = TXTMOVILX.TEXT
            If MSFX.Row > 1 Then
              TXTMOVILX.ToolTipText = REG1& - 1 & "," & COLU1&
            End If
            If MSFX.Row > 1 Then MSFX.Row = MSFX.Row - 1
'          MSFX.COL = 11
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
            
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
            .SetFocus
          End With
            On Error GoTo 0
        Case 40 '(flecha abajo)
            On Error Resume Next
            MSFX.TEXT = TXTMOVILX.TEXT
            If MSFX.Row < MSFX.Rows - 1 Then
              TXTMOVILX.ToolTipText = REG1& + 1 & "," & COLU1&
            End If
            If MSFX.Row < MSFX.Rows Then MSFX.Row = MSFX.Row + 1
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
            .SetFocus
          End With
          On Error GoTo 0
          
           Case 39 '(39derecha)
            On Error Resume Next
            'MSFX.SetFocus
            MSFX.TEXT = TXTMOVILX.TEXT
            If MSFX.Cols - 1 > MSFX.COL Then
              TXTMOVILX.ToolTipText = REG1& & "," & COLU1& + 1
'              MSFX.COL = MSFX.COL + 1
            End If

            If MSFX.COL < MSFX.Cols Then MSFX.COL = MSFX.COL + 1

          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
'            .ToolTipText = MSFX.Row
            .SetFocus
          End With
          On Error GoTo 0
           
           Case 37 '(37izquierda)
            On Error Resume Next
            'MSFX.SetFocus
            MSFX.TEXT = TXTMOVILX.TEXT

            If MSFX.COL > 1 Then
              TXTMOVILX.ToolTipText = REG1& & "," & COLU1& - 1
              MSFX.COL = MSFX.COL - 1
            End If
          With TXTMOVILX
            .ZOrder 0
            .Width = MSFX.CellWidth
            .Left = MSFX.CellLeft + MSFX.Left
            .Top = MSFX.CellTop + MSFX.Top
             .TEXT = TRIM$(MSFX.TEXT)
            .Visible = True
'            .ToolTipText = MSFX.Row
            .SetFocus
          End With
          On Error GoTo 0

    End Select

    '
End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
        Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
        On Error Resume Next
        REG& = XVALO(CADENATEX$(1))
        COLU1& = XVALO(CADENATEX$(2))
        MSF.TextMatrix(REG&, COLU1&) = TXTMOVIL
        TXTMOVIL.Visible = False
     End If
End Sub

'Private Sub TXTNOMBRELISTA_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXTCMB(Me.CMBNOMBRECAMPO)
'  End If
'
'End Sub


'Private Sub TXTORDPRES_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXT(Me.TXTFORMATO)
'  End If
'
'End Sub


'Private Sub TXTTIPOID_KeyPress(KeyAscii As Integer)
'  If KeyAscii = 13 Then
'     Call PINTATEXT(TXTNOMBRELISTA)
'  End If
'
'End Sub


Private Sub TXTMOVIL2_Change()
'     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0

End Sub


Private Sub TXTMOVIL2_DblClick()
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COL1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If REG1& = 3 Then ' CAMPO ECO
        ARMACONDI.APROBO% = 0
        'LEVANTO LA CONDICION SI YA ESTABA GUARDADA.
        ARMACONDI.Text7 = MSF4.TextMatrix(REG1&, COL1&)
        ARMACONDI.Show 1
'        MSFESPEJO.TextMatrix(REG1&, COL1&) = ""
        If ARMACONDI.APROBO% > 0 Then
          'COPIA LA CONDICION EN UNA GRILLA OCULTA
           MSF5.TextMatrix(REG1&, COL1&) = ARMACONDI.Text7
        End If
     End If

End Sub


Private Sub TXTMOVIL2_GotFocus()
     Call PINTATEXT(TXTMOVIL2)
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     COLU1& = XVALO(CADENATEX$(2))
     REG1& = XVALO(CADENATEX$(1))
     On Error GoTo 0

End Sub


Private Sub TXTMOVIL2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
'     If REG& = 1 Then REG& = 0
     MSF4.TextMatrix(REG&, COLU1&) = TXTMOVIL2.TEXT
     MSF4.COL = COLU1&
     TEPRUEBA.Caption = TXTMOVIL2.TEXT
     If TEPRUEBA.Width > MSF4.ColWidth(COLU1&) Then
       MSF4.ColWidth(COLU1&) = TEPRUEBA.Width + 300
     End If
     On Error GoTo 0
     '
     
     If MSF4.Cols - 1 > COLU1& Then
       COLU1& = COLU1& + 1
     Else
       COLU1& = 1
     End If
50   MSF4.COL = COLU1&
     With TXTMOVIL2
        .ZOrder 0
        .Width = MSF4.CellWidth
        .Left = MSF4.CellLeft + MSF4.Left
        .Top = MSF4.CellTop + MSF4.Top
        .Visible = True
        .ToolTipText = MSF4.Row & "," & COLU1&
        
        'ESTO SE PARA QUE CUANDO VAYA MOVIENDOSE TOME EL TEXTO INGRESADO EN EL ENCABEZADO
         Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
         On Error Resume Next
         REG& = XVALO(CADENATEX$(1))
         COLU1& = XVALO(CADENATEX$(2))
         TXTMOVIL2.TEXT = MSF4.TextMatrix(REG&, COLU1&)
        .SetFocus
     End With
 
   End If

End Sub


Private Sub TXTMOVIL2_LostFocus()
    If TXTMOVIL2.Visible = True Then
        Call TEXTOLOTES$(TXTMOVIL2.ToolTipText, ",")
        On Error Resume Next
        REG1& = XVALO(CADENATEX$(1))
        COLU1& = XVALO(CADENATEX$(2))
        On Error GoTo 0
        MSF4.TextMatrix(REG1&, COLU1&) = TXTMOVIL2
        TXTMOVIL2.Visible = False
    End If
End Sub




