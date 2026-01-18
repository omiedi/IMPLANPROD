VERSION 5.00
Begin VB.Form FRMSCRAP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORIGEN DE SCRAP"
   ClientHeight    =   4050
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "..."
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
      Left            =   930
      TabIndex        =   11
      Top             =   2040
      Width           =   175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
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
      Left            =   4000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "O.K."
      Top             =   3600
      Width           =   1485
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Balanza"
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
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Lee el Peso de La Balanza"
      Top             =   3000
      Width           =   2190
   End
   Begin VB.CommandButton CMDDE 
      Caption         =   "..."
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
      Left            =   930
      TabIndex        =   1
      Top             =   360
      Width           =   175
   End
   Begin VB.CommandButton CMDCM 
      Caption         =   "..."
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
      Left            =   930
      TabIndex        =   0
      Top             =   1200
      Width           =   175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1200
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg. Tara"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblkgtara 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4800
      TabIndex        =   15
      ToolTipText     =   "Si Tara Editable Doble Clic para Ingresar Peso de Tara"
      Top             =   2520
      Width           =   1185
   End
   Begin VB.Label LBLDESCTARA 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1200
      TabIndex        =   14
      Top             =   2040
      Width           =   4785
   End
   Begin VB.Label LBLTARA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecionar Peso de Tara"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label LblKgImpresos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3240
      TabIndex        =   10
      ToolTipText     =   "Kg. Con Tara Descontada"
      Top             =   3000
      Width           =   2745
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sector"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label LBLSECTOR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   555
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Categoría"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label LBLCATEGORIA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label LBLDESCSECTOR 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   4905
   End
   Begin VB.Label LBLDESCCATEGORIA 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   4905
   End
   Begin VB.Menu mnuT 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuAnulacionInfScrap 
         Caption         =   "Anulación de Informes de Scrap"
      End
      Begin VB.Menu mnuValeMaterialesScrap 
         Caption         =   "Vale de Materiales de Scrap"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuanularvalemateriales 
         Caption         =   "Anulación de Vale de Materiales"
      End
   End
   Begin VB.Menu mnuCl 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuIS 
         Caption         =   "Informes de Scrap"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuVerVM 
         Caption         =   "Visualizar/Re-Imprimir Vale de Materilas"
      End
      Begin VB.Menu mnuimpinfscrap 
         Caption         =   "Visualizar/Re-Imprimir Informes de Scrap"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConf 
      Caption         =   "Configuración"
      Begin VB.Menu mnuFormConsMat 
         Caption         =   "Formulario de Cosumo de Materiales"
      End
   End
End
Attribute VB_Name = "FRMSCRAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCM_Click()
   X = LISTASELECGENERAL("CATESCRP", , , 1)
   Me.LBLCATEGORIA = XVALO(TRIM$(RESP_ZELE_VECT(1)))

'   Me.LBLDESCCATEGORIA = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub CMDDE_Click()
   X = LISTASELECGENERAL("SECTSCRP", , , 1)
   Me.LBLSECTOR = XVALO(TRIM$(RESP_ZELE_VECT(1)))

'   Me.LBLDESCSECTOR = TRIM$(RESP_ZELE_VECT(2))

End Sub

Private Sub Command1_Click()
   X = LISTASELECGENERAL("TABLATARAS", , , 1)
   Me.LBLTARA = XVALO(RESP_ZELE_VECT(1))

'   Me.LBLDESCTARA = TRIM$(RESP_ZELE_VECT(2))
'   Me.lblkgtara = VALOBADA("DATASQL", "CAMPO3", "CONVERT(int, campo1) = '" & XVALO(RESP_ZELE_VECT(1)) & "' AND LISTA = 'TABLATARAS'", "NOMESS")

End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    CONDI$ = " CodSector = " & XVALO(Me.LBLSECTOR)
    CONDI$ = CONDI$ + " AND CodCateg = " & XVALO(Me.LBLCATEGORIA)
    CI1% = XVALO(VALOBADA("CONFSCRAP", "CiArt", CONDI$, "NOMESS"))

    
    If CI1% < 1 Then
      Call COMUNI("Falta Configurar Código Para Esta Combinación de Categoría y Sector")
      GoTo 99
    End If
    Codigo$ = CODEXT$(CI1%)
    If XVALO(LBLTARA) < 1 Then
      Call COMUNI("Tara No Válido")
      GoTo 99
    End If
    
    TARASELEC% = XVALO(Me.LBLTARA)
    
    Dim CTP As New Conteplast
    If InStr(TRIM$(UCase$(LBLDESCTARA)), "EDITABLE") < 1 Then
       PESOX# = CTP.PesoLeidoBalanza(0, TARASELEC%, OK%, 1)
    Else 'SI  ES TARA EDITABLE
       PESOX# = CTP.PesoLeidoBalanza(0, TARASELEC%, OK%, 1, XVALO(lblkgtara))
    End If
    If PESOX# < 0 Then PESOX# = 0
    
    LblKgImpresos = FORMATNUM$(PESOX#, "f12.2")
99  Command11.Enabled = True

End Sub

Private Sub Command2_Click()

    'VALIDAR DATOS
    Command2.Enabled = False
    If XVALO(Me.LBLSECTOR) < 1 Then
       Call COMUNI("Falta Ingresar Sector")
       GoTo 99
    End If
    'VALIDO POR CONSULTA POR SI LO QUE TRAJO EL CONTROL NO COINCIDE CON CAMBIOS EN LA TABLA DE CONFIGURACION
    REFESECTOR$ = TRIM$(VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLSECTOR) & "' AND LISTA = 'SECTSCRP'", "NOMESS"))
    If REFESECTOR$ = "" Then
       Call COMUNI("Sector Ingresado No Válido")
       GoTo 99
    End If
    '
    If XVALO(Me.LBLCATEGORIA) < 1 Then
       Call COMUNI("Falta Ingresar Categoría")
       GoTo 99
    End If
    
    'VALIDO POR CONSULTA POR SI LO QUE TRAJO EL CONTROL NO COINCIDE CON CAMBIOS EN LA TABLA DE CONFIGURACION
    REFECATEGORIA$ = TRIM$(VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLCATEGORIA) & "' AND LISTA = 'CATESCRP'", "NOMESS"))
    If REFECATEGORIA$ = "" Then
       Call COMUNI("Categoría Ingresada No Válida")
       GoTo 99
    End If
    '
    If XVALO(Me.LBLTARA) < 1 Then
       Call COMUNI("Falta Ingresar Tara")
       GoTo 99
    End If
    '
    CANTI = XVALO(LblKgImpresos)
    If CANTI <= 0 Then
       Call COMUNI("Falta Pesar Scrap")
       GoTo 99
    End If
    
    'VALIDO POR CONSULTA POR SI LO QUE TRAJO EL CONTROL NO COINCIDE CON CAMBIOS EN LA TABLA DE CONFIGURACION

     DESCRITARA$ = TRIM$(VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS"))
     If DESCRITARA$ = "" Then
       Call COMUNI("Tara Ingresada No Válida")
       GoTo 99
     End If
     KGTARA = XVALO(VALOBADA("DATASQL", "CAMPO3", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS"))
     If KGTARA <= 0 And InStr(TRIM$(UCase$(LBLDESCTARA)), "EDITABLE") < 1 Then
       OPX% = COMALTER("Atención Tara Seleccionada No Posee Valor de Kg.\\Cancelar\Grabar de Todos Modos")
       If OPX% < 2 Then GoTo 99
     End If
     
    'FUNCION QUE DEVUELVA EL CODIGO SEGUN CATEGORIA Y SECTOR
    CONDI$ = " CodSector = " & XVALO(Me.LBLSECTOR)
    CONDI$ = CONDI$ + " AND CodCateg = " & XVALO(Me.LBLCATEGORIA)
    DEPOIN% = XVALO(VALOBADA("CONFSCRAP", "DEPINGRE", CONDI$, "NOMESS"))
    
    CI1% = XVALO(VALOBADA("CONFSCRAP", "CiArt", CONDI$, "NOMESS"))
    If CI1% < 1 Then
      Call COMUNI("Falta Configurar Código Para La Combinación de Sector y Categoría Seleccionada")
      GoTo 99
    End If
    '
    'TODO ESTE PARRAFO ES POR SI LA COMBINACION DE SECTOR Y CATEGORIA DEVOLVIO MAS DE UN REGISTRO
    'EN ESE CASO SE PRESENTA LOS QUE MACHEAN EN UN SELECHO PARA QUE SE SELECCIONE.
    CANTREG& = CantRegistros&("CONFSCRAP", CONDI$, "NOMESS")
    If CANTREG& > 1 Then
        Campos$ = "CI/A0;Còdigo/A16;Descripciòn/A32"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        CAMPOSTABLA$ = "CIART,CODART,DESCART"
        TABLA$ = " CONFSCRAP "
        'LA CONDI ES LA MISMA QUE VIENE DE ARRIBA
        Dim obj As New RECORXET
        '
        Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
        Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
        
        On Error Resume Next
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        On Error GoTo 0
        FRMZELECHO.Width = 8500
        FRMZELECHO.Caption = "SELECCIONE CODIGO PARA EL SCRAP"
        FRMZELECHO.Show 1
        CI1% = XVALO(RESP_ZELE_VECT(1))
        If CI1% < 1 Then
            Call COMUNI("Operación Cancelada No Se Ha Seleccionado el Código\(La Combinación de Categoría y Sector, Arrojó más de un Registro)\Vuelva a Intentar y Seleccione un Código Realizando Doble Clic sobre el mismo")
            GoTo 99
        End If
    End If
    
    If DEPOIN% < 1 Then
      Call COMUNI("Falta Configurar Depósito Para La Combinación de Sector y Categoría Seleccionada")
      GoTo 99
    End If
    
   'GUARDAR EL CONTROL
    Call GRACONTROL(IDENTER$, "SECTSCRP", LBLSECTOR)
    Call GRACONTROL("*", "SECTSCRP", LBLSECTOR)
    Call GRACONTROL(IDENTER$, "CATSCRP", Me.LBLCATEGORIA)
    Call GRACONTROL("*", "CATSCRP", LBLCATEGORIA)
    Call GRACONTROL(IDENTER$, "TARASCRP", Me.LBLTARA)
    Call GRACONTROL("*", "TARASCRP", LBLTARA)
    '
    FE% = HOY(HO$)
    FEC$ = FETEXCOR1$(FE%)
    Codigo$ = CODEXT$(CI1%)
    DESCRICOD$ = DESCRIT0$(CI1%)

    SQLX$ = "SELECT * FROM DETASCRAP"
    SQLX$ = SQLX$ + " WHERE DETASCRAP_ID < 1"
    FE% = HOY(HO$)
    FEC$ = FETEXCOR1$(FE%)
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With RS
        .AddNew
        !CodSector = XVALO(Me.LBLSECTOR)
        !DescSector = Left$(REFESECTOR$, 32)
        !CodCateg = XVALO(Me.LBLCATEGORIA)
        !DescCateg = REFECATEGORIA$
        !CodArt = Left$(Codigo$, 24)
        !DescArt = Left$(DESCRICOD$, 64)
        !DepIngre = DEPOIN%
        !CiArt = CI1%
        !FECHA = FEC$
        !CODIEMPR = CodiEmp%
        !MARBORRA = 0
        !FEREMOV = Now
        !KILOS = XVALO(LblKgImpresos)
        .Update
    End With
        
    Set RS = Nothing
    '
    'OBTENGO EL ULTIMO IDENTITY GENERADO
    SQLX$ = "SELECT @@IDENTITY AS NUEVOID FROM DETASCRAP WITH(NOLOCK)"
    Set RS = READSET(SQLX$)
    NuevoIDGenerado& = XVALO(RS!NUEVOID)

   'GENERAR MOVIMIENTO DE STOCKS
   '****************************
   CMOV$ = "SC"
   NRO& = ProNroComprobante(CMOV$)
   HOII% = HOY(HOS$)
   FF% = HOII%
   DOPRI$ = "SC." + TRIM$(str$(NRO&))
   DOSEC$ = "SC." & NuevoIDGenerado&
   '
   NUORIT% = 0
   VUNI# = 0
   MONEII% = 0
   NC% = 0
   REFE$ = "Inf.de Scrap"
   Call MOVISTOK(FF%, CI1%, "", CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPOIN%, CANTI)
   '
   'FALTA LISTA DE MOVIMIENTOS DE DETALLE
   'ELIMINAR ALGUN MOVIMIENTO INFORMADO
   'REMITIR LLUGO O BIEN INFORAMR CONSUMO
   'NO HAGO MOVIMIENTO DE SALIDA SOLO DE INGRESO AL DEPOSITO DE SCRAP POR QUE NO ESTA IDENTIFICADO NI POR
   'LOTE Y APARTE ES EL SCRAP ES POR KG Y EL RESTO ES POR METRO VEREMOS COMO MANEJAMOS EL TEMA DE LA
   'RELACION DEL CONSUMO FINAL
   
   Call PRINTETIQSCRAP(CI1%, Me.LBLSECTOR, Me.LBLCATEGORIA, REFECATEGORIA$, CANTI, 1, NuevoIDGenerado&)


    
    'GENERAR MOVIMIENTO DEL CODIGO EN DEPOSITO
99   Screen.MousePointer = 1
     Command2.Enabled = True
     LblKgImpresos = 0
End Sub

 Sub PRINTETIQSCRAP(CI1%, CODISECTOR, CODICATEGORIA, REFECATEGORIA$, pesoReal, CANTIDADCOPIAS, NuevoIDGenerado&, Optional FE%)

' PRINTETIQSCRAP(NORFA$, LoteNorfa$, pesoReal, metrosReal, CANCOPIAS&)
        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""

        If CI1% < 1 Then
           Call COMUNI("IMPOSIBLE CONTINUAR NO SE PUDO OBTENER EL CÓDIGO/PERTENECIENTE A LA ORDEN DE FABRICACIÓN ACTIVA")
           Exit Sub
        End If
        PIMPRE$ = TRIM$(Control(IDENTER$, "PORTETIQ")) 'configurado en opvalemat (Pto.O/Etiqueta OF:)
        If PIMPRE$ = "" Then PIMPRE$ = Control$("*", "PORTETIQ")
        Call SETPRINPORT(PIMPRE$)
        
        Codigo$ = CODEXT$(CI1%)
        Descripcion$ = DESCRIT0$(CI1%)
        PESOENKG$ = TRIM$(FORMATNUM$(pesoReal, "F10.0"))
        If FE% < 1 Then 'SI SE PASA FECHA NO LA BUSCA
          CONDI$ = "MOVISTO_ID = " & NuevoIDGenerado& 'ACA LA BUSCA EN EL MOVISTO
          FECHAGENERACION% = CVINT(VALOBADA("MOVISTO", "FECHMOV", CONDI$, "NOMESS"))
          
        Else
          FECHAGENERACION% = FE% 'CASO QUE SE PASA FECHA LA LEE Y ASIGNA
          FECHA$ = TRIM$(FECHATEX$(FECHAGENERACION%))
        End If
        
        FORIPRE$ = TRIM$(Control("CONTEPLA", "ETISCRAP"))
        If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Sub

        IMAG$ = "IMASCRAP_QR.JPG"
        Identificador$ = "QR_ETSCRAP"
        ANCHOX% = 0
        ALTOX% = 0

        On Error Resume Next
        Kill LUCOM$ & IMAG$
        'MANTENGO EL POSICIONAMIENTO DE LA ETIQUETA DE LAS BOBINAS
        TEXTOQR$ = Codigo$ & "|" & " " & "|" & " " & "|" & PESOENKG$ & "|" & "0"
        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
        QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
        Call T_Espera(3)
        Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
        Call GRABAR_QR(LUCOM$ & IMAG$, Identificador$)

'        FILTX$ = LoteNorfa$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & RazonSocialProveedor$ & ";" & METROSREALES$ & ";" & PESOENKG$
        FILTX$ = str$(CODISECTOR) & ";" & str$(CODICATEGORIA) & ";" & REFECATEGORIA$ & ";" & Codigo & ";" & PESOENKG$ & ";" & FECHA$
        Call STDFORM(FORIPRE$, FILTX$, "", "", 0, CANCOPIAS&)

        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""
        'EL FRM ES: ETSC-CMB
'AHORA FALTA PROBAR SI EL QR TIENE LA INFO Y SI SE PUEDE CARGAR EN EL REMITO
 End Sub
Private Sub Form_Load()
    LBLSECTOR = XVALO(Control("*", "SECTSCRP"))
    If Control(IDENTER$, "SECTSCRP") <> "" Then LBLSECTOR = XVALO(Control(IDENTER$, "SECTSCRP"))
    
    LBLCATEGORIA = XVALO(Control("*", "CATSCRP"))
    If Control(IDENTER$, "CATSCRP") <> "" Then LBLCATEGORIA = XVALO(Control(IDENTER$, "CATSCRP"))
    
    Me.LBLTARA = XVALO(Control("*", "TARASCRP"))
    If Control(IDENTER$, "TARASCRP") <> "" Then LBLTARA = XVALO(Control(IDENTER$, "TARASCRP"))

End Sub

Private Sub LBLCATEGORIA_Change()
   Me.LBLDESCCATEGORIA = VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLCATEGORIA) & "' AND LISTA = 'CATESCRP'", "NOMESS")

End Sub

Private Sub lblkgtara_DblClick()
    If InStr(TRIM$(UCase$(LBLDESCTARA)), "EDITABLE") > 0 Then
       ValorIngresado$ = TRIM$(InputBox$("Informar", "Ingrese Tara (kg.)", lblkgtara.Caption))
       lblkgtara.Caption = FORMATNUM$(XVALO(ValorIngresado$), "F10.2")
    End If
    
    
End Sub

Private Sub LBLSECTOR_Change()
   Me.LBLDESCSECTOR = VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLSECTOR) & "' AND LISTA = 'SECTSCRP'", "NOMESS")

End Sub

Private Sub LBLTARA_Change()
   lblkgtara = ""
   Me.LBLDESCTARA = VALOBADA("DATASQL", "CAMPO2", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS")
   Me.lblkgtara = VALOBADA("DATASQL", "CAMPO3", "CONVERT(int, campo1) = '" & XVALO(LBLTARA) & "' AND LISTA = 'TABLATARAS'", "NOMESS")

End Sub

Private Sub mnuAd_Click()

End Sub

Private Sub mnuAnulacionInfScrap_Click()
  
    Call DESHASFECHA(Des%, Has%, PERIO$)
    
    If PERIO$ <> "" Then
      If Des% > Has% Then
        Call COMUNI("Rango de Fechas Seleccionado No Válido")
        Exit Sub
      End If
    Else
      Exit Sub
    End If
    Dim CTP As New Conteplast
    idinfScrap& = CTP.InformesDeScrap&(Des%, Has%, PERIO$)
    If idinfScrap& > 0 Then
       OPX% = COMALTER%("Realmente Desea Anular El Informe de Scrap Seleccionado (ID: " & idinfScrap& & ")\\Cancelar\Anular")
       If OPX% < 2 Then Exit Sub
       Call CTP.AnularInformeScrap(idinfScrap&)
    End If

    Set CTP = Nothing
End Sub



Private Sub mnuanularvalemateriales_Click()
   Call ANUVALEMAT
End Sub

Private Sub mnuFormConsMat_Click()
    ConfigDepo.Show 1
End Sub


Private Sub mnuimpinfscrap_Click()
    Call DESHASFECHA(Des%, Has%, PERIO$)
    
    If PERIO$ <> "" Then
      If Des% > Has% Then
        Call COMUNI("Rango de Fechas Seleccionado No Válido")
        Exit Sub
      End If
    Else
      Exit Sub
    End If
    Dim CTP As New Conteplast
    idinfScrap& = CTP.InformesDeScrap&(Des%, Has%, PERIO$)
    
    If idinfScrap& > 0 Then
     CI1% = XVALO(RESP_ZELE_VECT(9))
     CODISECTOR = XVALO(RESP_ZELE_VECT(2))
     CODICATEGORIA = XVALO(RESP_ZELE_VECT(4))
     REFECATEGORIA$ = RESP_ZELE_VECT(5)
     CANTI = XVALO(RESP_ZELE_VECT(11))
     FECHAX$ = RESP_ZELE_VECT(1)
     FEGEN% = FECHANUM(FECHAX$)
     Call PRINTETIQSCRAP(CI1%, CODISECTOR, CODICATEGORIA, REFECATEGORIA$, CANTI, 1, idinfScrap&, FEGEN%)
    End If

'    Set CTP = Nothing
'    Campos$ = "Fecha/D8;" '1
'    Campos$ = Campos$ + "C.Sect./F6;" '2
'    Campos$ = Campos$ + "Descripción/A32;"      '3
'    Campos$ = Campos$ + "C.Categ./F7;"      '4
'    Campos$ = Campos$ + "Descripción/A32;"   '5
'    Campos$ = Campos$ + "Cód.Art./A10;"       '6
'    Campos$ = Campos$ + "Descripción/A32;"        '7
'    Campos$ = Campos$ + "N°Dep./F6;"     '8
'    Campos$ = Campos$ + "CI/A0;" '9
'    Campos$ = Campos$ + "ID/A0;" '10
'    Campos$ = Campos$ + "KG./F8.2" '11

End Sub

Private Sub mnuIS_Click()
    Call DESHASFECHA(Des%, Has%, PERIO$)
    
    If PERIO$ <> "" Then
      If Des% > Has% Then
        Call COMUNI("Rango de Fechas Seleccionado No Válido")
        Exit Sub
      End If
    Else
      Exit Sub
    End If
    Dim CTP As New Conteplast
    idinfScrap& = CTP.InformesDeScrapReporte(Des%, Has%, PERIO$)


    Set CTP = Nothing
End Sub


Private Sub mnuValeMaterialesScrap_Click()

        PREBALANZA.Show 1
        PESOLEIDO# = XVALO(PREBALANZA.LblKgImpresos)
        If PESOLEIDO# <= 0 Then
           Call COMUNI("Peso Ingresado No Válido")
           Exit Sub
        End If
        VENTANA.Campox = "Códig/A16;Descripción/A32;Kgs/F10.2;Referencia/A48"
        VENTANA.Inicializar = 0
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 0
        VENTANA.AgregaRegistro = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Informe de Vale de Materiales de Scrap"
        VENTANA.CampEditables = "1;4"
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER;FAMILIA  = 'SC'"
        VENTANA.Inicializar = 1
        VENTABNEW.MSF.TextMatrix(1, 3) = FORMATNUM$(PESOLEIDO#, "F10.2")

        '*****************************************************'
'20      Call SELECHO("TADEPOSI/Depósito de Consumo")
'        DEPO$ = VALACT1$("TADEPOSI")
'        If DEPO$ = "" Then
'           OPX% = COMALTER%("Imposible Continuar, Depósito Mandatario\\Cancelar\Seleccionar Depósito")
'           If OPX% < 2 Then
'              Exit Sub
'           Else
'              GoTo 20
'           End If
'        End If
        DepOrig% = XVALO(Control("CONTEPLA", "DEPOSCRP"))
        If DepOrig% < 1 Then
           Call COMUNI("Falta Configurar Depósito de Scrap")
           Exit Sub
        End If
        
        DEPOSI$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))

30      VENTABNEW.Show 1
        Dim MSF As Object
        Set MSF = VENTABNEW.MSF
        '
        If VENTABNEW.APROBADO% < 1 Then Exit Sub

        If VENTABNEW.APROBADO% > 0 Then
          With VENTABNEW.MSF
           For i& = 1 To .Rows - 1
             COD$ = .TextMatrix(i&, 1)
             If CODINT(COD$) < 1 Then
               Call COMUNI("Código No Válido en fila: " & i&)
               GoTo 30
             End If
             XGS# = XVALO(.TextMatrix(i&, 3))
             If XGS# <= 0 Then
                  Call COMUNI("Faltan Informar Kgs. En Fila: " & i&)
                  GoTo 30
             End If
             OBS$ = XVALO(.TextMatrix(i&, 4))
             If TRIM$(OBS$) = "" Then
                  Call COMUNI("Falta Agregar Referencia en Fila:: " & i&)
                  GoTo 30
             End If
            Next i&
           End With
        End If
        '
        CMO$ = TRIM$(Control$("", "CODCONSU"))
        If CMO$ = "" Then CMO$ = "VM"
        NVM$ = TRIM$(ProNroComprobante(CMO$))
        NRO& = XVALO(NVM$)
        REMI$ = CMO$ + "." & NVM$

        'GENERA MOVIMIENTOS DE STOCK
        HOII% = HOY(HOS$)
        With VENTABNEW.MSF
           For i& = 1 To .Rows - 1
             COD$ = .TextMatrix(i&, 1)
             CI1% = CODINT(COD$)
             XGS# = XVALO(.TextMatrix(i&, 3))
             REFE$ = TRIM$(.TextMatrix(i&, 4))
             GoSub GENERARMOVIMIENTOMOVISTO
           Next i&
        End With
       '
       FORIPRE$ = TRIM$(Control$("", "FORCONSU"))
       If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Consumo de Materiales:\  \'" + RMCC$ + "'.")
           Exit Sub
         End If
       End If

       If TRIM$(FORIPRE$) <> "" Then
         '
         'VALIDACION DE MODO DE IMPRESION
         MODOIMPRE07.Check7.Caption = "Imprime Vale de Consumo"
         MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
         MODOIMPRE07.Check8.Value = 0
         Screen.MousePointer = 1
         MODOIMPRE07.Show 1
         
         FEX = HOII%
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = REMI$
         TIPODOC$ = "Vale de Consumo de Materiales"
         '
         CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
         CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
         
         Call CIERRARCH("!OBSERVOF")
         Call BLANARCH("!OBSERVOF")
         VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Observaciones de Vale de Materiales")

         CAPARDOC% = 2
         '
         CODTABU1$(1) = "COD-MPRI"
         CODTABU1$(2) = "DES-MPRI"
         CODTABU1$(3) = "CAN-MPRI"
         CODTABU1$(4) = "UNI-MPRI"

         CACOLTAB1% = 4
         '
         DESTIDOC% = 0
         CAITAB1% = 0
         CANCELPRINT% = 0
         If MODOIMPRE07.Label1.Caption = "OK" Then
          If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
         End If
         Screen.MousePointer = 11
         '
         With VENTABNEW.MSF
           For i& = 1 To .Rows - 1
             COD$ = .TextMatrix(i&, 1)
             DESX$ = .TextMatrix(i&, 2)
             CI1% = CODINT(COD$)
             XGS# = XVALO(.TextMatrix(i&, 3))
             REFE$ = TRIM$(.TextMatrix(i&, 4))
             UNID$ = TRIM$(Unimed$(CI1%))
             CANS$ = CSTRING$(MKD$(XGS#), 3, 9, 1, 2)
             
             CAITAB1% = CAITAB1% + 1
             TETABU1$(1, CAITAB1%) = COD$
             TETABU1$(2, CAITAB1%) = DESX$
             TETABU1$(3, CAITAB1%) = CANS$
             TETABU1$(4, CAITAB1%) = UNID$
           Next i&
         End With
         '
         Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
       End If
       '
       Set VENTANA = Nothing

       Exit Sub
       
GENERARMOVIMIENTOMOVISTO:
        LOTE$ = ""
        
        If CI1% > 0 Then
            CAN# = XGS#
            If CAN# > 0.005 Then
              DOSEC$ = "V.M.Scrap"
              CANTI = (-1) * CAN#
              NUORIT% = NUORIT% + 1
              VUNI# = 0
              Call MOVISTOK(HOII%, CI1%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)
            End If
        End If
Return

End Sub


Private Sub mnuVerVM_Click()

        Call CARGALISEL("PRIVALMA", "PDOC_TBL", "NUMEBUSQUE/A10;REFERENCIA/A64", "TIPODOCU='Vale de Consumo de Materiales'", "DISTINCT")
        '
5       Call SELECHO("PRIVALMA/Vale de Consumo de Materiales")
        RESUL& = XVALO(VALACT1$("PRIVALMA"))
        If RESUL& > 0 Then
           CMO$ = TRIM$(Control$("", "CODCONSU"))
           If CMO$ = "" Then CMO$ = "VM"
           Call MUESTRADOC(0, "", "Vale de Consumo de Materiales", str$(RESUL&), CMO$ & "." & CStr(RESUL&))
           GoTo 5
        End If

End Sub
