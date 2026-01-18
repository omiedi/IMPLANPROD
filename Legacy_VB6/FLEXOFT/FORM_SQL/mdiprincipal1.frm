VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.MDIForm mdiprincipal1 
   BackColor       =   &H00404040&
   Caption         =   "CENTRO MÉDICO ESCUELA"
   ClientHeight    =   7635
   ClientLeft      =   -105
   ClientTop       =   660
   ClientWidth     =   8850
   Icon            =   "mdiprincipal1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "mdiprincipal1.frx":058A
      Top             =   360
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnualumnos 
         Caption         =   "Alumnos"
      End
      Begin VB.Menu AbmProfesor 
         Caption         =   "Profesores"
      End
      Begin VB.Menu mnuCursos 
         Caption         =   "Cursos"
      End
      Begin VB.Menu mnuTablasAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuActivarCurso 
         Caption         =   "Alta de Nuevos Horarios"
      End
      Begin VB.Menu mnuAltaAlumnoACurso 
         Caption         =   "Alta de Alumno a un Curso"
      End
   End
   Begin VB.Menu mnuusuarios 
      Caption         =   "Usuarios"
      Begin VB.Menu mnuuser 
         Caption         =   "A.B.M.\Derechos de Usuarios "
      End
      Begin VB.Menu mnuderuser 
         Caption         =   "Derechos de Usuarios "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClave 
         Caption         =   "Cambio de Clave"
      End
   End
   Begin VB.Menu mnuconfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnunuevo 
         Caption         =   "Cuentas Contables"
      End
      Begin VB.Menu mnudsg 
         Caption         =   "Diseñador de Formularios"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
      End
   End
   Begin VB.Menu mntitulo 
      Caption         =   "Cuotas"
      Begin VB.Menu mnuCuotas 
         Caption         =   "Generar Cuotas"
      End
   End
   Begin VB.Menu mnurep 
      Caption         =   "Reportes"
      Begin VB.Menu mnuPresentes 
         Caption         =   "Generar Archivo de Presentes"
      End
      Begin VB.Menu mnugenerep 
         Caption         =   "Generador de Reporte"
      End
      Begin VB.Menu mnuAlumxDia 
         Caption         =   "Alumnos Filtrado por Día"
      End
      Begin VB.Menu mnupendcob 
         Caption         =   "Pendientes de Cobranza"
      End
   End
   Begin VB.Menu mnutsoreria 
      Caption         =   "Tesorería"
      Begin VB.Menu mnuaccrec 
         Caption         =   "Acciones de Recibos"
      End
      Begin VB.Menu mnucaja 
         Caption         =   "Caja"
      End
   End
   Begin VB.Menu historial 
      Caption         =   "Historial"
      Begin VB.Menu mnuhist 
         Caption         =   "Historial de Alumnos"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "mdiprincipal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Sub CREAR_CAMPOS_EXTRAS_MASTER()
'   Call CREACAMPO("ARTICULOS", "Master", "Codigo", 10, 24, 1)
'   Call CREACAMPO("ARTICULOS", "Master", "Descripcion", 10, 64)
'   Call CREACAMPO("ARTICULOS", "Master", "DescriComp", 10, 64)
'   Call CREACAMPO("ARTICULOS", "Master", "Observa", 12, 0)
'   Call CREACAMPO("ARTICULOS", "Master", "Fecosto", 8, 0)
   
   Call CREACAMPO("ARTICULOS", "Master", "Codint", 4, 0, 2)
   Call CREACAMPO("ARTICULOS", "Master", "Grucoven", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "COD_AREA", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "DESCR_AREA", 10, 128)
   Call CREACAMPO("ARTICULOS", "Master", "Monecos", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Fecosto", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeIniCurs", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeFinCurs", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "COD_DIA1", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "COD_DIA2", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "COD_DIA3", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DESC_DIA1", 10, 48)
   Call CREACAMPO("ARTICULOS", "Master", "DESC_DIA2", 10, 48)
   Call CREACAMPO("ARTICULOS", "Master", "DESC_DIA3", 10, 48)
   Call CREACAMPO("ARTICULOS", "Master", "ID_PROFESOR", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "ID_AULA", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Cosuni", 7, 8)
   Call CREACAMPO("ARTICULOS", "Master", "CantMeses", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "ValMatricula", 7, 8)
   Call CREACAMPO("ARTICULOS", "Master", "Capacidad", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CantHorCso", 7, 8)
   Call CREACAMPO("ARTICULOS", "Master", "DESCRICURSO", 10, 256)
   Call CREACAMPO("ARTICULOS", "Master", "REFERENCIA", 10, 256)
   Call CREACAMPO("ARTICULOS", "Master", "H_DESDE", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "H_HASTA", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "ID_CURSO", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "FEALTA", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "H_DESDE_FN", 4, 0)
   Call CREACAMPO("ARTICULOS", "Master", "COD_RESPONSABLE", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DESC_RESPONSABLE", 10, 48)
  
   If ValidaExisteCampo%("Master_ID", "Master") < 1 Then
        SQLX$ = "ALTER TABLE Master ADD Master_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
   End If

End Sub

Private Sub EXPE_Click()
   EXPENSAS.Show
End Sub


Sub GENERAR_ESTRUCTURAS_SQL_INSTITUTO()
'|INICIO AREAS
'
'FLEXCONN.Execute ("DELETE FROM  ENCAPRFUN")
'FLEXCONN.Execute ("DELETE FROM  ENCAPRF")

    CONDI$ = " LISTA =   'AREAS'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 15862
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'AREAS'"
      SQLX$ = SQLX$ + ",'LISTA DE AREAS'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,3"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,'22/01/18'"
      SQLX$ = SQLX$ + " ,'AREAS'"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) '"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CÓDIGO/F8'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'AREA/A84'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
'|INICIO DIAS_SEMANA
'
CONDI$ = " LISTA =   'DIAS_SEMANA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15863
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'DIAS_SEMANA'"
  SQLX$ = SQLX$ + ",'DIAS DE LA SEMANA'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'23/01/18'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT  CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NUMDIA/F8'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Día/A20'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN DIAS_SEMANA
'|INICIO AULAS
'
CONDI$ = " LISTA =   'AULAS'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15864
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'AULAS'"
  SQLX$ = SQLX$ + ",'TABLA DE AULAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'24/01/18'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'C.A./F4'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REF. AULA/A40'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN AULAS

'|INICIO MOTIVOS
'
CONDI$ = " LISTA =   'MOTIVOS'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15919
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'MOTIVOS'"
  SQLX$ = SQLX$ + ",'TABLA DE MOTIVOS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'15/03/18'"
  SQLX$ = SQLX$ + " ,'TABLA MOTIVOS'"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/F7'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'MOTIVO/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN MOTIVOS
'|INICIO RESPONSABLES
'
CONDI$ = " LISTA =   'RESPONSABLES'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15954
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'RESPONSABLES'"
  SQLX$ = SQLX$ + ",'TABLA DE RESPONSABLES'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'18/04/18'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'select campo1, campo2 from datasql with(nolock)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Codigo/F5'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Responsable/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 's'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN RESPONSABLES


End Sub

Private Sub AbmProfesor_Click()
   If DERACCE%("MAESPROF", "ACCESO A MAESTRO DE PROFESORES") < 2 Then Exit Sub

   FRMPROFESO1.Show 1
End Sub

Private Sub MDIForm_Load()
'

    
    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")
    '

    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
   '
   '
'    If HOYSERVER% > FECHANUM("31/04/18") Then
'       Call FINAL("")
'    End If
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    Me.Arrange 0
    Me.Caption = Me.Caption + "--- " + EPAC$

    mdiprincipal1.WindowState = 2
    '

    Call ABRECONEXION
    Call ABREMASTER
    Call ABRECABAN

    Call CREATABLA_ENCAPRF
     'TABLA ANOTA_ALUMNO
    Call CREACAMPO("", "ANOTA_ALUMNO", "LEGAJO", 4, 0, 1)
    Call CREACAMPO("", "ANOTA_ALUMNO", "ANOTACION", 12, 0)
    Call CREACAMPO("", "ANOTA_PROFESOR", "LEGAJO", 4, 0, 1)
    Call CREACAMPO("", "ANOTA_PROFESOR", "ANOTACION", 12, 0)
    If ValidaExisteCampo%("ANOTA_PROFESOR_ID", "ANOTA_PROFESOR") < 1 Then
        SQLX$ = "ALTER TABLE ANOTA_PROFESOR ADD ANOTA_PROFESOR_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CURSOS", "COD_CURSO", 10, 24, 1)
    Call CREACAMPO("", "CURSOS", "DESCRIPCION", 10, 512)
    Call CREACAMPO("", "CURSOS", "COD_AREA", 4, 0, 1)
    Call CREACAMPO("", "CURSOS", "AREA", 10, 128)
    Call CREACAMPO("", "CURSOS", "USUARIO", 4, 0, 1)
    Call CREACAMPO("", "CURSOS", "ValCurso", 7, 8)
    Call CREACAMPO("", "CURSOS", "ValMatricula", 7, 8)
    Call CREACAMPO("", "CURSOS", "ValDerExamen", 7, 8)
    '
    If ValidaExisteCampo%("CURSOS_ID", "CURSOS") < 1 Then
        SQLX$ = "ALTER TABLE CURSOS ADD CURSOS_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "TAPROVINCIA", "CODINUM", 3, 0, 1)
    Call CREACAMPO("", "TAPROVINCIA", "CODITEX", 10, 8)
    Call CREACAMPO("", "TAPROVINCIA", "DESCRI_PROV", 10, 36)
    
    If ValidaExisteCampo%("TAPROVINCIA_ID", "TAPROVINCIA") < 1 Then
      SQLX$ = "ALTER TABLE TAPROVINCIA ADD TAPROVINCIA_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
      Call CREACAMPO("BADACLIEN", "TAPROVINCIA", "TAPROVINCIA_ID", 4, 0, 1)
    End If

    '
    'TABLA DE ASOCIASION DE CURSOS A ALUMNOS
    Call CREACAMPO("", "ALUMNO_CURSO", "NUME_CLI", 4, 0, 1)
    Call CREACAMPO("", "ALUMNO_CURSO", "CURSOS_ID", 4, 0, 1)
    Call CREACAMPO("", "ALUMNO_CURSO", "FECHA", 8, 0)
    Call CREACAMPO("", "ALUMNO_CURSO", "Descripcion", 10, 512)
    Call CREACAMPO("", "ALUMNO_CURSO", "STATUS", 3, 0)
    Call CREACAMPO("", "ALUMNO_CURSO", "USUARIO", 3, 0, 1)
    If ValidaExisteCampo%("ALUMNO_CURSO_ID", "ALUMNO_CURSO") < 1 Then
        SQLX$ = "ALTER TABLE ALUMNO_CURSO ADD ALUMNO_CURSO_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If

    '
    Call CREACAMPO("", "PRESENTES", "FECHLIS", 8, 0, 1)
    Call CREACAMPO("", "PRESENTES", "DIAATEX", 10, 12)
    Call CREACAMPO("", "PRESENTES", "MESTEX", 10, 24)
    Call CREACAMPO("", "PRESENTES", "ANO", 3, 0)
    Call CREACAMPO("", "PRESENTES", "NUME_CLI", 3, 0, 1)
    Call CREACAMPO("", "PRESENTES", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "PRESENTES", "PRESENTE", 3, 0)
    Call CREACAMPO("", "PRESENTES", "CODIEMPR", 3, 2, 1)
    Call CREACAMPO("", "PRESENTES", "PRESTE", 3, 2)
    Call CREACAMPO("", "PRESENTES", "USERNUM", 3, 0)
    Call CREACAMPO("", "PRESENTES", "DESCRICURSO", 10, 256)
    Call CREACAMPO("", "PRESENTES", "OBSERVACIONES", 10, 256)
    
    If ValidaExisteCampo%("PRESENTES_ID", "PRESENTES") < 1 Then
        SQLX$ = "ALTER TABLE PRESENTES ADD PRESENTES_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    '
    If ValidaExisteCampo%("CAJABANC_ID", "CAJABANC") < 1 Then
        SQLX$ = "ALTER TABLE CAJABANC ADD CAJABANC_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    '
    If ValidaExisteCampo%("DEUDOR12_ID", "DEUDOR12") < 1 Then
        SQLX$ = "ALTER TABLE DEUDOR12 ADD DEUDOR12_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    '
    Call ABRESTOCKS
    
    If ValidaExisteCampo%("MOVISTO_ID", "MOVISTO") < 1 Then
        SQLX$ = "ALTER TABLE MOVISTO ADD MOVISTO_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If

    Call CREAR_CAMPOS_EXTRAS_MASTER
    Call GENERAR_ESTRUCTURAS_SQL_INSTITUTO
    Call CREATABLA_ENCAPRF
    
    
    


End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Close
   End
End Sub

Private Sub mnubackup_Click()
   Form38.Show 1
   Exit Sub
   MsgBox "PARA REALIZAR EL BACKUP NADIE DEBE ESTAR UTILIZANDO EL SISTEMA"
   Backup.Show 1
End Sub


Private Sub mnuClientes_Click()
End Sub

Private Sub mnuaccrec_Click()
     If DERACCE%("RECIBO10", "INGRESO A MÓDULO DE RECIBOS") < 2 Then Exit Sub

     ACONEC$ = "RECIBO10"
     Call CONECRUN(ACONEC$, "Generación de Recibo")

End Sub

Private Sub mnuActivarCurso_Click()
   If DERACCE%("ALTANUHO", "ALTA DE NUEVOS HORARIOS") < 2 Then Exit Sub

   ACTIVACURSO.Show 1
End Sub

Private Sub mnuAltaAlumnoACurso_Click()
   If DERACCE%("ALTAACUR", "ALTA DE ALUMNO A CURSO") < 2 Then Exit Sub

   ASIGCURSO.Show 1
End Sub

Private Sub mnualumnos_Click()
  If DERACCE%("MAESALUM", "ACCESO A MAESTRO DE ALUMNOS") < 2 Then Exit Sub
  FRMALUMNO1.Show 1
End Sub

Private Sub MNUANURECIBO_Click()
   ANURECIBO.Show 1
End Sub

Private Sub mnucaca_Click()
   
End Sub

Private Sub mnuAreas_Click()
'  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
'  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
'  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"
'
'  Dim obj As New RECORXET
'  Set RS1 = obj.RS_ENCAPRFLISTAS()
'  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
'  Set obj = Nothing
'  RS1.Close
'  Set RS1 = Nothing
'  EPAC$ = UCase$(TRIM$(EMPREAC$))
'  '
'
'  FRMZELECHO.Width = 8500
'  FRMZELECHO.Show 1
'  IDLIST& = XVALO(RESP_ZELE_VECT(1))
'  If InStr(EPAC$, "ENVAPLAST") > 0 Then
'    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
'    If LIXTA$ = "MATERIALES ENVAPLAST" Then
'       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
'       Exit Sub
'    End If
'  End If

  Call ABM_TABLA(IDLIST&, "AREAS")

End Sub

Private Sub mnuAlumxDia_Click()
    '
    If DERACCE%("REPOALDI", "REPORTE DE ALUMNOS POR DIA") < 2 Then Exit Sub

    'MUESTRA ZELECHO DE DIAS
    CAMPOS$ = "ID/f8;Día/A24"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "SELECCIONE EL DÍA"

    Dim obj As New RECORXET
    Set RS1 = obj.DIASSEMANATEX()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    Set obj = Nothing
    '
    FRMZELECHO.Width = 8500
20  FRMZELECHO.Show 1
    DIAXX$ = RESP_ZELE_VECT(2)
    If DIAXX$ = "" Then Exit Sub


    CAMPOS$ = "ID/f8;RESPONSABLE"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "SELECCIONE EL RESPONSABLE O TODOS"
    

    Dim obj1 As New RECORXET
    Set RS1 = obj.RS_RESPONSABLES_CME()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    Set obj1 = Nothing
    '
    FRMZELECHO.Width = 8500
25  FRMZELECHO.Show 1
    ID_RESPON% = XVALO(RESP_ZELE_VECT(1))
    DESC_RESPON$ = TRIM$(UCase$(RESP_ZELE_VECT(2)))
    If ID_RESPON% < 1 Then GoTo 20
    

    'CARGO EN UN VECTOR LA RAZON SOCIAL DE LOS PROFESORES
    Dim NAMEPROFE()
    SQLX$ = "SELECT NUME_CLI,RASO_CLI FROM PROFESORES WITH(NOLOCK) ORDER BY NUME_CLI ASC"
    Set rs = READSET(SQLX$)
    Do While Not rs.EOF
      NP = XVALO(rs!NUME_CLI)
      RS1 = SAFETEXT$(rs!RASO_CLI)
      ReDim Preserve NAMEPROFE(NP)
      NAMEPROFE(NP) = RS1
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    '
    Dim ALMN As New clientes
    ALMN.cargarvector
    '
    'TRAE TODOS LOS CURSOS QUE ESTEN VIGENTES. (QUE LA FECHA DE FINALIZACION NO HAYA SUPERADO LA FECHA DE HOY)
    SQLX$ = " select a.nume_cli, m.codint,m.descricurso,m.feinicurs,m.fefincurs,m.cantmeses,m.h_desde,m.h_hasta,m.desc_dia1,m.desc_dia2,desc_dia3 ,"
    SQLX$ = SQLX$ + " id_profesor,id_aula from alumno_curso  a WITH(NOLOCK) inner join master m on a.cursos_id = m.codint"
    SQLX$ = SQLX$ + " Where datediff(D, m.FeFinCurs, GETDATE()) <= 1 AND (M.STATUS >= 0 OR M.STATUS IS NULL) "
    SQLX$ = SQLX$ + " and (desc_dia1 = '" & DIAXX$ & "' or desc_dia2 = '" & DIAXX$ & "' or desc_dia3 = '" & DIAXX$ & "')"
    If DESC_RESPON$ <> "TODOS" Then
       SQLX$ = SQLX$ + " and COD_RESPONSABLE = " & ID_RESPON%
    End If
    SQLX$ = SQLX$ + " order by M.CODINT,a.nume_cli,h_desde_fn"
    CAMPOS$ = "Nombre Alum./A20;ID/f6;Curso/A24;Inicia/D8;Termina/d8;Meses/F5;Hor.Desde/a10;Hor.Hasta/a10;Día-S/a12;Profesor/A16;Aula/A10;Pendiente/F10.2;NC/F0"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir Grilla"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.Height = Screen.Height - 2000
    GRILLAGRAL.Width = Screen.Width - 1000
    GRILLAGRAL.NOPINTAR = 1
    GRILLAGRAL.SeleccionaFilaDobleClick% = 1
    '
    'CARGO EL RECORDSET
    j& = 0
    Set RS1 = READSET(SQLX$)
    With RS1
      Do While Not .EOF
         NC = XVALO(!NUME_CLI)
         rs = ALMN.nombre(NC)
         CI1 = XVALO(!CODINT)
         DESCR$ = SAFETEXT(!DESCRICURSO)
         FE% = CVINT(!feinicurs)
         FEH% = CVINT(!FeFinCurs)
         CANMES% = XVALO(!CantMeses)
         DIA1$ = SAFETEXT$(!desc_dia1)
         DIA$ = DIA1$
         DIA2$ = TRIM$(SAFETEXT$(!desc_dia2))
         DIA3$ = TRIM$(SAFETEXT$(!desc_dia3))
         'If DIA2$ <> "" Then DIA$ = DIA$ & " Y " & DIA2$
         DIA$ = DIA1$: If TRIM$(DIA2$) <> "" Then DIA$ = DIA$ & " Y " & DIA2$
         If TRIM$(DIA3$) <> "" Then DIA$ = DIA1$: If TRIM$(DIA2$) <> "" Then DIA$ = DIA$ & " , " & DIA2$ & " Y " & DIA3$
         NP = XVALO(!ID_PROFESOR)
         NOMBREPROFE$ = NAMEPROFE(NP)
         IDAULA& = XVALO(!ID_AULA)
         j& = j& + 1
         GRILLAGRAL.GRID.Rows = j& + 1
         GRILLAGRAL.GRID.TextMatrix(j&, 1) = rs
         GRILLAGRAL.GRID.TextMatrix(j&, 2) = CI1
         GRILLAGRAL.GRID.TextMatrix(j&, 3) = DESCR$
         GRILLAGRAL.GRID.TextMatrix(j&, 4) = FECHATEX$(FE%)
         GRILLAGRAL.GRID.TextMatrix(j&, 5) = FECHATEX$(FEH%)
         GRILLAGRAL.GRID.TextMatrix(j&, 6) = CANMES%
         GRILLAGRAL.GRID.TextMatrix(j&, 7) = SAFETEXT$(!H_DESDE)
         GRILLAGRAL.GRID.TextMatrix(j&, 8) = SAFETEXT$(!H_HASTA)
         GRILLAGRAL.GRID.TextMatrix(j&, 9) = DIA$
         
         GRILLAGRAL.GRID.TextMatrix(j&, 10) = NOMBREPROFE$
         GRILLAGRAL.GRID.TextMatrix(j&, 11) = AULA$(IDAULA&)
         GRILLAGRAL.GRID.TextMatrix(j&, 12) = SALDOAFECHAX#(NC, HOY(HO$))
         GRILLAGRAL.GRID.TextMatrix(j&, 13) = NC
         
      .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then
      Call COMUNI("NO SE DETECTARON CURSOS PARA EL DÍA Y/O RESPONSABLE SELECCIONADO")
      Exit Sub
    End If
    
    TX$ = "ALUMNOS FILTRADO POR DÍA"
    GRILLAGRAL.Caption = TX$
    '*******************************************
30  GRILLAGRAL.Show 1
    If GRILLAGRAL.MENU1 > 0 Then
        Dim MSFX As MSFlexGrid
        Set MSFX = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSFX = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSFX, CAMPOS$, "PRESENTES", "LISTADO DE PRESENTES", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
    End If
    
     If GRILLAGRAL.SeleccionaFilaDobleClick > 0 Then
         LEGAJO& = XVALO(RESP_ZELE_VECT(13))
         ACONEC$ = "RECIBO10/RECICUTA" + SAFETEXT$(LEGAJO&) ' NUME
         Call CONECRUN(ACONEC$, "Generación de Recibo")
    End If
   
    GoTo 20 ' VUELVE A LISTA DE SELECCION DE DIAS

     
End Sub

Private Sub mnucaja_Click()
   If DERACCE%("MANECAJA", "RESTRICCIÓN DE INGRESO A MANEJO DE CAJAS") < 2 Then Exit Sub
   CAJADET.Show 1
End Sub

Private Sub mnucopiatabla_Click()
     If CLACONTRA2% < 1 Then Exit Sub
     
     CACAM% = BaDaABRE_ADMIN_ALQ.TableDefs("USUARIOS").Fields.Count
     'Master.FindFirst ("CODINT = " & CII%)
     SQLX$ = "SELECT * FROM USUARIOS "
     Set rs = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenForwardOnly)

     SQLX$ = "SELECT * FROM USUARIOS "
     Dim RS2 As New ADODB.Recordset
     RS2.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText

     With rs
     Do While Not .EOF
        RS2.AddNew
        For KU% = 0 To CACAM% - 1
           NOMCAMPO$ = RS2.Fields(KU%).Name
           NOMCAMPO2$ = rs.Fields(KU%).Name
           If NOMCAMPO$ = NOMCAMPO2$ Then
              RS2.Fields(KU%).Value = rs.Fields(KU%).Value
           Else
              Call COMUNI("NO COINCIDE EL CAMPO")
           End If
        Next KU%
        RS2.Update
        .MoveNext
     Loop
     End With
     rs.Close
     Set rs = Nothing
     RS2.Close
     Set RS2 = Nothing
       

End Sub

Private Sub mnuClave_Click()
   Call FSETUPNEW.Command5_Click
End Sub

Private Sub mnuCuotas_Click()
   If DERACCE%("ALTACUOT", "GENERACION DE CUOTAS") < 2 Then Exit Sub

   ACTUCUOTA.Show 1
End Sub

Private Sub mnuCursos_Click()

    If DERACCE%("ALTCURPA", "ALTA DE CURSOS PADRE") < 2 Then Exit Sub

    Unload VENTABNEW
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "LISTADO DE CURSOS"
    VENTANA.CampEditables = "2;4;5;6;7"
    VENTANA.ColEcoCampo(2) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'AREAS'"
    VENTANA.Inicializar = 1
    Campox$ = "C.Curso/F6;C.Area/F5;Area/A30;Desc.Curso/A58;Val.Cuota/F10.2;Matrícula/F10.2;Der. Examen/F10.2"
    VENTANA.Campox = Campox$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS()
    Call Carga_MSF_Recordset(RS1, Campox$, VENTABNEW.MSF)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
30  VENTABNEW.Show 1
    
    
    If VENTABNEW.APROBADO% > 0 Then
       'VALIDO QUE TENGA CANTIDAD
       For i& = 1 To VENTABNEW.MSF.Rows - 1
         If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 2)) = "" Then
            MsgBox "Falta Informar Descripción en Fila " & i&
            GoTo 30
         End If
       Next i&
       For i& = 1 To VENTABNEW.MSF.Rows - 1
         If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 3)) = "" Then
            MsgBox "Falta Informar Area en Fila " & i&
            GoTo 30
         End If
       Next i&

       USNX% = USNBR% Mod 100
       For i& = 1 To VENTABNEW.MSF.Rows - 1
         If XVALO(VENTABNEW.MSF.TextMatrix(i&, 1)) > 0 Then
            SQLX$ = "UPDATE CURSOS SET COD_AREA = '" & VENTABNEW.MSF.TextMatrix(i&, 2) & "'"
            SQLX$ = SQLX$ + " ,AREA = '" & UCase$(VENTABNEW.MSF.TextMatrix(i&, 3)) & "'"
            SQLX$ = SQLX$ + " ,DESCRIPCION = '" & UCase$(VENTABNEW.MSF.TextMatrix(i&, 4)) & "' "
            SQLX$ = SQLX$ + " ,USUARIO = " & USNX% & " "
            SQLX$ = SQLX$ + " ,ValCurso = " & ROUND(XVALO(VENTABNEW.MSF.TextMatrix(i&, 5)), 2)
            SQLX$ = SQLX$ + " ,ValMatricula = " & ROUND(XVALO(VENTABNEW.MSF.TextMatrix(i&, 6)), 2)
            SQLX$ = SQLX$ + " ,ValDerExamen = " & ROUND(XVALO(VENTABNEW.MSF.TextMatrix(i&, 7)), 2)
            SQLX$ = SQLX$ + "  WHERE CURSOS_ID = " & XVALO(VENTABNEW.MSF.TextMatrix(i&, 1))
            FLEXCONN.Execute (SQLX$)
         Else
            SQLX$ = "INSERT INTO CURSOS"
            SQLX$ = SQLX$ + " (COD_AREA,AREA,DESCRIPCION,USUARIO,ValCurso,ValMatricula,ValDerExamen)"
            SQLX$ = SQLX$ + " VALUES(" & UCase(VENTABNEW.MSF.TextMatrix(i&, 2))
            SQLX$ = SQLX$ + " ,'" & UCase(VENTABNEW.MSF.TextMatrix(i&, 3)) & "'"
            SQLX$ = SQLX$ + " ,'" & UCase(VENTABNEW.MSF.TextMatrix(i&, 4)) & "'"
            SQLX$ = SQLX$ + " ," & USNX%
                        
            SQLX$ = SQLX$ + " ," & ROUND(VENTABNEW.MSF.TextMatrix(i&, 5), 2)
            SQLX$ = SQLX$ + " ," & ROUND(VENTABNEW.MSF.TextMatrix(i&, 6), 2)
            SQLX$ = SQLX$ + " ," & ROUND(VENTABNEW.MSF.TextMatrix(i&, 7), 2)

            SQLX$ = SQLX$ + ")"
            FLEXCONN.Execute (SQLX$)
         End If
       Next i&
    End If
          
       Set VENTANA = Nothing
       Unload VENTABNEW
End Sub

Private Sub mnuderuser_Click()
   NUMEUSER% = USNBR% Mod 100
   If NUMEUSER% <> 1 Then
      MsgBox "Solo El Supervisor del Sistema Pueda Ingresar"
      Exit Sub
   End If
   FSETUPNEW.Command16_Click


End Sub

Private Sub mnudsg_Click()
   If DERACCE%("DSGFRM09", "ACCESO A DISEÑADOR DE FORMULARIOS") < 2 Then Exit Sub
   If CLACONTRA% > 0 Then
      Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
   End If
End Sub

Private Sub mnugenerep_Click()
   If DERACCE%("GENREP07", "Generador de Reportes SQL") < 2 Then Exit Sub
   If CLACONTRA% > 0 Then
     Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   End If
End Sub

Private Sub mnuhist_Click()
    
''    HOO% = HOY(HOS$)
''    FECHA_ACTUAL$ = FECHATEX$(HOO%)
''    FECHA_VENCIMIENTO$ = "30/09/12"
''    If FECHANUM(FECHA_ACTUAL$) > FECHANUM(FECHA_VENCIMIENTO$) Then
''       TX$ = "MODULO HISTORIAL A MODO DE PRUEBA" + vbCrLf
''       TX$ = TX$ + "HA VENCIDO EL DÍA 30/09/12" + vbCrLf
''       TX$ = TX$ + "EN CASO DE HABILITACIÓN DEL MISMO COMUNICARSE CON EL PROGRAMADOR" + vbCrLf
''       MsgBox TX$
''       Exit Sub
''    End If

   SEGUIM_COLEGIO1.Show 1
End Sub

Private Sub mnuimpo_Click()

    MsgBox "ATENCION ESTE PROCESO ESTA RESERVADO PARA EL DISEÑADOR DEL SISTEMA"
10  CLAVE$ = TRIM$(InputBox$("INGRESE LA CONTRASEÑA ", , CLAVE$))
    '
    'GUARDA LA DIRECCION DE CORREO
    If TRIM$(CLAVE$) = "" Then Exit Sub
    If TRIM$(UCase$(CLAVE$)) <> "OMIEDI" Then GoTo 10
    
    
    IMPORTADOR.Show 1
 
End Sub

Private Sub MNUMAT_Click()
   RPT_MATR_GRADO.Show 1
End Sub

Private Sub mnumatricula_Click()
   REPORTE_GENERAL.Show 1
End Sub

Private Sub mnunuevo_Click()
   
   If DERACCE%("STPGESVE", "Configuracion de Setup de Ventas") = 2 Then
      If CLACONTRA% > 0 Then
        OPCLAV07.Show 1
        Call FINAL("")
      End If
   End If
End Sub
Sub CARGACUENTAS()
    
    CARGATABLA.TABLA1$ = "CUENTAS" ' ES LA TABLA
    CARGATABLA.CAMPO1$ = "COD_CUENTA"  ' ES EL CAMPO (VARIABLES A NIVEL DE FRM)
    CARGATABLA.CAMPO2$ = "REFERENCIA"     ' ES EL CAMPO DE CODIGO DE LA REFERENCIA
    CARGATABLA.CAMPO3$ = "STATUS"
    
    CARGATABLA.Caption = "CARGA TABLA DE CUENTAS CONTABLES"

    CARGATABLA.Show 1
End Sub
Private Sub mnuPendientes_Click()
   PENDIEN_COLE.Show 1
End Sub

Private Sub mnurecibo_Click()

End Sub

Private Sub mnupendcob_Click()
    SQLX$ = "SELECT * FROM SADEUPEN WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Status < 3 "
    SQLX$ = SQLX$ + " ORDER BY FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    Dim SACTEMP As ADODB.Recordset
    Set SACTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    On Error Resume Next
    FLIM = HOY(HOS$)
    With SACTEMP
      .MoveFirst
      If Err < 1 Then
        FLIM = CVINT(!FEVENCIM)
      End If
    End With
    SACTEMP.Close
    Set SACTEMP = Nothing
    On Error GoTo 0
    
    '
    FECHACLI.Text4 = ""
    DESX$ = FECHATEX$(HOY(HO$))
    DESX$ = "01" & Mid$(DESX$, 3)
    DESX1% = FECHANUM(DESX$)
    FECHACLI.Desde.Value = CVDAT(DESX1%)
    FECHACLI.Hasta.Value = CVDAT(HOY(HO$))
5   FECHACLI.Show 1
    '
    If IsFormLoaded(FECHACLI) = False Then
        Exit Sub
    End If

    DES% = CVINT(FECHACLI.Desde)
    HAS% = CVINT(FECHACLI.Hasta)
    If DES% < 33 Or HAS% < 33 Then
       Call COMUNI("Fecha Inválida")
       GoTo 5
    End If
    If DES% > HAS% Then
       Call COMUNI("Selección de Fecha Inválida")
       GoTo 5
    End If
    '
    CLISEL = XVALO(FECHACLI.Text4)
    

   
    CAMPOS$ = "Lgjo/F6;Nombre Alum./A20;Tele/a12;Documento/A30;Curso/A30;F.Emis/D8;Imp.Orig/f12.2;F.Venc./D8;Saldo/F12.2"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir Grilla"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    GRILLAGRAL.Height = Screen.Height - 2000
    GRILLAGRAL.Width = Screen.Width - 1000
    GRILLAGRAL.NOPINTAR = 1
'    GRILLAGRAL.SeleccionaFilaDobleClick% = 1
    ACUMUSALDEU# = 0
    SQLX$ = "SELECT * FROM SADEUPEN with(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE SADEUPEN.Status < 3 "
    SQLX$ = SQLX$ + "AND (ABS(SADEUPEN.SALDdebe - SADEUPEN.SALDAcre) >= 0.005 "
    SQLX$ = SQLX$ + "OR ABS(SADEUPEN.SALDdebe_Dola - SADEUPEN.SALDAcre_Dola) >= 0.005) "
    If CLISEL > 0 Then SQLX$ = SQLX$ + "AND SADEUPEN.NuClien = " & CLISEL & " "
    SQLX$ = SQLX$ + "AND SADEUPEN.FeVencim >= '" & FETEXCOR1$(DES%) & "' "
    SQLX$ = SQLX$ + "AND SADEUPEN.FeVencim <= '" & FETEXCOR1$(HAS%) & "' "
    SQLX$ = SQLX$ + "AND SADEUPEN.CODIEMPR = " & CodiEmp% & " "
    SQLX$ = SQLX$ + "ORDER BY SADEUPEN.FeVencim DESC "
    '
    Call APERBASDAT("CABAN")
    '
    j& = 0
    Set SACTEMP = READSET(SQLX$)
    With SACTEMP
        Do While Not .EOF
          UL& = UL& + 1
          Call TRACE(20, UL&, FIN&)
          NCLI = XVALO(!NUCLIEN)
          NFS$ = SAFETEXT$(!CodiCom_Cor)
          COCOLAR$ = SAFETEXT$(!CODICOM_LAR)
          FF% = CVINT(!FECHEMISI)
          FVI% = CVINT(!FEVENCIM)
          SALDEU# = XVALO(!SaldDebe) - XVALO(!SALDAcre)
          SALDOL# = 0
          On Error Resume Next
          SALDOL# = XVALO(!SaldDebe_DOLA) - XVALO(!SaldAcre_DOLA)
          On Error GoTo 0
          '
          CONDI$ = "CodiCom_Lar = '" & COCOLAR$ & "'"
          CONDI$ = CONDI$ + " AND (TipoMovim = 'FVEN') AND NUCLIEN =  " & NCLI
          If Left$(COCOLAR$, 3) = "FC-" Then
                CUOTA$ = SAFETEXT$(VALOBADA("MOVISTO", "IDENMOVI", "DOSECLAR='" & COCOLAR$ & "' AND Nume_Clie=" & NCLI))
                TICURSO$ = SAFETEXT$(VALOBADA("CAJABANC", "OBSERGEN", "CODICOM_LAR='" & COCOLAR$ & "' AND NUCLIEN=" & NCLI))
          End If

          If MODOMONEI% < 2 Then
             If Abs(SALDEU#) < 0.005 Then
                SALDOL# = 0
                GoTo 29
             End If
          End If
          IMORIG# = !IBruComp
          ES$ = "    "
          MONEMX% = 1
          On Error Resume Next
          MONEMX% = !Mone_Emis
          
          On Error GoTo 0
          '
          j& = j& + 1
          GRILLAGRAL.GRID.Rows = j& + 1
          GRILLAGRAL.GRID.TextMatrix(j&, 1) = NCLI
          GRILLAGRAL.GRID.TextMatrix(j&, 2) = rasocli$(NCLI)
          GRILLAGRAL.GRID.TextMatrix(j&, 3) = TEL1CLI$(NCLI)
          GRILLAGRAL.GRID.TextMatrix(j&, 4) = CUOTA$
          GRILLAGRAL.GRID.TextMatrix(j&, 5) = TICURSO$
          GRILLAGRAL.GRID.TextMatrix(j&, 6) = FECHATEX$(FF%)
          GRILLAGRAL.GRID.TextMatrix(j&, 7) = FORMATNUM$(IMORIG#, "f12.2")
          GRILLAGRAL.GRID.TextMatrix(j&, 8) = FECHATEX$(FVI%)
          GRILLAGRAL.GRID.TextMatrix(j&, 9) = FORMATNUM$(SALDEU#, "f12.2")
          ACUMUSALDEU# = ACUMUSALDEU# + SALDEU#
29        .MoveNext
        Loop
     End With
     GRILLAGRAL.Caption = "Pendientes de Cobranza (Total: " & TRIM$(FORMATNUM$(ACUMUSALDEU#, "f12.2")) & ")"
30   GRILLAGRAL.Show 1
    If GRILLAGRAL.MENU1 > 0 Then
        Dim MSFX As MSFlexGrid
        Set MSFX = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSFX = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSFX, CAMPOS$, "PRESENTES", "PENDIENTE DE COBRANZA", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
    End If
    GoTo 5
99
End Sub

Private Sub mnuPresentes_Click()
   If DERACCE%("INFOPRES", "GENERAR INFORMACION DE PRESENTES") < 2 Then Exit Sub
   FRMPRESENTE.Show 1
End Sub

Private Sub mnuSetup_Click()
    FSETUPNEW.Show 1
End Sub

Private Sub mnuTablasAux_Click()
'  Call GENERAR_ESTRUCTURAS_SQL
  If DERACCE%("MODTABAU", "MODIFICAR TABLAS AUXILIARES") < 2 Then Exit Sub

  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set RS1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
  Set obj = Nothing
  RS1.Close
  Set RS1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub mnutot_cob_Click()
   TOT_COBRANZAS.Show 1
End Sub

Private Sub mnuuser_Click()
    If DERACCE%("FSETUP07/", "CONFIGURACION GENERAL") < 2 Then GoTo 99
 
    FSETUPNEW.Command16_Click
   
'   If DERACCE1%("ABMUSUAR", "ABM DE USUARIOS") > 1 Then
'     ABMUSUARIOS.Show 1
99
End Sub
Sub CARGALISTAUSUARIOS()
    Call ABRECONEXION
    '
    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NUMEUSUARIO > 0"
    CNRG& = CantRegistros&("USUARIOS", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Usuarios Cargados"
      Exit Sub
    End If
    '
    JJ& = 0
    FRMBUSCAR.List1.Clear

    SQLX$ = "SELECT * FROM USUARIOS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMEUSUARIO > 0"
    SQLX$ = SQLX$ + " AND (STATUS >= 0 OR STATUS IS NULL)"
    SQLX$ = SQLX$ + " ORDER BY NUMEUSUARIO"
    '
    Set MENSATMP = BASSSEX.Execute(SQLX$)
    With MENSATMP
      Do While Not .EOF
       COCU$ = XVALO(!NUMEUSUARIO)
       DESCRI$ = SAFETEXT$(!nombre) + " " + SAFETEXT$(!APELLIDO)
       Call REPLA(TTXX$, COCU$, 1, 10)
       Call REPLA(TTXX$, DESCRI$, 12, 40)
       JJ& = JJ& + 1
       FRMBUSCAR.List1.AddItem TTXX$
       .MoveNext
      Loop
    End With
    
    MENSATMP.Close
    Set MENSATMP = Nothing
    '
    FRMBUSCAR.Caption = "DERECHOS"
    FRMBUSCAR.Show

End Sub

Private Sub mnuvarias_Click()
'  If DERACCE1%("ACCIVARI", "ACCIONES VARIAS DE RECIBOS") > 1 Then
'     EDIRECIBO.Show 1
'  End If
End Sub

Private Sub mnuver_Click()
   VERRECIBOS.Show 1
End Sub

