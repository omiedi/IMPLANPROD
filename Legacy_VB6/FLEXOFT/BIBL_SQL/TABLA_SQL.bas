Attribute VB_Name = "TABLA_SQL"
Sub CARUBIPRE()
  '\\\OT-08-1041-OD-13/11/08///
   Call SETULTREG("UBIPRE", 0)
   I& = 0
   Call ABRESTOCKS
   '
   'RECORRO LA TABLA Y GRABO EN EL ARCHIVO UBIPRE
   '
   Call SELECHO("TADEPOSI")
   DEPOO$ = VALACT1$("TADEPOSI")
   DEPOSI$ = VALACT2$("TADEPOSI")
   DEPO% = XVALO(DEPOO$)
   If DEPO% < 1 Then Exit Sub
   '
   Call ACTUREGISTRO("UBIPRE", "DEPOUBIC", "DEPOUBIC IS NULL", 1, REGS&, "NOMESS")
   '
   SQLX$ = "SELECT * FROM UBIPRE WHERE DEPOUBIC = " & DEPO% & "ORDER BY codiubic"
   Dim UBIPRETMP As ADODB.Recordset
   Set UBIPRETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With UBIPRETMP
     Do While Not .EOF
       I& = I& + 1
       CODD$ = SAFETEXT$(!CODIUBIC)
       REFER$ = SAFETEXT$(!REFERENCIA)
       STAT% = XVALO(!Status)
       Call REPLA(TX$, CODD$, 1, 16)
       Call REPLA(TX$, REFER$, 17, 64)
       Call REPLA(TX$, MKI$(STAT%), 81, 2)
       Call GRAREG("!UBIPRE", TX$, I&)
     .MoveNext
     Loop
   End With
   '
   Call SETULTREG("!UBIPRE", I&)
   '
   ' MUESTRO LA VENTANA
10 VTB% = VENTABU%("UBIPRE")
   If VTB% < 0 Then Exit Sub
   '
   'MARCA TODOS LOS REGISTROS  PARA BORRAR CON 1
   SQLX$ = "Update UBIPRE SET Marborra = 1 WHERE DEPOUBIC = " & DEPO%
   FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
   For u& = 1 To ULTREG&("!UBIPRE")
     X$ = REGLEIDO$("!UBIPRE", u&)
     STATT% = CVI(Mid$(X$, 81, 2))
     If STATT% < 0 Or STATT% > 1 Then
        Call MENSERR(24, "El Status debe ser '0' o '1' ")
        GoTo 10
     End If
   Next u&
   '
   SQLX$ = "Select * From UBIPRE"
   '
   Set UBIPRETMP = New ADODB.Recordset
   UBIPRETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   'RECORRE LA VENTANA
   For u& = 1 To ULTREG&("!UBIPRE")
      X$ = REGLEIDO$("!UBIPRE", u&)
      CODI$ = Mid$(X$, 1, 16)
      '
      If CODI$ <> "" Then
        REFFF$ = Mid$(X$, 17, 64)
        STATT% = CVI(Mid$(X$, 81, 2))
        '
        With UBIPRETMP
          .AddNew  ' AGREGA NUEVA UBICACION FISICA   ' YA CONVERTIDO
              On Error Resume Next
          !CodiEmpr = CodiEmp%
              On Error GoTo 0
          !DEPOUBIC = XVALO(DEPO%)
          !CODIUBIC = SAFETEXT(CODI$)
          !REFERENCIA = SAFETEXT(REFFF$)
          !Status = XVALO(STATT%)
          !MarBorra = 0
          .Update
        End With
        '
      End If
   Next u&
   UBIPRETMP.Close
   Set UBIPRETMP = Nothing
   '
   ' BORRO TODOS LOS QUE FUERON MARCADOS CON 1
   SQLX$ = "DELETE FROM UBIPRE WHERE Marborra > 0 "
   FLEXCONN.Execute (FILTSQL$(SQLX$))
   '
   Call CIERRARCH("UBIPRE")
   Call CIERRARCH("!UBIPRE")
   '\\\OT-08-1041-OD-FIN///
End Sub
'

Sub CARTAVENDE()
    '
    Call ABRECLIEN
25  Call SELECHO("VENDEDOR")
    OP$ = VALACT1$("VENDEDOR")
    NUVENDD% = XVALO(OP$)
    VENDD$ = TRIM$(VALACT2$("VENDEDOR"))
    CONDI$ = "Vend_Cli = " & NUVENDD%
    REGSS& = CantRegistros&("VENDEDOR", CONDI$)
    If REGSS& < 1 Then
        Call CreaRegistro("VENDEDOR", "Vend_Cli", NUVENDD%, "NOMESS")
        Call ACTUREGISTRO("VENDEDOR", "RasoVend", CONDI$, VENDD$, REGS&, "NOMESS")
        X$ = FILTERGETRECORD("DIMAIVEN", 1, Chr$(NUVENDD%))
        DIMAI$ = TRIM$(Mid$(X$, 2, 63))
        If DIMAI$ <> "" Then
            Call ACTUREGISTRO("VENDEDOR", "EmailVend", CONDI$, DIMAI$, REGS&, "NOMESS")
        End If
    End If
    If NUVENDD% > 0 Then
        DOMVEND$ = VALOBADA("VENDEDOR", "DomiVend", CONDI$, "NOMESS")
        LOCVEND$ = VALOBADA("VENDEDOR", "LocaVend", CONDI$, "NOMESS")
        TEL1VEND$ = VALOBADA("VENDEDOR", "TeleVend_1", CONDI$, "NOMESS")
        TEL2VEND$ = VALOBADA("VENDEDOR", "TeleVend_2", CONDI$, "NOMESS")
        MAILVEND$ = VALOBADA("VENDEDOR", "EmailVend", CONDI$, "NOMESS")
        OBSERVEND$ = VALOBADA("VENDEDOR", "Observaciones", CONDI$, "NOMESS")
        '
        Call REPLA(VDEF$, DOMVEND$, 1, 48)
        Call REPLA(VDEF$, LOCVEND$, 49, 48)
        Call REPLA(VDEF$, TEL1VEND$, 97, 32)
        Call REPLA(VDEF$, TEL2VEND$, 129, 32)
        Call REPLA(VDEF$, MAILVEND$, 161, 48)
        Call REPLA(VDEF$, MAILVEND$, 209, 63)

        VDEF$ = OBJPLA$("DATCOVEND", VDEF$)
        If VDEF$ <> "" Then
            Call ACTUREGISTRO("VENDEDOR", "DomiVend", CONDI$, Mid$(VDEF$, 1, 48), REGS&, "NOMESS")
            Call ACTUREGISTRO("VENDEDOR", "LocaVend", CONDI$, Mid$(VDEF$, 49, 48), REGS&, "NOMESS")
            Call ACTUREGISTRO("VENDEDOR", "TeleVend_1", CONDI$, Mid$(VDEF$, 97, 32), REGS&, "NOMESS")
            Call ACTUREGISTRO("VENDEDOR", "TeleVend_2", CONDI$, Mid$(VDEF$, 129, 32), REGS&, "NOMESS")
            Call ACTUREGISTRO("VENDEDOR", "EmailVend", CONDI$, Mid$(VDEF$, 209, 63), REGS&, "NOMESS")
            Call ACTUREGISTRO("VENDEDOR", "Observaciones", CONDI$, Mid$(VDEF$, 161, 48), REGS&, "NOMESS")
        End If
        GoTo 25
        Else
               Exit Sub
        End If
End Sub

Sub CARMAQUI()
    '
    If EXISTE%("EQUPRO09.EXE") Then
       Call CONECRUN("EQUPRO09.EXE", "Equipos de Produccion")
       Exit Sub
    End If
    '
'    Call ABRECONEXION
'    Call CREATABLA_MAQUINARIAS
'    SQLX$ = "SELECT * FROM EQUIPOS WITH(NOLOCK)"
'    Dim EQSEL As ADODB.Recordset
'    Set EQSEL = New ADODB.Recordset
'    Set EQSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    With EQSEL
'       EQAGRE% = 0
'       Do While Not .EOF
'          CODEQ1$ = SAFETEXT$(!Codigo_Equ)
'          DESEQ$ = SAFETEXT$(!Descri_Equ)
'          SECEQ = XVALO(!Ubica_Equ)
'          XXX$ = TRIM$(FILTERGETRECORD("PADMAQ", 1, CODEQ1$))
'          If XXX$ = "" Then
'             XXX$ = REGBLAN$("PADMAQ")
'             Call REPLA(XXX$, CODEQ1$, 1, 6)
'             Call REPLA(XXX$, DESEQ$, 7, 30)
'             Call REPLA(XXX$, MKS$(SECEQ), 37, 4)
'             Call REGAPP("PADMAQ", XXX$)
'             EQAGRE% = 1
'          End If
'       .MoveNext
'       Loop
'    End With
'    '
'    If EQAGRE% > 0 Then
'       Call CIERRARCH("PADMAQ")
'       Call FILESORT("PADMAQ", "", 1, 1, "ASC")
'       Call COMUNI("Unificados Equipos con Tabla de Mantenimiento")
'    End If
    '
    j& = 0
    For u& = 1 To ULTREG&("PADMAQ")
        X$ = REGLEIDO$("PADMAQ", u&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            j& = j& + 1
            Call GRAREG("!PADMAQ", X$, j&)
        End If
    Next u&
    Call SETULTREG("!PADMAQ", j&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
    VTB% = VENTABU%("CARMAQUI")
    If VTB% < 0 Then Exit Sub
    '
    j& = 0
    For u& = 1 To ULTREG&("!PADMAQ")
        X$ = REGLEIDO$("!PADMAQ", u&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        
        If COMAQ$ <> "" Then
            j& = j& + 1
            Call GRAREG("PADMAQ", X$, j&)
            '
            If CantRegistros("EQUIPOS", "CODIGO_EQU = '" & COMAQ$ & "'") < 1 Then
               Call CreaRegistro("EQUIPOS", "CODIGO_EQU", COMAQ$, "NOMESS")
            End If
            DEMAQ$ = TRIM$(Mid$(X$, 7, 30))
            Call ACTUREGISTRO("EQUIPOS", "DESCRI_EQU", "CODIGO_EQU = '" & COMAQ$ & "'", DEMAQ$, RAFE&)
            '
        End If
    Next u&
    '
    Call SETULTREG("PADMAQ", j&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    Call APAREA_TABLA_PADMAQSQL
    '
End Sub

Sub APAREA_TABLA_PADMAQSQL()
   If EXISTE_TABLA%("DATASQL") = 1 Then
      'MARCO CON MARBORRA
      DESCRITABL$ = "PADRON DE MAQUINAS"
      IDLI& = XVALO(VALOBADA("DATASQL", "ID_LISTA", "LISTA = '" & DESCRITABL$ & "'", "NOMESS"))
      Call ACTUREGISTRO("DATASQL", "MARBORRA", "ID_LISTA = " & IDLI&, 1, REGAF&)
      
      For u& = 1 To ULTREG&("PADMAQ")
        X$ = REGLEIDO$("PADMAQ", u&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        
        If COMAQ$ <> "" Then
            j& = j& + 1
            '
            DEMAQ$ = TRIM$(Mid$(X$, 7, 30))
            VALOX$ = "" & IDLI& & " , '" & DESCRITABL$ & "'"
            VALOX$ = VALOX$ & ", '" & COMAQ$ & "' , '" & DEMAQ$ & "'"
            SQLX$ = "INSERT INTO DATASQL "
            PRICAMP$ = "ID_LISTA, LISTA, Campo1, Campo2"
            SQLX$ = SQLX$ + " (" + PRICAMP$ + RET_CAMPOS$ + ") "
            SQLX$ = SQLX$ + " VALUES (" & VALOX$ & ")"
            FLEXCONN.Execute SQLX$

            '
        End If
    Next u&
    '
    Call SETULTREG("PADMAQ", j&)
    Call CIERRARCH("PADMAQ")
    Call BORRAREGISTROS("DATASQL", "ID_LISTA = " & IDLI& & " AND MARBORRA = 1", REGAF&, "NOMESS")
   End If
End Sub

Function PROVINSEL%()
  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "COD_NUM/F0;COD_TEX/A12;NOMBRE PCIA./A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "COD_NUM/F0;COD_TEX/A12;NOMBRE PCIA./A48", FRMZELECHO)
  FRMZELECHO.Caption = "TABLA DE PROVINICIAS"
  '
  Dim PROVINCIAX(24)
    PROVINCIAX(0) = "C CIUDAD AUTONOMA DE BUENOS AIRES"
    PROVINCIAX(1) = "B BUENOS AIRES"
    PROVINCIAX(2) = "W CATAMARCA"
    PROVINCIAX(3) = "X CORDOBA"
    PROVINCIAX(4) = "I CORRIENTES"
    PROVINCIAX(5) = "E ENTRE RIOS"
    PROVINCIAX(6) = "Y JUJUY"
    PROVINCIAX(7) = "M MENDOZA"
    PROVINCIAX(8) = "K LA RIOJA"
    PROVINCIAX(9) = "S SALTA"
    PROVINCIAX(10) = "J SAN JUAN"
    PROVINCIAX(11) = "L SAN LUIS"
    PROVINCIAX(12) = "Q SANTA FE"
    PROVINCIAX(13) = "G SANTIAGO DEL ESTERO"
    PROVINCIAX(14) = "T TUCUMAN"
    PROVINCIAX(16) = "A CHACO"
    PROVINCIAX(17) = "H CHUBUT"
    PROVINCIAX(18) = "N FORMOSA"
    PROVINCIAX(19) = "O MISIONES"
    PROVINCIAX(20) = "U NEUQUEN"
    PROVINCIAX(21) = "P LA PAMPA"
    PROVINCIAX(22) = "R RIO NEGRO"
    PROVINCIAX(23) = "Z SANTA CRUZ"
    PROVINCIAX(24) = "F TIERRA DEL FUEGO"
    
  Dim obj As New RECORXET
  Set RS1 = obj.RS_TABLA_PROVINCIA()
  Call Carga_MSF_Recordset(RS1, "COD_NUM/F6;COD_TEX/A4;NOMBRE PCIA./A32", FRMZELECHO.msf2)
  Set obj = Nothing
  If obj.RS_TABLA_PROVINCIA.EOF Then
      SQLX$ = "DELETE FROM TAPROVINCIA"
      FLEXCONN.Execute (SQLX$)
      Dim clprovi As New clientes
      For I& = 0 To 24
        XX$ = PROVINCIAX(I)
        codProvinciaFlx$ = TRIM$(Mid$(XX$, 1, 1))
        nombreP$ = TRIM$(Mid$(XX$, 2))
        Call ALTAPROVINCIASSQL(I&, codProvinciaFlx$, nombreP$)
      Next I&
'    End If
  End If
  RS1.Close
  Set RS1 = Nothing
  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  PROVINSEL% = XVALO(RESP_ZELE_VECT(1))

End Function


