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
       REFER$ = SAFETEXT$(!Referencia)
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
   For U& = 1 To ULTREG&("!UBIPRE")
     X$ = REGLEIDO$("!UBIPRE", U&)
     STATT% = CVI(Mid$(X$, 81, 2))
     If STATT% < 0 Or STATT% > 1 Then
        Call MENSERR(24, "El Status debe ser '0' o '1' ")
        GoTo 10
     End If
   Next U&
   '
   SQLX$ = "Select * From UBIPRE"
   '
   Set UBIPRETMP = New ADODB.Recordset
   UBIPRETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   'RECORRE LA VENTANA
   For U& = 1 To ULTREG&("!UBIPRE")
      X$ = REGLEIDO$("!UBIPRE", U&)
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
          !Referencia = SAFETEXT(REFFF$)
          !Status = XVALO(STATT%)
          !MarBorra = 0
          .Update
        End With
        '
      End If
   Next U&
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
    Call ABRECONEXION
    Call CREATABLA_MAQUINARIAS
    SQLX$ = "SELECT * FROM EQUIPOS WITH(NOLOCK)"
    Dim EQSEL As ADODB.Recordset
    Set EQSEL = New ADODB.Recordset
    Set EQSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With EQSEL
       EQAGRE% = 0
       Do While Not .EOF
          CODEQ1$ = SAFETEXT$(!Codigo_Equ)
          DESEQ$ = SAFETEXT$(!Descri_Equ)
          SECEQ = XVALO(!Ubica_Equ)
          XXX$ = TRIM$(FILTERGETRECORD("PADMAQ", 1, CODEQ1$))
          If XXX$ = "" Then
             XXX$ = REGBLAN$("PADMAQ")
             Call REPLA(XXX$, CODEQ1$, 1, 6)
             Call REPLA(XXX$, DESEQ$, 7, 30)
             Call REPLA(XXX$, MKS$(SECEQ), 37, 4)
             Call REGAPP("PADMAQ", XXX$)
             EQAGRE% = 1
          End If
       .MoveNext
       Loop
    End With
    '
    If EQAGRE% > 0 Then
       Call CIERRARCH("PADMAQ")
       Call FILESORT("PADMAQ", "", 1, 1, "ASC")
       Call COMUNI("Unificados Equipos con Tabla de Mantenimiento")
    End If
    '
    J& = 0
    For U& = 1 To ULTREG&("PADMAQ")
        X$ = REGLEIDO$("PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("!PADMAQ", X$, J&)
        End If
    Next U&
    Call SETULTREG("!PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
    VTB% = VENTABU%("CARMAQUI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!PADMAQ")
        X$ = REGLEIDO$("!PADMAQ", U&)
        COMAQ$ = TRIM$(Left$(X$, 6))
        
        If COMAQ$ <> "" Then
            J& = J& + 1
            Call GRAREG("PADMAQ", X$, J&)
            '
            If CantRegistros("EQUIPOS", "CODIGO_EQU = '" & COMAQ$ & "'") < 1 Then
               Call CreaRegistro("EQUIPOS", "CODIGO_EQU", COMAQ$, "NOMESS")
            End If
            DEMAQ$ = TRIM$(Mid$(X$, 7, 30))
            Call ACTUREGISTRO("EQUIPOS", "DESCRI_EQU", "CODIGO_EQU = '" & COMAQ$ & "'", DEMAQ$, RAFE&)
            '
        End If
    Next U&
    '
    Call SETULTREG("PADMAQ", J&)
    Call CIERRARCH("PADMAQ")
    Call CIERRARCH("!PADMAQ")
    '
End Sub



