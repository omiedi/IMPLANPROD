Attribute VB_Name = "MAST_FUN"
Public PUNCODIG As String
Public BUFERCOD As String * 2048
Public NUSU%, COSU%(64), SUCU$(64), ADEPO$(64), DISTO%(64), CATEO#(64), CAFIN#(64)
Dim CIANT%, XXX$
Public PHABIT%(32767)

Public UPDATMASTER%, UPDAPLANOS%
Dim MasterTemp As Recordset
'
Sub AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
   '
   Screen.MousePointer = 11
   '
   'validar que no exista el codigo
   'validar que el codigo sea >3 y <15
   'validar que la descripcion no sea nula
   'dar el update
   'hacer un gencomas.
   '
   'validaciones
   If Len(CODINUE$) < 3 Or Len(CODINUE$) > 15 Then
        Call MENSERR(24, "El Codigo de Articulo no debe ser Inferior a tres Cararcteres. No se ha podido dar del Alta el Articulo.")
        Screen.MousePointer = 1
        GoTo 99
   End If
   '
   If Len(DESCRINUE$) <= 0 Then
        Call MENSERR(24, "El Articulo no tiene Descripcion. No se ha Podido dar de Alta el Articulo.")
        Screen.MousePointer = 1
        GoTo 99
   End If
   '
   CODII% = NUMAS% + 1
20 On Error Resume Next
   Call APERBASDAT("MASTER")
    Master.MoveFirst
   If Err < 1 Then
     On Error GoTo 0
     ABUS$ = "Codint = " & CODII%
     Master.FindFirst "Codint = " & CODII%
     If Master.NoMatch = False Then
       CODII% = CODII% + 1
       GoTo 20
     End If
   End If
   '
   SQLX$ = "SELECT * FROM MASTER "
   SQLX$ = SQLX$ + " WHERE TRIM(CODIGO) = '" & CODINUE$ & "' "
   Set COCODTMP = Articulos.OpenRecordset(SQLX$, 4)
   With COCODTMP
     On Error Resume Next
     .MoveFirst
     If Err Then 'el codigo no es pre-existente LO aGREGO
       On Error GoTo 0
       With Master
         .AddNew
           !CODIGO = CODINUE$
           !DESCRIPCION = DESCRINUE$
           !Status = 1
           !CODINT = CODII%
           !UMEDIDA = UNID$
           !TIMATE = TIMA%
         .Update
       End With
       '
       ' generar un registro para los saldos de stock
       Call ABRESTOCKS
       With SalStock
         .FindFirst "CodiInte = " & CODII%
         If .NoMatch = True Then
           .AddNew
           !CodiInte = CODII%
           !VeSalSto = String$(256, 0)
           .Update
         End If
       End With
       '
       UPDATMASTER% = 1
       Call GENECOMAS
       OKX% = 1
       Exit Sub
       '
     End If
     On Error GoTo 0
     Call MENSERR(24, "Código de Articulo no válido o Pre-Existente. No se ha Podido dar de Alta el Articulo.")
     GoTo 99
   End With
   '
99 Screen.MousePointer = 1
End Sub

Sub GENECOMAS()
  '
  If UPDATMASTER% > 0 Then
    '
    Screen.MousePointer = 11
    '
    Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT,NPLANO FROM Master WHERE STATUS > 0 ORDER BY CODIGO ASC;", dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    MASTERECO.MoveLast
    If Err Then GoTo 15
    '
    CAREC% = MASTERECO.RecordCount
    If CAREC% < 1 Then GoTo 15
    '
    FIN& = CAREC%
    REBLA$ = REGBLAN$("ECOMAST")
    JJ& = 0: JJ1& = 0
    MASTERECO.MoveFirst
    '
10  Call TRACE(20, JJ1&, FIN&)
    EL1$ = REBLA$
    JJ1& = JJ1& + 1
    CODEXX$ = MASTERECO!CODIGO
    '
    If Mid$(CODEXX$, 16, 1) = Chr$(96) Then
      CIXXI% = MASTERECO!CODINT
      If CIXXI% > 0 Then
        Master.FindFirst ("CODINT = " & CIXXI%)
        If Master.NoMatch = False Then
          Master.Edit
          Master!Status = (-1)
          Master!CODIGO = TRIM$(Left$(CODEXX$, 15))
          Master.Update
          GoTo 12
        End If
      End If
    End If
    '
    If TRIM$(CODEXX$) <> "" Then
      Call REPLA(EL1$, CODEXX$, 1, 16)
        DEPLANO$ = TRIM$(MASTERECO!DESCRIPCION)
        If IsNull(MASTERECO!NPLANO) = False Then
          NXPLANO$ = TRIM$(MASTERECO!NPLANO)
          If NXPLANO$ <> "" Then
            If NXPLANO$ <> TRIM$(MASTERECO!CODIGO) Then
              While Len(DEPLANO$ + NXPLANO$) > 43
                DEPLANO$ = Left$(DEPLANO$, Len(DEPLANO$) - 1)
              Wend
              DEPLANO$ = DEPLANO$ + " (" + NXPLANO$ + ")"
            End If
          End If
        End If
      Call REPLA(EL1$, DEPLANO$, 17, 46)
      Call REPLA(EL1$, MKI$(MASTERECO!CODINT), 63, 2)
      JJ& = JJ& + 1
      Call GRAREG("ECOMAST", EL1$, JJ&)
    End If
    '
12  If JJ1& < CAREC% Then
      MASTERECO.MoveNext
      GoTo 10
    End If
    '
    MASTERECO.Close
    Call SETULTREG("ECOMAST", JJ&)
    Call CIERRARCH("ECOMAST")
    Call FRESHECHO("ECOMAST")
    Call BAJALDISCO
  End If
  '
15 If UPDATMASTER% > 0 Or UPDAPLANOS% > 0 Then
    Screen.MousePointer = 11
    Set MASTERECO = Articulos.OpenRecordset("SELECT NPlano,CODIGO,DESCRIPCION,CODINT FROM Master WHERE STATUS > 0 AND NPLANO <> '' ORDER BY NPLANO ASC;", dbOpenDynaset, dbReadOnly)
    On Error Resume Next
    MASTERECO.MoveLast
    If Err Then GoTo 99
    '
    CAREC% = MASTERECO.RecordCount
    If CAREC% < 1 Then GoTo 99
    '
    FIN& = CAREC%
    REBLA$ = REGBLAN$("IPLANOS")
    JJ& = 0
    MASTERECO.MoveFirst
    '
20  Call TRACE(20, JJ&, FIN&)
    EL1$ = REBLA$
    Call REPLA(EL1$, MASTERECO!NPLANO, 1, 24)
    Call REPLA(EL1$, MASTERECO!DESCRIPCION, 25, 54)
    Call REPLA(EL1$, MKI$(MASTERECO!CODINT), 79, 2)
    JJ& = JJ& + 1
    Call GRAREG("IPLANOS", EL1$, JJ&)
    '
    If JJ& < CAREC% Then
      MASTERECO.MoveNext
      GoTo 20
    End If
    '
    MASTERECO.Close
    Call SETULTREG("IPLANOS", JJ&)
    Call CIERRARCH("IPLANOS")
    Call FRESHECHO("IPLANOS")
    Call BAJALDISCO
   End If
   '
99 UPDATMASTER% = 0
   UPDAPLANOS% = 0
   Screen.MousePointer = 1
End Sub
'
Function NUMAS%()
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If CONTROL$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Call APERBASDAT("MASTER")
   On Error Resume Next
   Master.MoveLast
   NUMAS% = Master.RecordCount
   If NUMAS% > 0 Then
     If Master!CODINT > NUMAS% Then
       Call MENSERR(24, "La Base de Datos Maestro de Artículos\está Dañada !!! - Consulte Urgente !!!")
       Call FINAL("")
     End If
   End If
   '
End Function

Function PRICOD%()
   '
   PRICOD% = 0
   Call APERBASDAT("MASTER")
   Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
   On Error Resume Next
   With MasterTp
     .MoveFirst
     If Err < 1 Then
       PRICOD% = CODINT%(!CODIGO)
     End If
   End With
   On Error GoTo 0
   '
End Function

Function ULTCOD%()
   '
   ULTCOD% = 0
   Call APERBASDAT("MASTER")
   Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
   On Error Resume Next
   With MasterTp
     .MoveLast
     If Err < 1 Then
       ULTCOD% = CODINT%(!CODIGO)
     End If
   End With
   On Error GoTo 0
   '
End Function

Function STATUIT%(CI%)
   Static CIANT%, STATUIA%, FIDAMASANT#
   '
   ' = 1 --> Existe
   ' = 0 --> No existe el item
   '
   If STATUIA% < 1 Then
     Call APERBASDAT("MASTER")
   End If
   '
   If CI% = CIANT% Then
     If STATUIA% > 0 Then
       STATUIT% = STATUIA%
       Exit Function
     End If
   End If
   '
   STATUIT% = 0
   If CI% > 0 Then
     'CALL APERBASDAT("MASTER")
     FIDAMAS# = FIDATIM#("ECOMAST")
     If FIDAMAS# <> FIDAMASANT# Then
       FIDAMASANT# = FIDAMAS#
       On Error Resume Next
       Master.Close
       On Error GoTo 0
       Call APERBASDAT("MASTER")
     End If
     '
     Set MasterTemp = Articulos.OpenRecordset("SELECT * FROM Master WHERE MASTER!CODINT=" & CI%, dbOpenDynaset, dbReadOnly)
     On Error Resume Next '
     MasterTemp.MoveFirst
     MasterTemp.FindFirst "CODINT =" & CI%
     If MasterTemp.NoMatch = False Then
       STATUIT% = 1
     End If
     On Error GoTo 0
   End If
   '
   CIANT% = CI%
   STATUIA% = STATUIT%
   '
End Function
 
Function STATITEM%(CI%)
'\\\OT-06-0353-MH-29/08/06///
   ' DEVUELVE  1 SI ESTÁ ACTIVO
   '          -1 SI ESTA DE BAJA
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then STATITEM% = MasterTemp!Status
'\\\OT-06-0353-MH-FIN///
End Function
Function CODINT%(CODEXTE$)
   '
   Static FIDAMASANT#
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If CONTROL$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   If CODEXTE$ <> "" Then
     '
     FIDAMAS# = FIDATIM#("ECOMAST")
     If FIDAMAS# <> FIDAMASANT# Then
       FIDAMASANT# = FIDAMAS#
       On Error Resume Next
       Master.Close
       On Error GoTo 0
       Call APERBASDAT("MASTER")
     End If
     '
     Call APERBASDAT("MASTER")
     On Error Resume Next
     Master.MoveFirst
     CDXX$ = TRIM$(CODEXTE$): CDYY$ = AJUSTI$(CODEXTE$, 15) + Chr$(96)
     SQLX$ = "CODIGO ='" + CDXX$ + "' OR CODIGO = '" & CDYY$ & "' "
       If INCLUALIAS% = 1 Then SQLX$ = SQLX$ + "OR NPlano = '" + CDXX$ + "' "
     Master.FindFirst SQLX$
     If Master.NoMatch = False Then
       CODINT% = Master!CODINT * Master!Status
     End If
     On Error GoTo 0
   End If
   '
End Function

Function CODINPLA%(NUPLAX$)
   '
   If NUPLAX$ <> "" Then
     Call APERBASDAT("MASTER")
     On Error Resume Next
     Master.MoveFirst
     Master.FindFirst "NPLANO ='" + TRIM$(NPLAX$) + "'"
     If Master.NoMatch = False Then
       CODINPLA% = Master!CODINT
     End If
     On Error GoTo 0
   End If
   '
End Function

Function CODEXT$(CI%)
   '
   Static CDXT$()
   Static CTX1 As Integer
   '
   If CTX1% < 1 Then
     ReDim CDXT$(NUMAS%)
     CTX1% = 1
   End If
   '
   CIK1% = Abs(CI%)
   If CIK1% > UBound(CDXT$) Then
     ReDim Preserve CDXT$(CIK1%)
   End If
   '
   If UPDATMASTER% > 0 Then
      For KKU& = 1 To UBound(CDXT$)
        CDXT$(KKU&) = ""
      Next KKU&
   End If
   '
   If CDXT$(CIK1%) <> "" Then
     CODEXT$ = CDXT$(CIK1%)
     Exit Function
   End If
   '
   On Error Resume Next
   If STATUIT%(CIK1%) <> 0 Then CODEXT$ = TRIM$(MasterTemp!CODIGO)
   CDXT$(CIK1%) = CODEXT$
   '
End Function
 
Function DESCRIT$(CI%)
   On Error Resume Next
   DESCRIT$ = ""
   If CI% < 1 Then
       DESCRIT$ = "Material No Codificado"
     Else
       DESCRIT$ = DESCRIT0$(CI%)
       If TRIM$(DESCRIT$) <> "" Then DESCRIT$ = DESCRIT$ + " (" + UNIMED$(CI%) + ")"
   End If
End Function

Function DESCRIT0$(CI%)
   On Error Resume Next
   DESCRIT0$ = ""
   If CI% < 1 Then
       DESCRIT0$ = "Material No Codificado"
     Else
       If STATUIT%(CI%) <> 0 Then DESCRIT0$ = TRIM$(MasterTemp!DESCRIPCION)
   End If
End Function

Function DESCRIT2$(CI%)
   On Error Resume Next
   DESCRIT2$ = ""
   If STATUIT%(CI%) <> 0 Then DESCRIT2$ = TRIM$(MasterTemp!DESCRIDOS)
End Function

Function UNIMED$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then UNIMED$ = TRIM$(MasterTemp!UMEDIDA)
End Function

Function TIMATE%(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then TIMATE% = CDbl(MasterTemp!TIMATE)
End Function

Function CRITREP%(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then CRITREP% = CDbl(MasterTemp!CRITREP)
End Function

Function GRUCOVEI%(CI%)
   On Error Resume Next
   If STATUIT%(CI%) > 0 Then GRUCOVEI% = CDbl(MasterTemp!Grucoven)
End Function

Function GRUPIVA%(CI%)
   Static GRUMAXI%
   If GRUMAXI% = 0 Then
     GRUMAXI% = ULTREG&("GRUPIVA")
   End If
   On Error Resume Next
   If STATUIT%(CI%) > 0 Then GRUPIVA% = CDbl(MasterTemp!GRUPIVA)
   If GRUPIVA% > GRUMAXI% Then GRUPIVA% = 0
End Function

Function LOTESTAN(CI%)
   On Error Resume Next
   If STATUIT%(CI%) > 0 Then LOTESTAN = CDbl(MasterTemp!LOREPOS)
End Function
  
Function COSUNI(CI%)
   On Error Resume Next
   If STATUIT%(CI%) > 0 Then COSUNI = CDbl(MasterTemp!COSUNI)
End Function
'
Function MONECOSTO%(CI%)
   MONECOSTO% = 1
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then MONECOSTO% = CDbl(MasterTemp!MONECOS)
End Function

Function FECOSTO%(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then FECOST$ = Format(MasterTemp!FECOSTO, "dd/mm/yy")
   FECOSTO% = FECHANUM(FECOST$)
End Function
'
Function COSPESOS(CI%)
    COSPESOS = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%)) * (1 + PORNACIO(CI%) / 100)
End Function
'
Function PORNACIO(CI%)
   PORNACIO = 0
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PORNACIO = CDbl(MasterTemp!PORNACIO)
End Function
'
Function COEFNACIO(CI%)
   COEFNACIO = 1 + PORNACIO(CI%) / 100
End Function
'
Function PORCOMVE(CI%)
   PORCOMVE = 0
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PORCOMVE = CDbl(MasterTemp!PorComVen)
End Function
'
Function PROVHABI%(CI%)
   PROVHABI% = 0        ' FABRICACION PROPIA
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PROVHABI% = CDbl(MasterTemp!PROVHABI)
   If PROVHABI% > 0 Then
     If ECOARCH$("ACREDOR", MKI$(PROVHABI%)) = "" Then
       PROVHABI% = 0
     End If
   End If
End Function
 
Function COLOR$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then COLOR$ = TRIM$(MasterTemp!COLOR)
End Function
'
Function TALLE$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then TALLE$ = TRIM$(MasterTemp!TALLE)
End Function
'
Function PESUNI(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PESUNI = CDbl(MasterTemp!PESUNI)
End Function
'
Function PESMETRO(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PESMETRO = CDbl(MasterTemp!PESMETRO)
End Function
'
Function LTITPE$(CP)
   'convierte peso unitario en titulo de hilado
   LCLR$ = ""
   If CP > 0 Then
        LCLR$ = CSTRING$(MKS$(CP), 3, 10, 5, 2)
      ElseIf CP < 0 Then
        LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
        LCLR$ = REPLACAR$(LCLR$, ".", "/")
   End If
   LTITPE$ = LCLR$
   '
End Function

Function OBSERVAIT$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then OBSERVAIT$ = MasterTemp!Observa
End Function
'
Function ESPECIFIT$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then ESPECIFIT$ = MasterTemp!ESPECIFICACION
End Function
'
Function NUPLANO$(CI%)
   '
   NUPLANO$ = ""
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then NUPLANO$ = TRIM$(MasterTemp!NPLANO)
   '
End Function
'\\\OT-05-0702-WAR-01/11/05///
Function ATRIB1$(CI%) 'ATRIBUTO UNO
   '
   ATRIB1$ = ""
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then ATRIB1$ = TRIM$(MasterTemp!TALLE)
   '
End Function
Function ATRIB2$(CI%) 'ATRIBUTO DOS
   '
   ATRIB2$ = ""
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then ATRIB2$ = TRIM$(MasterTemp!COLOR)
   '
End Function
'
Function UNICOM$(CI%, Optional NPROVE%)
   '
   UNICO$ = ""
   '
   If CI% > 0 Then  ' que no lo haga con material libre
     SQLX$ = "SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
       If NPROVE% > 0 Then
         SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROVE% & " "
       End If
     SQLX$ = SQLX$ + "ORDER BY Femi_Ocom ASC, Nume_Ocom ASC"
     Set OcomdetaTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     On Error Resume Next
     With OcomdetaTemp
       .MoveLast
       If Err < 1 Then
         On Error Resume Next
         UNICO$ = !Uni_Com
         On Error GoTo 0
       End If
     End With
     '
     If ECOARCH$("UNIMED", UNICO$) = "" Then UNICO$ = ""
     If UNICO$ = "" Then UNICO$ = UNIMED$(CI%)
   End If
   '
   UNICOM$ = UNICO$
   '
End Function

Function CANOCPEN(CI%) ' DEVUELVE LAS CANTIDADES COMPRADAS PENDIENTES DE RECEPCIÓN
    ' '\\\OT-07-0386-FL-16/10/07///
    CANPEN# = 0
    If CI% > 0 Then
        Call APERBASDAT("COMPRAS")
        SQLX$ = "SELECT *  FROM OCOMDETA "
        SQLX$ = SQLX$ + "WHERE CANT_OCOM > CANT_REC "
        SQLX$ = SQLX$ + "AND COD_INTE = " & CI% & " "
        SQLX$ = SQLX$ + "AND STAT_OCOM< 3 "
        Set OcomdetaTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With OcomdetaTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
               CANPEN# = CANPEN# + (!Cant_Ocom - !Cant_Rec)
             .MoveNext
             Loop
           End If
        End With
    End If
    CANOCPEN = CANPEN#
   ' '\\\OT-07-0386-FIN///
End Function
'
Function CANFAPEN(CI%) 'DEVUELVE LA CANTIDAD TOTAL EN PROCESO DE FABRICACIÓN DE UN ARTÍCULO
'\\\OT-07-0386-FL-17/10/07///
    CANFA# = 0
    If CI% > 0 Then
        Call APERBASDAT("ORFAB")
        SQLX$ = "SELECT * FROM ORDEFABR "
        SQLX$ = SQLX$ + "WHERE CANAPRO < CANPROY "
        SQLX$ = SQLX$ + "AND COD_INTE = " & CI% & " "
        SQLX$ = SQLX$ + "AND StatuOrf < 3 "
        Set OrdefabrTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With OrdefabrTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             Do While Not .EOF
               CANFA# = CANFA# + (!CanProy - !CanApro)
             .MoveNext
             Loop
           End If
        End With
    End If
    CANFAPEN = CANFA#
    '\\\OT-07-0386-FIN///
End Function

Function CANRECEP(CI%)
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + "AND Cod_Inte = " & CI% & " "
    '
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    TORECEPE = 0
    With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         K& = 0
         Do While Not .EOF
            CARECEP = !CANTNOM - !CanToApro
            If CARECEP > 0 Then
               TORECEPE = TORECEPE + CARECEP
            End If
         .MoveNext
         Loop
       End If
    End With
    CANRECEP = TORECEPE
End Function
'
Function REVICOD$(CI%)
   REVICOD$ = ""
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then REVICOD$ = TRIM$(MasterTemp!Revision)
End Function
'
Function FEREVI%(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then FEREVIST$ = Format(MasterTemp!FEREVIS, "dd/mm/yy")
   FEREVI% = FECHANUM(FEREVIST$)
End Function
'
' ***************************************************************************
'
Function FEREFOR%(CI%)   ' FECHA REVISION FORMULA
  '
  FRF% = 0
  If CI% > 0 Then
    XX$ = FILTERGETRECORD$("REVIFORM", 1, MKI$(CI%))
    FRF% = CVI(Mid$(XX$, 3, 2))
  End If
  FEREFOR% = FRF%
  '
End Function
'
' ***************************************************************************
'
Function USREFOR$(CI%)    ' USUARIO RESPONSABLE REVISION
  '
  URF$ = ""
  If CI% > 0 Then
    XX$ = FILTERGETRECORD$("REVIFORM", 1, MKI$(CI%))
    URF$ = TRIM$(Mid$(XX$, 5, 28))
  End If
  USREFOR$ = URF$
  '
End Function
'
' ***************************************************************************
'
Function CCOLOR%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TACOLOR")
           XX$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       If ULTREG&("TACOLOR") >= 255 Then
            Call MENSERR(24, "Saturación de Tabla de Colores.\Imposible Agregar Color '" + TC$ + "' a Tabla.")
            CC% = 0
          Else
            CC% = 1 + ULTREG&("TACOLOR")
            Call REGAPP("TACOLOR", TC$)
            Call CIERRARCH("TACOLOR")
            Call FRESHECHO("TACOLOR")
       End If
    End If
    '
99  CCOLOR% = CC%
    '
End Function
'

Function LCOLOR$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TACOLOR") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
      End If
   End If
   '
99 LCOLOR$ = LCLR$
   '
End Function
'
' ***************************************************************************
'

Function CTALLE%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TATALLES")
           XX$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       If ULTREG&("TATALLES") >= 255 Then
            Call MENSERR(24, "Saturación de Tabla de Talles.\Imposible Agregar Talle '" + TC$ + "' a Tabla.")
            CC% = 0
          Else
            CC% = 1 + ULTREG&("TATALLES")
            Call REGAPP("TATALLES", TC$)
            Call CIERRARCH("TATALLES")
            Call FRESHECHO("TATALLES")
       End If
    End If
    '
99  CTALLE% = CC%
    '
End Function
'
Function LTALLE$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TATALLES") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
      End If
   End If
   '
99 LTALLE$ = LCLR$
   '
End Function
'
' ***************************************************************************
'
Function TIPOHIL$(CI%)
'
TIPOHIL$ = TALLE$(CI%)
End Function
'
' ***************************************************************************
'

Function TITULO$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
TITULO$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CP = CVS(Mid$(XXX$, 99, 4))
       If CP < 0 Then
          LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
          TITULO$ = REPLACAR$(LCLR$, ".", "/")
       End If
   End If
End If
End Function
'
' ***************************************************************************
'

' ***************************************************************************
'

Function COSCALCU(CI%, MODOCO%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 89, 8))
    End If
  End If
  '
  MODOCO% = 1
  If COU < 0.0005 Then
    COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%)) * (1 + PORNACIO(CI%) / 100)
    MODOCO% = 2
  End If
  '
  COSCALCU = COU
  '
End Function
' ***************************************************************************
'
Function PRECALCU(CI%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 97, 8))
    End If
  End If
  '
  PRECALCU = COU
  '
End Function
' ***************************************************************************
'
Function COSMATCAL(CI%, MONECO%)
   '
   ' determina el costo calculado de suministros (M.Prima,
   ' Piezas y Trabajos de Terceros en la misma forma que
   ' el programa ANACOS01. Si MONECO%=1 en pesos
   '                       Si MONECO%=2 en dolares
   '
   Static DP%(32767)
   Static FESTRU#, NVI%, NMM%
   '
   SUCOS = 0
   CICON% = CI%
   If CICON% < 1 Or CICON% > NUMAS% Then GoTo 90
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
     '
     NVI% = NUINV%
     NMM% = NUMAS%
     '
     FESTRU# = FESTRUC#
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = REGLEIDO$("ESTRUNUE", I&)
       CI1% = CVI(Left$(R0$, 2))
       If CI1% > 0 Then
         If CI1% <= NMM% Then
           If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
             DP%(CI1%) = 1
           End If
         End If
       End If
     Next I&
   End If
   '
   Call EXPLOMUL(CICON%)
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
     COSUDOL = COSUNI(CIK%(KKI%))
     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
        '
        MONEI% = MONECOSTO%(CIK%(KKI%))
        COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
        SUCOS = SUCOS + CANT(KKI%) * COSUN
        '
      Else
        '
        If COSUDOL > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
            If PROVINT%(PHB%) = 0 Then
              MONEI% = MONECOSTO%(CIK%(KKI%))
              COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
              SUCOS = SUCOS + CANT(KKI%) * COSUN
            End If
          End If
        End If
        '
     End If
   Next KKI%
   '
   If MONECO% = 2 Then
      TCM = TICAMONE(2)
      If TCM < 1 Then TCM = 1
      SUCOS = SUCOS / TCM
   End If
   '
90 COSMATCAL = SUCOS
   '
End Function

'
' ***************************************************************************
'

' ***************************************************************************
'
Function PROVINT%(NPRX%)
   '
   Static PVIT%(32767), CPVI%
   '
   If CPVI% < 1 Then
     CPVI% = 1
     For UI& = 0 To 32767
       PVIT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NPPX% = CVI(Left$(X$, 2))
       If NPPX% > 0 Then
         PVIT%(NPPX%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
   End If
   '
   PROVINT% = 0
   If NPRX% > 0 Then
     PROVINT% = PVIT%(NPRX%)
   End If
   '
End Function
'
' ***************************************************************************
'
Function PREBASE(CI%)
'
PREBASE = COSUNI(CI%)
End Function
'
' ***************************************************************************
'

' ***************************************************************************
'
Function PORCEIVA(CILOCAL%)
   '
   Static CTX%, URE&, PRIVA(15), REMA As String * 128
   If CTX% < 1 Then
     CTX% = 1
     URE& = ULTREG&("GRUPIVA")
     'If URE& > 4 Then URE& = 4
     For U& = 1 To URE&
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
     Next U&
     Call CIERRARCH("GRUPIVA")
   End If
   '
   GRIVA% = GRUPIVA%(CILOCAL%)
   If GRIVA% < 0 Or GRIVA% > URE& - 1 Then GRIVA% = 0
   '
   PORCEIVA = PRIVA(GRIVA%)
   '
End Function
'
' ***************************************************************************
'
Function STOMIN(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then STOMIN = CDbl(MasterTemp!STOMIN)
End Function
'
' ***************************************************************************
'
' ***************************************************************************
'\\\OT-05-0720-WAR-25/11/05///
Function LINDIVIS(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then LINDIVIS = CDbl(MasterTemp!LINDIVIS)
End Function
'\\\OT-05-0720-WAR-FIN///
'
' ***************************************************************************
'
Function LOREPOMA(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then LOREPOMA = CDbl(MasterTemp!STOMAX)
End Function
'
' ***************************************************************************
'
Sub MUESTRADATCOD(CODDX$)
   '
   CICDX% = CODINT%(CODDX$)
   If CICDX% > 0 Then
     If CICDX% <= NUMAS% Then
       REMASX& = CICDX% + 1
       RMCDX$ = REGBLAN$("MASTER") + String$(52, 0)
       Call REPLA(RMCDX$, MKI$(CICDX%), 1, 2)
       Call REPLA(RMCDX$, DESCRIT0$(CICDX%), 5, 64)
       Call REPLA(RMCDX$, CODEXT$(CICDX%), 77, 16)
       Call REPLA(RMCDX$, Chr$(TIMATE%(CICDX%)), 93, 1)
       Call REPLA(RMCDX$, Chr$(CRITREP%(CICDX%)), 94, 1)
       Call REPLA(RMCDX$, UNIMED$(CICDX%), 109, 2)
       Call REPLA(RMCDX$, MKS$(COSUNI(CICDX%)), 111, 4)
       Call REPLA(RMCDX$, MKI$(FECOSTO%(CICDX%)), 115, 2)
       Call REPLA(RMCDX$, MKI$(PROVHABI%(CICDX%)), 123, 2)
       Call REPLA(RMCDX$, MKS$(STOCKACT(CICDX%)), 129, 4)
       Call REPLA(RMCDX$, MKS$(STOMIN(CICDX%)), 133, 4)
       Call REPLA(RMCDX$, MKS$(LOTESTAN(CICDX%)), 137, 4)
       Call REPLA(RMCDX$, MKS$(LOREPOMA(CICDX%)), 141, 4)
       Call REPLA(RMCDX$, TALLE$(CICDX%), 145, 16)
       Call REPLA(RMCDX$, COLOR$(CICDX%), 161, 16)
       Call REPLA(RMCDX$, MKS$(PESUNI(CICDX%)), 177, 4)
       If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
         Call REPLA(RMCDX$, MKS$(99999999999#), 111, 4)
       End If
       OBXXX$ = OBJPLA$("MUESTRADATCOD", RMCDX$)
     End If
   End If
End Sub

Sub MUESTRADATPRO(NPROV$)
   NPY& = XVALO(NPROV$)
   If NPY& > 0 Then
     If NPY& <= 32767 Then
       NPX% = NPY&
       If ECOARCH$("PROVED1", MKI$(NPX%)) <> "" Then
         RMCDX$ = String$(1024, 0)
         NC1% = (-1) * NPX%
         Call REPLA(RMCDX$, MKI$(NPX%), 1, 2)
         Call REPLA(RMCDX$, RASOCLI$(NC1%), 3, 62)
         Call REPLA(RMCDX$, TEL2CLI$(NC1%), 65, 32)
         Call REPLA(RMCDX$, DOMICLI$(NC1%), 97, 64)
         Call REPLA(RMCDX$, CPOSCLI$(NC1%), 161, 16)
         Call REPLA(RMCDX$, LOCACLI$(NC1%), 177, 48)
         Call REPLA(RMCDX$, TEL1CLI$(NC1%), 257, 24)
         Call REPLA(RMCDX$, PAISCLI$(NC1%), 225, 32)
         Call REPLA(RMCDX$, PROVCLI$(NC1%), 529, 1)
         Call REPLA(RMCDX$, FAXCLI$(NC1%), 281, 24)
         Call REPLA(RMCDX$, EMAILCLI$(NC1%), 305, 48)
         Call REPLA(RMCDX$, Chr$(PIVACLI%(NC1%)), 530, 1)
         Call REPLA(RMCDX$, CUITCLI$(NC1%), 513, 16)
         Call REPLA(RMCDX$, Chr$(TICUEPRO%(NC1%)), 531, 1)
         Call REPLA(RMCDX$, Chr$(CPAGCLI%(NC1%)), 532, 1)
         OBXXX$ = OBJPLA$("MUESTRADATPRO", RMCDX$)
       End If
     End If
   End If
End Sub

Sub REPUNCODIG()
   '
Exit Sub
   Screen.MousePointer = 11
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BFC$ = "": JI% = 0
   For U& = 1 To NUMAS%
      X$ = REGLEIDO$("MASTER", U& + 1)
      BFC$ = BFC$ + Mid$(X$, 77, 16)
      If Len(BFC$) >= 2048 Then
         BUFERCOD$ = BFC$
         JI% = JI% + 1
         Put #NX%, JI%, BUFERCOD$
         BFC$ = ""
      End If
   Next U&
   '
   If Len(BFC$) > 0 Then
      While Len(BFC$) < 2048
         BFC$ = BFC$ + Chr$(0)
      Wend
      BUFERCOD$ = BFC$
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
   End If
   '
   For JI1% = JI% + 1 To LOF(NX%) / 2048
      BUFERCOD$ = String$(2048, 0)
      Put #NX%, JI1%, BUFERCOD$
   Next JI1%
   '
   Close #NX%: NX% = 0
   PUNCODIG$ = ""
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARDEPOS()
   '
   If NUSU% < 1 Then
      SMP = Screen.MousePointer
      Screen.MousePointer = 11
      FIN& = ULTREG&("SUCURSAL")
      For U& = 1 To FIN&
         If U& <= 62 Then
            X$ = REGLEIDO$("SUCURSAL", U&)
            NUSU% = U&
            If Len(X$) >= 64 Then
               COSU%(U&) = Asc(Mid$(X$, 64, 1))
               SUCU$(U&) = Left$(X$, 20)
               ADEPO$(U&) = Mid$(X$, 21, 8)
               If TRIM$(ADEPO$(U&)) = "" Then ADEPO$(U&) = Left$(X$, 8)
               CATEO#(U&) = 0
               CAFIN#(U&) = 0
               DISTO%(U&) = 0
               If UCase$(Mid$(X$, 62, 2) = "NO") Then DISTO%(U&) = (-1)
            End If
         End If
      Next U&
      '
      If NUSU% < 1 Then
         NUSU% = 1
         COSU%(1) = 1
         SUCU$(1) = "Deposito Central"
         ADEPO$(1) = "Central"
         CATEO#(1) = 0
         CAFIN#(1) = 0
      End If
      Screen.MousePointer = SMP
   End If
   '
End Sub
'
Function STOCKACT(CI%)
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   Set SalStTemp = Stocks.OpenRecordset("SELECT * FROM SALSTOCK WHERE CodiInte = " & CI%, dbOpenSnapshot)
   With SalStTemp
     On Error Resume Next
     .MoveFirst
     If Err Then
         VESAL$ = String$(256, 0)
       Else
         VESAL$ = !VeSalSto
     End If
   End With
   Set SalStTemp = Nothing
   '
   STAC = 0
   For KKI% = 1 To NUSU%
     If DISTO(KKI%) >= 0 Then
       STAC = STAC + CVS(Mid$(VESAL$, 4 * KKI% - 3, 4))
     End If
   Next KKI%
   '
   STOCKACT = STAC
   '
End Function
'
Function STODEPOS(CI%, DEPO%)
   '
   If DEPO% < 1 Then
      STAC = STOCKACT(CI%)
      GoTo 99
   End If
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   Set SalStTemp = Stocks.OpenRecordset("SELECT * FROM SALSTOCK WHERE CodiInte = " & CI%, dbOpenSnapshot)
   With SalStTemp
     On Error Resume Next
     .MoveFirst
     If Err Then
         VESAL$ = String$(256, 0)
       Else
         VESAL$ = !VeSalSto
     End If
   End With
   Set SalStTemp = Nothing
   '
   STAC = 0
   For KKI% = 1 To NUSU%
     If COSU%(KKI%) = DEPO% Then
       STAC = CVS(Mid$(VESAL$, 4 * KKI% - 3, 4))
       Exit For
     End If
   Next KKI%
   '
99 STODEPOS = STAC
   '
End Function
'
Function STODEPFECH(CI%, DEPO%, FECH%)
   '
   ' DEVUELVE LA EXISTENCIA EN EL DEPÓSITO DEPO% A LA FECHA FECH% - 04/04/07 (OT-07-0141)
   ' SI DEPO% = 0 DEVUELVE EXISTENCIA TOTAL DEL ITEM A ESA FECHA
   If CI% > 0 Then
      SCMA = Screen.MousePointer
      Screen.MousePointer = 11
      '
      CARRAS# = 0
      HASDAT = CDbl(CVDAT(FECH%))
      '
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND FechMov <= " & HASDAT
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      Set MOVITEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With MOVITEMP
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             CARRAS# = CARRAS# + !CantIngre - !CANTSALID
           .MoveNext
           Loop
         End If
      End With
      '
   End If
   '
   STODEPFECH = CARRAS#
   Screen.MousePointer = SCMA
   '
End Function
'
Function STOCKASI(CI%)
    '
10  ASII = 0
    UREOD& = ULTREG&("MAVENDID")
    RGXX$ = RECFILT$("MAVENDID", 20, MKI$(CI%))
    '
    For U& = 1 To Len(RGXX$) Step 4
      REG& = CVS(Mid$(RGXX$, U&, 4))
      If REG& > 0 Then
        If REG& <= UREOD& Then
          X$ = REGLEIDO$("MAVENDID", REG&)
          If Asc(Mid$(X$, 63, 1)) <= 1 Then  ' STATUS 1
            CASIG = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
            If CASIG < 0 Then CASIG = 0
            ASII = ASII + CASIG
          End If
        End If
      End If
    Next U&
    '
    STOCKASI = ASII
    '
End Function
'
Function STOCKRES(CI%)
    '
10  ASII = 0
    '
35  Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    '
    Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ReserTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          TORE# = !CantRes
          CACO# = !CantCons
          '
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
             TORE# = TORE# / 1000
             CACO# = CACO# / 1000
          End If
          '
          SARE = TORE# - CACO#
          If SARE < 0 Then SARE = 0
          ASII = ASII + SARE
        .MoveNext
        Loop
      End If
    End With
    '
    STOCKRES = ASII
    '
End Function
'
Function STORESOF(CI%, NORFAI&)
    '
    NBRS$ = TRIM$(Str$(NORFAI&))
    While Len(NBRS$) < 6
      NBRS$ = "0" + NBRS$
    Wend
    NBRS$ = "OF-" + NBRS$ + "-*"
    '
10  ASII = 0
    '
35  Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "AND IdSubOr LIKE '" & NBRS$ & "' "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    '
    Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ReserTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          TORE# = !CantRes
          CACO# = !CantCons
          '
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
             TORE# = TORE# / 1000
             CACO# = CACO# / 1000
          End If
          '
          SARE = TORE# - CACO#
          If SARE < 0 Then SARE = 0
          ASII = ASII + SARE
        .MoveNext
        Loop
      End If
    End With
    '
    STORESOF = ASII
    '
End Function
'
Function STOCKDIS(CI%)
    '
    STOCKDIS = STOCKACT(CI%) - STOCKASI(CI%) - STOCKRES(CI%)
    '
End Function
'

Function STOCKDISNEVEN(CI%)
    '
    STOCKDISNEVEN = STODEPOS(CI%, DEPOEXPE%) - STOCKASI(CI%) - STOCKRES(CI%) - PEDNOASI(CI%)
    '
End Function
'

Function STOCKDISNE(CI%)
    '
    STOCKDISNE = STOCKDIS(CI%) - PEDNOASI(CI%)
    '
End Function
'
'
Function STORFAB(CI%)
   '
   STORF = 0
   RFF$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
   For UK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK&, 4))
      X$ = REGLEIDO$("ORDEFABR", U&)
      STAT% = Asc(Mid$(X$, 108, 1))
      If STAT% <= 2 Then
        CACOM = CVD(Mid$(X$, 89, 8))
        CAREC = CVD(Mid$(X$, 119, 8))
        CAPEN = CACOM - CAREC
        If CAPEN > 0 Then
          STORF = STORF + CAPEN
        End If
      End If
   Next UK&
   '
   If STORF < 0 Then STORF = 0
   STORFAB = STORF
   '
End Function

Function STOCOPEN(CI%)
     '
     CAPEN = 0
18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        CIIX% = CVI(Mid$(X$, 9, 2))
        STAT% = Asc(Mid$(X$, 124, 1))
        If STAT% < 9 Then      ' no esta anulada
          NUORSE& = CVS(Mid$(X$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(X$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
            CAPEN = CACOM - CAREC
          End If
        End If
     Next UK&
     '
     If CAPEN < 0 Then CAPEN = 0
     STOCOPEN = CAPEN

End Function

Function STOLOTES(CI%)
    '
    STLOT = 0
    '
    If CI% > 0 Then
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT * FROM LOTINGRE "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
      '
      With LotIngTemp
25      On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            STLOT = STLOT + !CanSaldo
          .MoveNext
          Loop
        End If
      End With
    End If
    '
    STOLOTES = STLOT
    '
End Function
'
Function UBISTO$(CI%)
    '
20  UBX$ = ""
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           If UBX$ <> "" Then UBX$ = UBX$ + ", "
           UBX$ = UBX$ + TRIM$(Mid$(X$, 3, 24))
        End If
    Next U&
    '
    UBISTO$ = UBX$
    '
End Function
'
Function SEPRODUC(CI%)
    '
20  SPROD = 0
    Call APERBASDAT("ORFAB")
    '
    SQLX$ = "SELECT NumeOrFa, SecciPro FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
    Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    On Error Resume Next
    With ORFABTEMP
      .MoveFirst
      If Err < 1 Then
        SPROD = !SecProd
      End If
    End With
    '
    SEPRODUC = SPROD
    '
End Function
'
Function PEDNOASI(CI%)
    '
10  ASII = 0
    RGXX$ = RECFILT$("ODESEMI", 6, MKI$(CI%))
    '
    For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("ODESEMI", REG&)
        CIP% = CVI(Mid$(X$, 11, 2))
        If CIP% <> CI% Then        ' fuerza reindexacion
            MPANTE = Screen.MousePointer
            Screen.MousePointer = 11
            X$ = REGLEIDO$("ODESEMI", 1)
            Call GRAREG("ODESEMI", X$, 1)
            Call CIERRARCH("ODESEMI")
            Screen.MousePointer = MPANTE
            GoTo 10
        End If
        '
        ASII = ASII + CVS(Mid$(X$, 17, 4))
        '
    Next I&
    '
    PEDNOASI = ASII
    '
End Function
'
Function NUOCOCLI$(NPED&, CIXI%)
    '
    NUOCOCLI$ = ""
    If NPED& < 1 Then Exit Function
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        NUOCOCLI$ = !NroOcom
      End If
    End With
    '
End Function
'
Function NOCOMCLI$(NPED&)
    '
    NOCOMCLI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDENCA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        NOCOMCLI$ = !NroOcom
      End If
    End With
    '
End Function
'
Function PRECIOPED#(NPED&, CODIMAT%, MONEPE%)
    '
    PRECIOPED# = 0
    MONEPE% = 1
    '
    If NPED& > 0 Then
      Call APERBASDAT("PEDCLI")
      SQLX$ = " SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
      SQLX$ = SQLX$ + " AND CodiInt = " & CODIMAT% & " "
      SQLX$ = SQLX$ + " AND Status < 9 "    ' NO ANULADOS
      '
      Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With PEDTEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          PRECIOPED# = !PreUnid
          MONEPE% = !MoneEmis
        End If
        On Error GoTo 0
      End With
    End If
    '
End Function
'
Function MONEMIPED%(NPED&)   ' MONEDA DE EMISION DEL PEDIDO
    '
    MONEMIPED% = 1    ' POR DEFAULT PESOS
    '
    If NPED& > 0 Then
      Call APERBASDAT("PEDCLI")
      SQLX$ = " SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
      SQLX$ = SQLX$ + " AND Status < 9 "    ' NO ANULADOS
      '
      Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      With PEDTEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          MONEMIPED% = !MoneEmis
        End If
        On Error GoTo 0
      End With
    End If
    '
End Function
'
Function SUBCODI$(NPED&, CIXI%)
    '
    SUBCODI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        If IsNull(!SubCod) = False Then
          SUBCODI$ = !SubCod
        End If
      End If
    End With
    '
End Function
'

Function PRELISTA#(APREC$, CIIXXI%, Optional FEPRECIO%)
   '
   Static APRECA$, UREPRE&
   '
   PPXX% = InStr(APREC$, "/NEW")
   If PPXX% > 0 Then
      APREC$ = Left$(APREC$, PPXX% - 1)
      UREPRE& = ULTREG&(APREC$)
   End If
   '
   If APREC$ <> APRECA$ Then
     APRECA$ = APREC$
     UREPRE& = ULTREG&(APREC$)
   End If
   '
   PRELISTA# = 0: FEPRECIO% = 0: VUELTA% = 0
   If CIIXXI% > 0 Then
     If CIIXXI% <= NUMAS% Then
       AP$ = APREC$
10     RFFF$ = RECFILT$(AP$, 1, MKI$(CIIXXI%))
       If Len(RFFF$) >= 4 Then
         REPRE& = CVS(Left$(RFFF$, 4))
         If REPRE& > 0 Then
           If REPRE& <= UREPRE& Then
             TTZ$ = REGLEIDO$(AP$, REPRE&)
             If CVI(Left$(TTZ$, 2)) = CIIXXI% Then
                 PRELISTA# = CVD(Mid$(TTZ$, 62, 8))
                 FEPRECIO% = CVI(Mid$(TTZ$, 70, 2))
               Else
                 If VUELTA% < 1 Then
                   VUELTA% = VUELTA% + 1
                   TTZ$ = REGLEIDO$(AP$, 1)
                   Call GRAREG(AP$, TTZ$, 1)
                   Call CIERRARCH(AP$)
                   GoTo 10
                 End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   If CONVERPRECIO% = 1 Then
     NLIS% = XVALO(Mid$(APREC$, 6, 3))
     PRELISTA# = PRELISTA# * TICAMONE(MONELISTA%(NLIS%)) / TICAMONE(MONEMI%)
   End If
   '
End Function
'
Function MONELISTA%(ILISTA%) ' se traslado esta funcion desde el moduo glo_variabs.bas
   '
   Static CTXX%
   Static MNL%(256)
   Static FICATO#
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   If EXISTE%(LUDAT$ + "LISTAS.DAT") > 0 Then
     If FileDateTime(LUDAT$ + "LISTAS.DAT") <> FICATO# Then
       CTXX% = 0
       FICATO# = FileDateTime(LUDAT$ + "LISTAS.DAT")
     End If
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("LISTAS")
       XX$ = REGLEIDO$("LISTAS", U&)
       CTG% = Asc(Left$(XX$, 1))
       MNL%(CTG%) = Asc(Mid$(XX$, 50, 1))
     Next U&
     MNL%(0) = 1
   End If
   '
   MONELISTA% = MNL%(ILIS%)
   '
End Function
'
Function DECILISTA%(ILISTA%)           ' DECIMALES DE LISTA DE PRECIOS
   '
   Static CTXX%
   Static DCL%(256)
   Static FICATO#
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   If EXISTE%(LUDAT$ + "LISTAS.DAT") > 0 Then
     If FileDateTime(LUDAT$ + "LISTAS.DAT") <> FICATO# Then
       CTXX% = 0
       FICATO# = FileDateTime(LUDAT$ + "LISTAS.DAT")
     End If
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("LISTAS")
       XX$ = REGLEIDO$("LISTAS", U&)
       CTG% = Asc(Left$(XX$, 1))
       CANDECI% = Asc(Mid$(XX$, 49, 1))
         If CANDECI% < 2 Then CANDECI% = 2
         If CANDECI% > 5 Then CANDECI% = 5
      DCL%(CTG%) = CANDECI%
     Next U&
     DCL%(0) = 2
   End If
   '
   DECILISTA% = DCL%(ILIS%)
   '
End Function
'
Function DENOLISTA$(ILISTA%) ' se traslado esta funcion desde el moduo glo_variabs.bas
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   DNL$ = "Precios Concertados (Pesos)"
   If ILIS% > 0 Then
     DNL$ = TRIM$(ECOARCH$("LISTAS", Chr$(ILIS%)))
     MNL% = MONELISTA%(ILIS%)
     If MNL% = 2 Then
         DNL$ = DNL$ + " (Dólares)"
       Else
         DNL$ = DNL$ + " (Pesos)"
     End If
   End If
   DENOLISTA$ = DNL$
   '
End Function

Function DEPOPRE%(CI%)
'
DPP% = 0
TMT% = TIMATE%(CI%)
For U& = 1 To ULTREG&("TABTIMA")
   X$ = REGLEIDO$("TABTIMA", U&)
   If Asc(Left$(X$, 1)) = TMT% Then
      DPP% = Asc(Mid$(X$, 17, 1))
      GoTo 10
   End If
Next U&
'
10 If DPP% > 0 Then
     If ECOARCH$("ECODEP", Chr$(DPP%)) = "" Then DPP% = 0
     If DPP% = 99 Then DPP% = 0
End If
'
DEPOPRE% = DPP%
'
End Function
'
Function DEPOEXPE%()
    '
    Static CTXX%, DPXXP%
    If CTXX% = 0 Then
      CTXX% = 1
      DPXXP% = 0
      DPXX% = XVALO(CONTROL$("", "DEPOEXPE"))
      If DPXX% > 0 Then
        If DPXX% < 99 Then
          If ECOARCH$("TADEPOSI", Chr$(DPXX%)) <> "" Then
            DPXXP% = DPXX%
          End If
        End If
      End If
    End If
    '
    DEPOEXPE% = DPXXP%
    '
End Function
'
Function FEPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         FEX = CVI(Mid$(RPLA$, 23, 2))
         NPLX$ = FECHATEX$(FEX)
       End If
     End If
   End If
   '
   FEPLANO$ = NPLX$
   '
End Function

Function DEPENART%(COD1%) 'DEPOSITO POR OMISION PARA ENTRADA POR O/F
   '
   Call APERBASDAT("MASTER")
   If COD1% > 0 Then
       Master.FindFirst ("CODINT = " & COD1%)
       If Master.NoMatch = False Then
          On Error Resume Next
          DEPENART% = Master!DepoEntra
       End If
   End If
   '
End Function
'
Function DEPCONART%(COD2%) 'DEPOSITO POR OMIDION PARA CONSUMO DE O/F
   '
   Call APERBASDAT("MASTER")
   If COD2% > 0 Then
       Master.FindFirst ("CODINT = " & COD2%)
       If Master.NoMatch = False Then
          On Error Resume Next
          DEPCONART% = Master!DepoCon
       End If
   End If
  '
End Function

Function PZPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 53, 12))
       End If
     End If
   End If
   '
   PZPLANO$ = NPLX$
   '
End Function

Sub PRIDOGRAF(CIXI%, IGRA%)
   '
   Static URE&, RUTARGRA$
   '
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORGEN$ = Mid$(TRIM$(CONTROL$("", "FORGRAFI")), 2)
   CAIMAGEN% = 0
   '
   If PRIFORGEN$ <> "" Then
     If CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$) = 1 Then
       HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       If HIMAGEN > VIMAGEN Then
            PRIFORM$ = "H" + PRIFORGEN$
          Else
            PRIFORM$ = "V" + PRIFORGEN$
       End If
       '
       If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             '
             CAPARDOC% = 0
             CACOLTAB1% = 0
             CAITAB1% = 0
             CACOLTAB2% = 0
             CAITAB2% = 0
             '
             CODPARAM$(1) = "DES-MAT": DOCPARAM$(1) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
             CAPARDOC% = 1
             '
             COPLANO$ = AGRAFIC$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMAGEN% = 1
                 CODIMAGEN$(1) = "GRAFIPIE"
                 ARCHIMAGEN$(1) = COPLANO$
               End If
             End If
           End If
         End If
       End If
       Call PRINTDOC("@" + PRIFORM$)
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
'
Function CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   LPX% = 0: AGRAFIC$ = ""
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Function
   End If
   '
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   '
   APLANO$ = "": Index = 0
   If CIXI% > 0 Then
     If CIXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65 + 48 * IGRA%, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   AGRAFIC$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   CARDOGRAF% = LPX%
   '
End Function

Function DEGRUVEN$(CL%)
   DX$ = "Varios"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOVEN", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "Varios"
   DEGRUVEN$ = DX$
End Function
'
'
Function NUHOPROF$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   NHPF$ = ""
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SecopTemp.MoveFirst
   With SecopTemp
     NHPF$ = !HoProFab
   End With
   '
90 NUHOPROF$ = NHPF$
   '
End Function
'
Function MAQUISTAN$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   NPLX$ = ""
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SecopTemp.MoveFirst
   With SecopTemp
     NPLX$ = !CodMaq0
   End With
   '
90 MAQUISTAN$ = NPLX$
   '
End Function
'
Function TIESTAN(CIXI%, NOPE%, CANTID#, MODO%, Optional ALTEROP%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   Call ABRESECOP
   TSTA = 0
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SecopTemp.MoveFirst
   With SecopTemp
     CALOTES = 1
     CAOPS = !CantiOps
       If CAOPS < 0.01 Then CAOPS = 1
     CICLOHE = !PieCiclo
       If CICLOHE < 1 Then CICLOHE = 1
     MINPREP = 60 * !HorsPrep * CALOTES
       If MODO% > 0 Then MINPREP = 0
     MINFIJO = 60 * !HorsFijo * CALOTES
     MINVARI = CANTID# * !MinuStan
     MINHERR = CANTID# * 60 * !HorsMHer / CICLOHE
     MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
     MINLIMP = 60 * !HorsLimp * CALOTES
     MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
     '
     TSTA = CAOPS * MINLOTE
     '
   End With
   '
90 TIESTAN = TSTA
   '
End Function
'
Function TIEMAQUI(CIXI%, NOPE%, CANTID#, MODO%, Optional ALTEROP%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   Call ABRESECOP
   TSTA = 0
   
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SecopTemp.MoveFirst
   With SecopTemp
     CALOTES = 1
     CICLOHE = !PieCiclo
       If CICLOHE < 1 Then CICLOHE = 1
     MINPREP = 60 * !HorsPrep * CALOTES
       If MODO% > 0 Then MINPREP = 0
     MINFIJO = 60 * !HorsFijo * CALOTES
     MINVARI = CANTID# * !MinuStan
     MINHERR = CANTID# * 60 * !HorsMHer / CICLOHE
     MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
     MINLIMP = 60 * !HorsLimp * CALOTES
     MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
     '
     TSTA = MINLOTE
     '
   End With
   '
90 TIEMAQUI = TSTA
   '
End Function
'\\\OT-07-0028-RG-25/01/07///
Function CADSTAN(CIXI%, NOPE%, Optional ALTEROP%)
   CADSTAN = 0
   TSTANBA = TIEMAQUI(CIXI%, NOPE%, 1, 1, ALTEROP%)
   If TSTANBA > 0 Then CADSTAN = 60 / TSTANBA
   If CADSTAN >= 1000 Then CADSTAN = Int(CADSTAN + 0.5)
End Function
'\\\OT-07-0028-RG-FIN///
Function CANOPS%(CIXI%)
   '
   CNP% = 0
   Set SecopTemp = Articulos.OpenRecordset("SELECT NumeOper FROM SecOper WHERE CodIConj=" & CIXI% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   On Error Resume Next
   SecopTemp.MoveLast
   If Err < 1 Then
     CNP% = SecopTemp.RecordCount
   End If
   '
90 CANOPS% = CNP%
   '
End Function
'
Function DESCRIOP$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   Call APERBASDAT("MASTER")
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   NPLX$ = ""
   SQLX$ = "SELECT DescOper FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   '
   If CAREC% > 0 Then
     SecopTemp.MoveFirst
     NPLX$ = SecopTemp!DescOper
   End If
   '
90 DESCRIOP$ = NPLX$
   '
End Function
'
Function ULTIMOP%(CIXI%)
   '
   Call APERBASDAT("MASTER")
   ULTX1% = 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND StatOper = 1 "
   SQLX$ = SQLX$ + "ORDER BY NumeOper DESC"
   Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   '
   On Error Resume Next
   SecopTemp.MoveFirst
   If Err < 1 Then
      ULTX1% = SecopTemp!NumeOper
   End If
   '
90 ULTIMOP% = ULTX1%
   '
End Function
'
Function CANOPREP(CIXI%, NOPE%)  ' CANTIDAD DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 1
   Set SecopTemp = Articulos.OpenRecordset("SELECT Nuoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   '
   If CAREC% > 0 Then
     SecopTemp.MoveFirst
     If SecopTemp!NUOPER_PREP > 0 Then
       NPLY = SecopTemp!NUOPER_PREP
     End If
   End If
   '
90 CANOPREP = NPLY
   '
End Function
'
Function CATOPREP%(CIXI%, NOPE%)  ' CATEGORIA DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 0
   Set SecopTemp = Articulos.OpenRecordset("SELECT Catoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   '
   If CAREC% > 0 Then
     SecopTemp.MoveFirst
     If SecopTemp!CATOPER_PREP > 0 Then
       NPLY = SecopTemp!CATOPER_PREP
     End If
   End If
   '
90 CATOPREP% = NPLY
   '
End Function
'

Function OBSEROP$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   Set SecopTemp = Articulos.OpenRecordset("SELECT ObseOper FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   CAIT% = 0
   On Error Resume Next
   SecopTemp.MoveLast
   If Err Then GoTo 90
   CAREC% = SecopTemp.RecordCount
   '
   If CAREC% > 0 Then
     SecopTemp.MoveFirst
     NPLX$ = SecopTemp!OBSEOPER
   End If
   '
90 OBSEROP$ = NPLX$
   '
End Function

Function IAJUINF#(FEBAS%, FEACT%)
   '
   Static FIDATA#, AJUIN$
   Static CAPER%, DESPE%(512), HASPE%(512), IAJU#(512)
   '
   If AJUIN$ = "" Or FIDATIM#("FLEXCRL") <> FIDATA# Then
     FIDATA# = FIDATIM#("FLEXCRL")
     AJUIN$ = CONTROL$("ANACOSTO", "AJUINFLA")
     If AJUIN$ <> "SI" Then AJUIN$ = "NO"
   End If
   '
   If AJUIN$ <> "SI" Then
     IAJC# = 1
     GoTo 99
   End If
   '
   If CAPER% < 1 Then
     CAPER% = ULTREG&("INDIAJU")
     For U& = 1 To CAPER%
       XX$ = REGLEIDO$("INDIAJU", U&)
       DESPE%(U&) = CVI(Mid$(XX$, 17, 2))
       HASPE%(U&) = CVI(Mid$(XX$, 19, 2))
       IAJU#(U&) = CVD(Mid$(XX$, 21, 8))
     Next U&
     If CAPER% < 1 Then
       CAPER% = 1
       DESPE%(1) = (-32767)
       HASPE%(1) = 32767
       IAJU#(1) = 1
     End If
   End If
   '
10 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEBAS% Then
       IAJ1# = IAJU#(KKI%)
       GoTo 20
     End If
   Next KKI%
   IAJ1# = IAJU#(CAPER%)
   '
20 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEACT% Then
       IAJ2# = IAJU#(KKI%)
       GoTo 30
     End If
   Next KKI%
   IAJ2# = IAJU#(CAPER%)
   '
30 IAJC# = 1
   If IAJ1# >= 1 Then
     IAJC# = IAJ2# / IAJ1#
   End If
   '
99 IAJUINF# = IAJC#
   '
End Function

Function RVPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 19, 4))
       End If
     End If
   End If
   '
   RVPLANO$ = NPLX$
   '
End Function

Function NUHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 5, 16)
       End If
     End If
   End If
   '
   NUHOREI$ = NPLX$
   '
End Function

Function ESPECIFI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CIXI%))
   URN& = ULTREG&("ESPECIFI")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
         If NPLX$ <> "" Then
           NPLX$ = NPLX$ + Chr$(13) + Chr$(10)
         End If
         NPLX$ = NPLX$ + RTrim$(TTYY$)
       End If
     End If
   Next U&
   '
   ESPECIFI$ = NPLX$
   '
End Function

Function NUHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 33, 16))
       End If
     End If
   End If
   '
   NUHOJES$ = NPLX$
   '
End Function

Function RVHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 49, 4))
       End If
     End If
   End If
   '
   RVHOJES$ = NPLX$
   '
End Function

Function LEERPLANO%(CIXXI%, ARCHIPLANO$)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   LPX% = 0: ARCHIPLANO$ = ""
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   APLANO$ = "": Index = 0
   If CIXXI% > 0 Then
     If CIXXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   ARCHIPLANO$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   LEERPLANO% = LPX%
   '
End Function

Function NUPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 25, 16)
       End If
     End If
   End If
   '
   NUPLIEN$ = NPLX$
   '
End Function

Function RVPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 41, 4)
       End If
     End If
   End If
   '
   RVPLIEN$ = NPLX$
   '
End Function

Function RVHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 21, 4)
       End If
     End If
   End If
   '
   RVHOREI$ = NPLX$
   '
End Function
'\\\OT-06-0105-WAR-23/03/06///
Function DESCOMPO%(CI%)
   '
   DESCOMPO% = 0
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then DESCOMPO% = MasterTemp!DesCompAut
   '
 '\\\OT-06-0105-WAR-FIN//
End Function
'
Function CODORIG$(CIXXX%, NCXXX%)
   '
   ASELE$ = "REPLACOC"
   If NCXXX% < 0 Then ASELE$ = "REPLACOP"
   '
   CDXXX$ = CODEXT$(CIXXX%)
   UREP& = ULTREG&(ASELE$)
   RFF$ = RECFILT$(ASELE$, 3, MKI$(CIXXX%))
   '
   For KKI% = 1 To Len(RFF$) Step 4
     REREP& = CVS(Mid$(RFF$, KKI%, 4))
     If REREP& > 0 Then
       If REREP& <= UREP& Then
         XXYY$ = REGLEIDO$(ASELE$, REREP&)
         If CVI(Mid$(XXYY$, 1, 2)) = Abs(NCXXX%) Then
           CDYYY$ = TRIM$(Mid$(XXYY$, 5, 28))
           If CDYYY$ <> "" Then
             CDXXX$ = CDYYY$
             GoTo 99
           End If
         End If
       End If
     End If
   Next KKI%
   '
99 CODORIG$ = CDXXX$
   '
End Function

Function LOTESEL$(CIXI%)
   '
   Dim TLOTE$(256), CALOTE#(256)
   LSX$ = ""
   '
   If CIXI% > 0 Then
     CANLOT% = 0
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       ENTRA# = CVD(Mid$(X$, 88, 8))
       If ENTRA# >= 0.05 Then
         IDENLO$ = TRIM$(Mid$(X$, 5, 16))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 23, 10))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 33, 10))
         IDENLO$ = AJUSTI$(IDENLO$, 16)
         '
         For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Then
             CALOTE#(KKI%) = CALOTE#(KKI%) + ENTRA#
             TTXX$ = TLOTE$(KKI%)
             Call REPLA(TTXX$, CSTRING$(MKD$(CALOTE#(KKI%)), 4, 12, 1, 2), 27, 12)
             TLOTE$(KKI%) = TTXX$
             GoTo 19
           End If
         Next KKI%
         FEX = CVI(Left$(X$, 2))
         CANLOT% = CANLOT% + 1
         TTXX$ = AJUSTI$(IDENLO$, 40)
         Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
         Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
         TLOTE$(CANLOT%) = TTXX$
         CALOTE#(KKI%) = ENTRA#
         '
       End If
19   Next U&
     '
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       SALID# = CVD(Mid$(X$, 96, 8))
       If SALID# >= 0.05 Then
         IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
         IDENLOX$ = TRIM$(IDENLO$)
         '
35       For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Or IDENLOX$ = "" Then
             If SALID# <= CALOTE#(KKI%) Then
                 CALOTE#(KKI%) = CALOTE#(KKI%) - SALID#
                 SALID# = 0
               Else
                 SALID# = SALID# - CALOTE#(KKI%)
                 CALOTE#(KKI%) = 0
             End If
           End If
         Next KKI%
         If SALID# >= 0.05 Then
           IDENLOX$ = ""
           GoTo 35
         End If
         '
       End If
     Next U&
     '
     REBLA$ = REGBLAN$("STOLOTE"): JJ& = 0
     For KKI% = 1 To CANLOT%
       If CALOTE#(KKI%) > 0.05 Then
         ZZ$ = REBLA$
         Call REPLA(ZZ$, TLOTE$(KKI%), 1, 38)
         Call REPLA(ZZ$, MKD$(CALOTE#(KKI%)), 41, 8)
         JJ& = JJ& + 1
         Call GRAREG("!STOLOTE", ZZ$, JJ&)
       End If
     Next KKI%
     Call SETULTREG("!STOLOTE", JJ&)
     Call CIERRARCH("!STOLOTE")
     '
     Call FRESHECHO("!STOLOTE")
     Call SELECHO("!STOLOTE/Id.Lote                       F.Alta                Q.Orig                      Saldo")
     LSX$ = VALACT1$("!STOLOTE")
   End If
   '
   LOTESEL$ = LSX$
   '
End Function

Function GRUCOFAB%(CIJK%)
  '
  ' GRUCOFAB%=1 --> Compra
  ' GRUCOFAB%=2 --> Fabricacion
  '
  Static GRF(32767) As Integer
  '
  If GRF%(CIJK%) < 1 Then
    GRF%(CIJK%) = 1
    Call AREXPLO(CIJK%, 1, 1)
    '
    If CAIT% > 0 Then
      GRF%(CIJK%) = 2
      PBL% = PROVHABI%(CIJK%)
      If PBL% > 0 Then
        If PROVINT%(PBL%) = 0 Then
          GRF%(CIJK%) = 1
        End If
      End If
    End If
  End If
  '
  GRUCOFAB% = GRF%(CIJK%)
  '
End Function
' ***************************************************************************
'
Function COTIPIF$(CI%)
  RECO$ = FILTERGETRECORD$("COEQTIP", 1, MKI$(CI%))
  COTIPIF$ = TRIM$(Mid$(RECO$, 9, 24))
End Function

Function NUINV%()
'
X$ = REGLEIDO$("INVERT", 1&)
NUINV% = CVI(Mid$(X$, 17, 2))
End Function
' ***************************************************************************

Sub DESHASCOD(COD1$, COD2$, CIKX1%, CIKX2%)
   '
10 COD1$ = "": COD2$ = "": CIKX1% = 0: CIKX2% = 0
   XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   '
   CIKX1% = CVI(Left$(X$, 2))
   COD1$ = CODEXT$(CIKX1%): If TRIM$(COD1$) = "" Then Exit Sub
   CIKX2% = CVI(Mid$(X$, 3, 2))
   COD2$ = CODEXT$(CIKX2%): If TRIM$(COD2$) = "" Then Exit Sub
   If CIKX1% < 1 Or CIKX2% < 1 Then GoTo 10
   If COD2$ < COD1$ Then GoTo 10
   '
End Sub
'
Function SECPERTE(EQP$)
   '
   EQP1$ = TRIM$(EQP$)
   If EQP1$ <> "" Then
      Y$ = FILTERGETRECORD$("PADMAQ", 1, EQP1$)
      SEC = CVS(Mid$(Y$, 37, 4))
      If SEC > 0 Then
         SECPERTE = SEC
      End If
   End If
   '
End Function

Function FECREALOT%(IDLOT$)
   '
   FECREALOT% = 0
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOT$ & "' "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With LotIngTemp
25   On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        FECREALOT% = CVINT%(!FechAlta)
     End If
     On Error GoTo 0
   End With
   '
End Function

Function SALNOCONS(IDLOT$, CIXI%)
   '
   SALNOCONS = 0
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOT$ & "' "
   SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With LotIngTemp
25   On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        SALNOCONS = !CanAlta - !CanConsu
     End If
     On Error GoTo 0
   End With
   If SALNOCONS < 0 Then SALNOCONS = 0
   '
End Function

Function BIENCAP%(CI%)
    BICA% = 0
    If CI% < 1 Then GoTo 99
    '
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then BICA% = XVALO(MasterTemp!BIEN_CAP)
   
99  BIENCAP% = BICA%
End Function

Function NOMENC$(CI%)
    'NOMENCLADOR DE LA AFIP QUE SE GUARDA EN MAS DATOS DE A.MAESTRO DE ARTICULOS
    NOME$ = ""
    If CI% < 1 Then GoTo 99
    '
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then NOME$ = TRIM$(MasterTemp!CODIGO_NCM)
   '
99  NOMENC$ = NOME$
End Function
Function UNIMEX$(CI%)
    ' UNIDAD DE MEDIDA AFIP QUE SE GUARDA EN A.MAESTRO DE ARTICULOS EN MAS DATOS
    UNIM$ = ""
    If CI% < 1 Then GoTo 99
    '
    On Error Resume Next
    If STATUIT%(CI%) <> 0 Then UNIM$ = TRIM$(MasterTemp!PRO_UMED)
    '
99  UNIMEX$ = UNIM$
End Function

Function UNIMEAFIP%(CI%)

   'MODIFICADO 09 POR QUE NO DEVOLIVA POR DEFECTO LA 'UNIDAD' 7
   UNIMEAFIP% = 7 ' POR DEFECTO DEVUELVE UNIDAD
   If CI% < 1 Then Exit Function
   '
   UNIM$ = UNIMED$(CI%)
   X$ = FILTERGETRECORD$("UNIMED", 1, TRIM$(UNIM$))
   UMEDAFIP% = CVI(Mid$(X$, 41, 2))
   '
   If UMEDAFIP% > 0 Then
     UNIMEAFIP% = UMEDAFIP%
     Exit Function
   End If
   '
End Function
Function COTIVA$(CI%)
    ' IDENTIFICADOR DE IVA AFIP QUE SE GUARDA EN A.MAESTRO DE ARTICULOS EN MAS DATOS
    COTI$ = ""
    If CI% < 1 Then GoTo 99
    '
    On Error Resume Next
    If STATUIT%(CI%) <> 0 Then COTI$ = TRIM$(MasterTemp!IVA_ID)
    '
99  COTIVA$ = COTI$
End Function
Function COSEC$(CI%)
    ' CODIGO SECUNDARIO  AFIP QUE SE GUARDA EN A.MAESTRO DE ARTICULOS EN MAS DATOS
    COSE$ = ""
    If CI% < 1 Then GoTo 99
    '
    On Error Resume Next
    If STATUIT%(CI%) <> 0 Then COSE$ = TRIM$(MasterTemp!CODIGO_SEC)
    '
99  COSEC$ = COSE$

End Function

