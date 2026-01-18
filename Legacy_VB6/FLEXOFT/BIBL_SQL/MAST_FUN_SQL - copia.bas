Attribute VB_Name = "MAST_FUN_SQL"
Public PUNCODIG As String
Public BUFERCOD As String * 2048
Public NUSU%, COSU%(64), SUCU$(64), ADEPO$(64), DISTO%(64), CATEO#(64), CAFIN#(64)
Dim CIANT%, XXX$
Public PHABIT%()
Public UPDATMASTER%, UPDAPLANOS%
'
Dim MasterTemp As ADODB.Recordset
'
Function CostoRecpMovisto(CI1%, RECEP&, Optional MONEX%)
'TENER EN CUENTA QUE EN EL CASO QUE EN UNA MISMA ORDEN DE RECEPCION HAYA MAS DE UNA ORDEN DE COMPRA
'PUEDE DEVOLVER VALORES QUE NO CORRESPONDAN CUANDO SE REPITA EL CODIGO  CON DIFRENTEES PRECIOS,
'ESTO ES RARO QUE SUCEDA EN AHP POR QUE LA MAYORIA SON DOLARES Y NO DEBERIA CAMBIAR EL MISMO,
'AVISAR QUE HAGAN UNA RECEPCION POR ORDEN DE COMPRA

   CONDI$ = "COD_INTE = " & CI1%
   CONDI$ = CONDI$ + " AND SUBSTRING(DOPRICOR,4,7) = " & RECEP&
   CONDI$ = CONDI$ + " AND CODIMOVI = 'BR'"
   CONDI$ = CONDI$ + " AND ORDER BY IDENLOTE DESC"
   MON% = XVALO(VALOBADA("MOVISTO", "MONEVAL", CONDI$, "NOMESS"))
   TIPCAM# = XVALO(VALOBADA("MOVISTO", "TIPOCAM", CONDI$, "NOMESS"))
   COSMOVISTO# = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
   COEFCON = 1
   If MONEX% <> MON% And MONEX% > 0 Then
      COEFCON = TICAMONE(MON%) / TICAMONE(MONEX%)
   End If
   COSMOVISTO# = COSMOVISTO# * COEFCON
   
   'SI ES 100 DOLARES Y PASO MONEDA $ = COSMOVISTO# * 9  = 900
   
End Function

Sub ActualizarCostoMovisto(CI1%, NBORE&, NUEVOVALOR#, MONEDAX%)
   'ACTUALIZA EL PRECIO DEL MOVISTO VERIFICA SI ES LA MISMA MONEDA
   CONDI$ = "COD_INTE = " & CI1%
   CONDI$ = CONDI$ + " AND SUBSTRING(DOPRICOR,4,7) = " & RECEP&
   CONDI$ = CONDI$ + " AND CODIMOVI = 'BR'"
   Call ACTUREGISTRO("MOVISTO", "VALOUNIT", CONDI$, FORMATNUM$(NUEVOVALOR#, "F16.5"), REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVISTO", "MONEVAL", CONDI$, MONEDAX%, REGAF&, "NOMESS")
   If MONEDAX% = 1 Then Call ACTUREGISTRO("MOVISTO", "TIPOCAM", CONDI$, 1, REGAF&, "NOMESS")
End Sub
Sub AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%, Optional TRAZ%)
   '
   Screen.MousePointer = 11
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
   SQLX$ = "SELECT MAX(CODINT)AS CODIMAXIM FROM MASTER "
   Dim MAXIMMMO As ADODB.Recordset
   Set MAXIMMMO = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MAXIMMMO
      If !CODIMAXIM < 1 Then
         CODII% = 1
       Else
         CODII% = XVALO(!CODIMAXIM) + 1
      End If
   End With
   MAXIMMMO.Close
   Set MAXIMMMO = Nothing
   '
   SQLX$ = "SELECT * FROM MASTER "
   SQLX$ = SQLX$ + " WHERE CODIGO = '" & TRIM$(CODINUE$) & "' "
   Dim COCODTMP As ADODB.Recordset
   Set COCODTMP = New ADODB.Recordset
25 On Error Resume Next
   COCODTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With COCODTMP
     On Error Resume Next
     .MoveFirst
     If Err Then 'el codigo no es pre-existente LO aGREGO
       On Error GoTo 0
       SQLX$ = "SELECT * FROM MASTER "
       Set MASTTTER = New ADODB.Recordset
26     On Error Resume Next
       MASTTTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
       End If
       On Error GoTo 0
       '
       With MASTTTER
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
           !CODIGO = CODINUE$
           !Descripcion = DESCRINUE$
           !Status = 1
           !CODINT = CODII%
           !UMEDIDA = UNID$
           !TIMATE = TIMA%
              If TRAZ% <> 1 Then TRAZ% = 0
           !TRAZABLE = TRAZ%
           .Update
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
'
Sub GENECOMAS()
  '
      If UPDATMASTER% > 0 Then
        '
        Screen.MousePointer = 11
        '
        CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
        If CONTROL$("", "IDSECUN") = "ALIAS" Then
            Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
        Else
            Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
        End If
        '
      End If
      '
      If UPDATMASTER% > 0 Or UPDAPLANOS% > 0 Then
          CONDI$ = " CODIGO <> '' AND STATUS > 0 AND NPLANO <> ''"
          Call CARGALISEL("IPLANOS", "MASTER", "NPLANO/A24;DESCRIPCION/A54;CODINT/F5.0", CONDI$, "NOMESS")
      End If
      '
      UPDATMASTER% = 0
      UPDAPLANOS% = 0
      Screen.MousePointer = 1
      '

    'Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT,NPLANO FROM Master WHERE STATUS > 0 ORDER BY CODIGO ASC;", dbOpenDynaset, dbReadOnly)
    'sql
''    Dim MASTERECO As ADODB.Recordset
''    Set MASTERECO = New ADODB.Recordset
''    MASTERECO.CursorType = adOpenStatic
''    MASTERECO.LockType = adLockReadOnly
''
''    STRSQL = "SELECT CODIGO,DESCRIPCION,CODINT,NPLANO FROM Master WITH(NOLOCK) WHERE STATUS > 0 ORDER BY CODIGO ASC"
''25  On Error Resume Next
''    MASTERECO.Open FILTSQL$(STRSQL), FLEXCONN, , , adCmdText
''    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''      On Error GoTo 0
''      Call CapturaErrorOpen
''      GoTo 25
''    End If
''    On Error GoTo 0
''    '
''    CAREC& = CantRegistros("MASTER", "STATUS > 0")
''    If CAREC& < 1 Then GoTo 15
''    '
''    FIN& = CAREC&
''    REBLA$ = REGBLAN$("ECOMAST")
''    JJ& = 0: JJ1& = 0
''    '
''10  Call TRACE(20, JJ&, FIN&)
''    EL1$ = REBLA$
''    JJ1& = JJ1& + 1
''    CODEXX$ = MASTERECO!CODIGO
''    '
''    If TRIM$(CODEXX$) <> "" Then
''      LOFI1% = LENFIELD%("ECOMAST", 1)
''      POFI2% = POSFIELD%("ECOMAST", 2)
''      POFI3% = POSFIELD%("ECOMAST", 3)
''      Call REPLA(EL1$, CODEXX$, 1, LOFI1%)
''        DEPLANO$ = TRIM$(MASTERECO!Descripcion)
''        If IsNull(MASTERECO!NPLANO) = False Then
''          NXPLANO$ = TRIM$(MASTERECO!NPLANO)
''          If NXPLANO$ <> "" Then
''            'If NXPLANO$ <> TRIM$(MASTERECO!CODIGO) Then
''            If InStr(TRIM$(MASTERECO!CODIGO), NXPLANO$) < 1 Then
''              While Len(DEPLANO$ + NXPLANO$) > 43
''                DEPLANO$ = Left$(DEPLANO$, Len(DEPLANO$) - 1)
''              Wend
''              DEPLANO$ = DEPLANO$ + " (" + NXPLANO$ + ")"
''            End If
''          End If
''        End If
''      Call REPLA(EL1$, DEPLANO$, POFI2%, 46)
''      If MASTERECO!CODINT <= 32767 Then
''          Call REPLA(EL1$, MKI$(MASTERECO!CODINT), POFI3%, 2)
''      End If
''      JJ& = JJ& + 1
''      Call GRAREG("ECOMAST", EL1$, JJ&)
''    End If
''    '
''12  If JJ1& < CAREC& Then
''      MASTERECO.MoveNext
''      GoTo 10
''    End If
''    '
''    MASTERECO.Close
''    Call SETULTREG("ECOMAST", JJ&)
''    Call CIERRARCH("ECOMAST")
''    Call FRESHECHO("ECOMAST")
''    Call BAJALDISCO
''  End If
  '
''15 If UPDATMASTER% > 0 Or UPDAPLANOS% > 0 Then
''    Screen.MousePointer = 11
''
''    'Set MASTERECO = Articulos.OpenRecordset("SELECT NPlano,CODIGO,DESCRIPCION,CODINT FROM Master WHERE STATUS > 0 AND NPLANO <> '' ORDER BY NPLANO ASC;", dbOpenDynaset, dbReadOnly)
''    SQLX$ = " select NPlano,CODIGO,DESCRIPCION,CODINT"
''    SQLX$ = SQLX$ + " from Master"
''    SQLX$ = SQLX$ + " WHERE STATUS > 0 "
''    SQLX$ = SQLX$ + " AND NPLANO <> ''"
''    ' SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
''
''    'Dim MASTERECO As ADODB.Recordset
''    Set MASTERECO = FLEXCONN.Execute(FILTSQL$(SQLX$))
''
''    On Error Resume Next
''    MASTERECO.MoveLast
''    If Err Then GoTo 99
''    '
''    CAREC& = MASTERECO.RecordCount
''    If CAREC& < 1 Then GoTo 99
''    '
''    FIN& = CAREC&
''    REBLA$ = REGBLAN$("IPLANOS")
''    JJ& = 0
''    MASTERECO.MoveFirst
''    '
''20  Call TRACE(20, JJ&, FIN&)
''    EL1$ = REBLA$
''    Call REPLA(EL1$, MASTERECO!NPLANO, 1, 24)
''    Call REPLA(EL1$, MASTERECO!Descripcion, 25, 54)
''    Call REPLA(EL1$, MKI$(MASTERECO!CODINT), 79, 2)
''    JJ& = JJ& + 1
''    Call GRAREG("IPLANOS", EL1$, JJ&)
''    '
''    If JJ& < CAREC& Then
''      MASTERECO.MoveNext
''      GoTo 20
''    End If
''    '
''    MASTERECO.Close
''    Set MASTERECO = Nothing
''    Call SETULTREG("IPLANOS", JJ&)
''    Call CIERRARCH("IPLANOS")
''    Call FRESHECHO("IPLANOS")
''    Call BAJALDISCO
''   End If
''   '
''99 UPDATMASTER% = 0
''   UPDAPLANOS% = 0
''   Screen.MousePointer = 1
End Sub
Sub GENEECOMABAJ()
    CONDI$ = " CODIGO <> '' AND STATUS < 0 ORDER BY CODIGO ASC"
    Call CARGALISEL("ECOMABAJ", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
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
   
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strsql = ("SELECT max(codint) as maximo FROM MASTER")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
   If Not IsNull(MIRS!maximo) Then
      NUMAS% = MIRS!maximo
   Else
      NUMAS% = 0
   End If
   MIRS.Close
   Set MIRS = Nothing
   
End Function

Function PRICOD%()
   '
   PRICOD% = 0
   Call APERBASDAT("MASTER")
   'Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
    strsql = "SELECT MIN(Codigo)as Primero FROM Master WHERE Status >=0  "
    Dim MasterTp As ADODB.Recordset
    Set MasterTp = FLEXCONN.Execute(FILTSQL$(strsql))
    If Not IsNull(MasterTp!Primero) Then 'si no esta vacio
       PRICOD% = CODINT%(MasterTp!Primero)
    End If
    On Error GoTo 0
    MasterTp.Close
    Set MasterTp = Nothing
End Function

Function ULTCOD%()
   '
   ULTCOD% = 0
   Call APERBASDAT("MASTER")
   'Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
   strsql = "SELECT MAX(Codigo) as Ultimo FROM Master WHERE Status >=0 "
   Dim MasterTp As ADODB.Recordset
   Set MasterTp = FLEXCONN.Execute(FILTSQL$(strsql))
   If Not IsNull(MasterTp!Ultimo) Then 'si no es null (devuelve un sola fila)
      ULTCOD% = CODINT%(MasterTp!Ultimo)
    End If
   On Error GoTo 0
   MasterTp.Close
   Set MasterTp = Nothing
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
       On Error Resume Next
       CIAA% = MasterTemp!CODINT
       On Error GoTo 0
       If CI% = CIAA% Then Exit Function
     End If
   End If
   '*********NOTA : SI NO ENTRA AQUI DA  UN ERROR EN CODEXT **********
   STATUIT% = 0
   If CI% > 0 Then
     strsql = "SELECT * FROM Master WITH(NOLOCK) WHERE CODINT=" & CI%
     Set MasterTemp = FLEXCONN.Execute(FILTSQL$(strsql))
     If Not (MasterTemp.EOF And MasterTemp.BOF) Then 'si no esta vacio
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
'
Function CODINT%(CODEXTE$)
   '
   Static FIDAMASANT#, LONCODI%
   '
   If LONCODI% < 1 Then
      LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
      If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16
   End If
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If CONTROL$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   If CODEXTE$ <> "" Then
     '
     Call APERBASDAT("MASTER")
     On Error Resume Next
     CDXX$ = TRIM$(CODEXTE$): CDYY$ = AJUSTI$(CODEXTE$, LONCODI% - 1) + Chr$(96)
     SQLX$ = "Select * from Master where CODIGO ='" + CDXX$ + "' OR CODIGO = '" & CDYY$ & "' "
     
     If INCLUALIAS% = 1 Then SQLX$ = SQLX$ + "OR NPlano = '" + CDXX$ + "' "
     Set MasterTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     If Not (MasterTemp.EOF And MasterTemp.BOF) Then
        CODINT% = MasterTemp!CODINT * MasterTemp!Status
     End If
     On Error GoTo 0
   End If
   '
End Function

Function CODINPLA%(NUPLAX$)
   '
   If NUPLAX$ <> "" Then
     Call APERBASDAT("MASTER")
     strsql = "SELECT codint FROM MASTER WHERE NPLANO ='" & TRIM$(NPLAX$) + "'"
     Dim MIRS As ADODB.Recordset
     Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
     On Error Resume Next
     If Not (MIRS.EOF And MIRS.BOF) Then
      CODINPLA% = MIRS!CODINT
     End If
     MIRS.Close
     Set MIRS = Nothing
   
     'On Error Resume Next
     'Master.MoveFirst
     'Master.FindFirst "NPLANO ='" + TRIM$(NPLAX$) + "'"
     'If Master.NoMatch = False Then
     '  CODINPLA% = Master!CODINT
     'End If
     On Error GoTo 0
   End If
   '
End Function

Function CODEXT$(CI%)
   '
   If CI% = 0 Then
        CODEXT$ = ""
     Else
        If STATUIT%(CI%) <> 0 Then
            On Error Resume Next
            CODEXT$ = TRIM$(MasterTemp!CODIGO)
            If Err Then GoTo 20
          Else
20          On Error GoTo 0
            Call APERBASDAT("MASTER")
            strsql = "SELECT Codigo FROM MASTER WHERE Codint = " & CI%
            Dim MIRS As ADODB.Recordset
            Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
            If Not (MIRS.EOF And MIRS.BOF) Then
              CODEXT$ = MIRS!CODIGO
            End If
            MIRS.Close
            Set MIRS = Nothing
        End If
   End If
   '
End Function
 
Function DESCRIT$(CI%)
   On Error Resume Next
   DESCRIT$ = ""
   If CI% < 1 Then
       If InStr(UCase$(App.EXEName), "WINTAB") < 1 Then
          DESCRIT$ = "Material No Codificado"
       End If
     Else
       DESCRIT$ = DESCRIT0$(CI%)
       If TRIM$(DESCRIT$) <> "" Then DESCRIT$ = DESCRIT$ + " (" + UNIMED$(CI%) + ")"
   End If
End Function

Function DESCRIT0$(CI%)
   On Error Resume Next
   DESCRIT0$ = ""
   If CI% < 1 Then
       If InStr(UCase$(App.EXEName), "WINTAB") < 1 Then
         DESCRIT0$ = "Material No Codificado"
       End If
     Else
       If STATUIT%(CI%) <> 0 Then DESCRIT0$ = TRIM$(MasterTemp!Descripcion)
   End If
End Function

Function DESCRIT2$(CI%)
   On Error Resume Next
   DESCRIT2$ = ""
   If STATUIT%(CI%) <> 0 Then DESCRIT2$ = TRIM$(MasterTemp!DESCRIDOS)
End Function
Function USUPREPES$(CI%)
   On Error Resume Next
   USUPREPES$ = ""
   If STATUIT%(CI%) <> 0 Then USUPREPES$ = TRIM$(MasterTemp!USUPREPES)
End Function
Function USUPREDOL$(CI%)
   On Error Resume Next
   USUPREDOL$ = ""
   If STATUIT%(CI%) <> 0 Then USUPREDOL$ = TRIM$(MasterTemp!USUPREDOL)
End Function

Function UNIMED$(CI%)
   '
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then UNIMED$ = TRIM$(MasterTemp!UMEDIDA)
   If TRIM$(UNIMED$) = "" Then UNIMED$ = "U."
   '
End Function

Function DEUNIMED$(CI%)
   UME$ = UNIMED$(CI%)
   DEUNIMED$ = UME$
   DEUNI$ = TRIM$(ECOARCH$("UMEDIDA", UME$))
   If DEUNI$ <> "" Then
      DEUNIMED$ = DEUNI$
   End If
End Function
'
Function CODFAMIL$(CI%)
   '
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then CODFAMIL$ = TRIM$(MasterTemp!FAMILIA)
   '
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
   If LOTESTAN < 0 Then LOTESTAN = 100 * Abs(LOTESTAN)
End Function
  
Function COSUNI(CI%)
   '
   On Error Resume Next
   If STATUIT%(CI%) > 0 Then COSUNI = CDbl(MasterTemp!COSUNI)
   '
End Function
'
Function MONECOSTO%(CI%)
   '
   MONECOSTO% = 1
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then MONECOSTO% = CDbl(MasterTemp!Monecos)
   '
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
   On Error GoTo 0
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
   On Error GoTo 0
End Function
'
Function MANO_OBRA_REAL(CI%)
   MANO_OBRA_REAL = 0
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then MANO_OBRA_REAL = CDbl(MasterTemp!MANO_OBRA_REAL)
   On Error GoTo 0
End Function
'

Function PROVHABI%(CI%)
   PROVHABI% = 0        ' FABRICACION PROPIA
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then PROVHABI% = CDbl(MasterTemp!PROVHABI)
   On Error GoTo 0
   '
   If PROVHABI% > 0 Then
     If CantRegistros("PROVED12", "Nume_Cli = " & PROVHABI%) < 1 Then
       PROVHABI% = 0
     End If
   End If
End Function
 
Function COLOR$(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then COLOR$ = TRIM$(MasterTemp!COLOR)
   On Error GoTo 0
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
        CANDE% = 5: If CP < 10 Then CANDE% = 7
        LCLR$ = CSTRING$(MKS$(CP), 3, 10, CANDE%, 2)
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
Function CODINTERNA$(CI%)
   '
   CODINTERNA$ = ""
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then CODINTERNA$ = TRIM$(MasterTemp!Cod_Interna)
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
Function ARMEMOTEC$(CI%)
   '
   ARMEMOTEC$ = VALOBADA("MASTER", "ARCHIPDF", "CODINT=" & CI%)
   '
End Function
'
Function CONSUMOPROMEDIO#(CI%)
   '
   CONSUMOPROMEDIO# = XVALO(VALOBADA("MASTER", "consuprom", "CODINT=" & CI%))
   '
End Function
Function COSCALC#(CI%)
   '
   COSCALC# = XVALO(VALOBADA("MASTER", "CosCal", "CODINT=" & CI%))
   '
End Function
Function FECOSCAL%(CI%)
   '
   FECOSCAL% = FECHANUM("01/01/80")
   On Error Resume Next
   FECOSCAL% = CVINT%(VALOBADA("MASTER", "FeCosCal", "CODINT=" & CI%))
   On Error GoTo 0
   '
End Function
'
Function UNICOM$(CI%, Optional NPROVE%)
   '
   UNICO$ = ""
   '
   Call ABRECOMPRAS
   If CI% > 0 Then  ' que no lo haga con material libre
     SQLX$ = "SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
       If NPROVE% > 0 Then
         SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROVE% & " "
       End If
     SQLX$ = SQLX$ + "ORDER BY Femi_Ocom ASC, Nume_Ocom ASC"
     'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim OcomdetaTemp As ADODB.Recordset
     Set OcomdetaTemp = New ADODB.Recordset
25   On Error Resume Next
     OcomdetaTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
     End If
     On Error GoTo 0
     '
     On Error Resume Next
     With OcomdetaTemp
       .MoveLast
       If Err < 1 Then
         On Error Resume Next
         UNICO$ = SAFETEXT$(!Uni_Com)
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
'
Function ULTICOM$(NPRO%, CIIX%)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Dim OcomdetaTemp  As ADODB.Recordset
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPRO%), 7, 2)
   Call REPLA(UCO$, MKI$(CIIX%), 9, 2)
   '
   Call ABRECOMPRAS
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIIX% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPRO% & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + "ORDER BY Femi_Ocom DESC, Nume_Ocom DESC"
   Set OcomdetaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With OcomdetaTemp
     If Not (.BOF And .EOF) Then
         Call REPLA(UCO$, MKS$(!nume_ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, SAFETEXT$(!Uni_Mas), 11, 4)
         Call REPLA(UCO$, SAFETEXT$(!Uni_Com), 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   '
   ULTICOM$ = UCO$
   '
End Function

Function ULTIPRE#(NPRO%, CIIX%)
  '
  'DEVUELVE EL ULTIMO PRECIO DE COMPRA DEL ARTICULO CIIX%
  'AL PROVEEDOR NPRO%
  '
  UC$ = ULTICOM$(NPRO%, CIIX%)
  MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
  ULTIPRE# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
  '
End Function

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
Function FEREFOR%(CI%)   ' FECHA REVISION FORMULA
  '
  FRF% = 0
  On Error Resume Next
  If STATUIT%(CI%) <> 0 Then
      FRF% = CVINT(MasterTemp!Revi_Formu)
      FRF2% = CVINT(MasterTemp!FechApFor)
      If FRF2% > 0 Then FRF% = FRF2%
  End If
  FEREFOR% = FRF%
  '
End Function
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
       CC% = 1 + ULTREG&("TACOLOR")
       Call REGAPP("TACOLOR", TC$)
       Call CIERRARCH("TACOLOR")
       Call FRESHECHO("TACOLOR")
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

Function COSTOSUMI(CI%, MODOCO%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 33, 8))        ' COSTO TOTAL DE SUMINISTROS
    End If
  End If
  '
  MODOCO% = 1
  If COU < 0.0005 Then
    COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%)) * (1 + PORNACIO(CI%) / 100)
    MODOCO% = 2
  End If
  '
  COSTOSUMI = COU
  '
End Function
' ***************************************************************************
'
Function PRECALCU(CI%)
  '
  COU = 0
  RFG$ = RECFILT$("LICOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("LICOSCAL", REGB&)
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
   Static DP%(32767), FESTRU#
   
   If CI% < 1 Then GoTo 90
   CICON% = CI%
   '
   NVI% = NUINV%
   NMM% = NUMAS%
   '
   Call EXPLOMUL(CI%)
   SUMPPROP = 0: SUMPREC = 0
   SUPIPROP = 0: SUPIREC = 0
   SUSEPROP = 0: SUSEREC = 0
   '
   For KKI% = 1 To CONTAMUL%
     For KKJ% = KKI% + 1 To CONTAMUL%
       If CIK%(KKJ%) = CIK%(KKI%) Then
         CANT(KKI%) = CANT(KKI%) + CANT(KKJ%)
         CANT(KKJ%) = 0
       End If
     Next KKJ%
   Next KKI%
   '
   ' costo de materias primas y piezas de terceros
   '
   SUPRI = 0: SUPIE = 0: SUSER = 0
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
     If Abs(CANT(KKI%)) < 0.0000000000005 Then GoTo 19
     If KKI% >= CONTAMUL% Or NIVE%(KKI% + 1) <= NIVE%(KKI%) Then  ' NO TIENE ESTRUCTURA
15      MONEI% = MONECOSTO%(CIK%(KKI%))
        DEMONE$ = DEMONECO$(MONEI%)
        COSUDOL = COSUNI(CIK%(KKI%))
        If MONEI% > 1 Then
            COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
          Else
            FEBAS% = FECOSTO%(CIK%(KKI%))
            FEACT% = HOY(HOS$)
            COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
        End If
        '
        ' CONVERSION DE UNIDADES
        UNXX$ = UNIMED$(CIK%(KKI%))
        USON = CANT(KKI%)
        Call CONVERUNID(USON, UNXX$)
        '
        Z$ = Space$(256)
        Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
        DESCRI$ = DESCRIT0$(CIK%(KKI%))
        Call REPLA(Z$, DESCRI$, 17, 23)
        Call REPLA(Z$, DESCRI$, 129, 40)
        ' Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
        Call REPLA(Z$, CSTRING$(MKS$(USON), 3, 10, 3, 2), 40, 10)
        ' Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
        Call REPLA(Z$, UNXX$, 52, 4)
        Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
        Call REPLA(Z$, DEMONE$, 179, 10)
        Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
        UNIME$ = UCase$(UNIMED$(CIK%(KKI%)))
        If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
             If KKI% = 1 Then
                  SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                Else
                  SUPIREC = SUPIREC + CANT(KKI%) * COSUN
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 65, 10)
             SUPIE = SUPIE + CANT(KKI%) * COSUN
          Else
             If KKI% = 1 Then
                  SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                Else
                  SUMPREC = SUMPREC + COSUN * CANT(KKI%)
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 55, 10)
             SUPRI = SUPRI + CANT(KKI%) * COSUN
        End If
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
        '
      Else
        '
        If COSUNI(CIK%(KKI%)) > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
             If PROVINT%(PHB%) = 0 Then
               MONEI% = MONECOSTO%(CIK%(KKI%))
               DEMONE$ = DEMONECO$(MONEI%)
               COSUDOL = COSUNI(CIK%(KKI%))
               If MONEI% > 1 Then
                   COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                 Else
                   FEBAS% = FECOSTO%(CIK%(KKI%))
                   FEACT% = HOY(HOS$)
                   COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
               End If
               '
               Z$ = Space$(256)
               Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
               DESCRI$ = DESCRIT0$(CIK%(KKI%))
               Call REPLA(Z$, DESCRI$, 17, 23)
               Call REPLA(Z$, DESCRI$, 129, 40)
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
               Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
               Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
               Call REPLA(Z$, DEMONE$, 179, 10)
               Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
               If KKI% = 1 Then
                    SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                  Else
                    SUSEREC = SUSEREC + CANT(KKI%) * COSUN
               End If
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 75, 10)
               SUSER = SUSER + CANT(KKI%) * COSUN
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
             End If
           End If
        End If
        '
     End If
19 Next KKI%

   A1 = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   A2 = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   A3 = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   A4 = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   A5 = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   A6 = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(A1)
   SUMA1 = SUMA1 + Val(A2)
   SUMA1 = SUMA1 + Val(A3)
   SUMA1 = SUMA1 + Val(A4)
   SUMA1 = SUMA1 + Val(A5)
   SUMA1 = SUMA1 + Val(A6)
   SUCOS = SUMA1

   '
'   Static DP%(32767)
'   Static FESTRU#, NVI%, NMM%
'   '
'   SUCOS = 0
'   CICON% = CI%
'   If CICON% < 1 Or CICON% > NUMAS% Then GoTo 90
'   '
'   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
'   '
''   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
''   If FESTRUC# <> FESTRU# Then
''     '
''     NVI% = NUINV%
''     NMM% = NUMAS%
''     '
''     FESTRU# = FESTRUC#
''     FIN& = ULTREG&("ESTRUNUE")
''     For I& = 1 To FIN&
''       Call TRACE(20, I&, FIN&)
''       R0$ = REGLEIDO$("ESTRUNUE", I&)
''       CI1% = CVI(Left$(R0$, 2))
''       If CI1% > 0 Then
''         If CI1% <= NMM% Then
''           If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
''             DP%(CI1%) = 1
''           End If
''         End If
''       End If
''     Next I&
''   End If
'   '
'   Call EXPLOMUL(CICON%)
'   '
'   FIN& = CONTAMUL
'   For KKI% = 1 To CONTAMUL
'     COSUDOL = COSUNI(CIK%(KKI%))
'     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
'        '
'        MONEI% = MONECOSTO%(CIK%(KKI%))
'        COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
'        SUCOS = SUCOS + CANT(KKI%) * COSUN
'        '
'      Else
'        '
'        If COSUDOL > 0 Then
'          PHB% = PROVHABI%(CIK%(KKI%))
'          If PHB% > 0 Then
'            If PROVINT%(PHB%) = 0 Then
'              MONEI% = MONECOSTO%(CIK%(KKI%))
'              COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
'              SUCOS = SUCOS + CANT(KKI%) * COSUN
'            End If
'          End If
'        End If
'        '
'     End If
'   Next KKI%
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
   If STOMIN < 0 Then STOMIN = 100 * Abs(STOMIN)
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

Function LINDIVEN(CI%)
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then LINDIVEN = CDbl(MasterTemp!LINDIVEN)
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
       If EXISTE%(LUGRA$ + CODDX$ + ".JPG") = 0 Then
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
            '
            Call REPLA(RMCDX$, NUPLANO$(CICDX%), 181, 20) ' CODIGO ALIAS
            Call REPLA(RMCDX$, CODINTERNA$(CICDX%), 201, 48) ' CODIFICACION INTERNA
            '
            If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
              Call REPLA(RMCDX$, MKS$(99999999999#), 111, 4)
            End If
            OBXXX$ = OBJPLA$("MUESTRADATCOD", RMCDX$)
       Else
            '
            '\\\OT-1069-FL-24/11/08/// Carga los datos en el formulario FORDACOD07 en caso que el
            'articulo tenga asignado una imagen, sino carga el obejeto "MUESTRADATCOD"
            FORDACOD07.Label2.Caption = CODDX$
            FORDACOD07.Show 1
            '\\\OT-1069-FL-FIN///
       End If
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
               If UCase$(Mid$(X$, 62, 2) = "NO") Then
               DISTO%(U&) = (-1)
               End If
            End If
         End If
         '
         If U& = 1 Then
            ' ACTUALIZA LOS REGISTROS DE MOVIMIENTOS EN DEPOSITO 0 (CERO)
            CONDI$ = "COD_INTE > 0 AND DEPOMOVI = 0"
            QTR& = CantRegistros("MOVISTO", CONDI$)
            If QTR& > 0 Then
               Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", CONDI$, COSU%(1), RAFE&, "NOMESS")
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
   Static MIRS As ADODB.Recordset
   
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
'''   STAC = 0
'''   STRSQL = "Select sum(CantIngre) AS SUINGRE, SUM(CantSalid) as SUSALI from movisto WITH(NOLOCK) where cod_inte = " & CI%
'''   '
'''   Set MIRS = New ADODB.Recordset
'''   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
'''   If Not (MIRS.EOF And MIRS.BOF) Then   ' IsNull(mirs!SALDO) Then
'''      STAC = MIRS!SUINGRE - MIRS!SUSALI
'''   End If
'''   MIRS.Close
   'Call CARDEPOS
   CONDI1$ = ""
   For UI% = 1 To NUSU%
    If DISTO%(UI%) < 0 Then
       CONDI1$ = CONDI1$ + " AND DepoMovi <> " & COSU%(UI%)
    End If
   Next UI%
   '
   CONDI$ = "Cod_inte = " & CI%
   If CONDI1$ <> "" Then CONDI$ = CONDI$ & " AND (" & Mid$(CONDI1$, 5) & ")"
   'CONDI$ = "cod_inte = " & CI%
   SUINGRE = XVALO(VALOBADA("MOVISTO", "sum(CantIngre)", CONDI$, "NOMESS"))
   SUSALI = XVALO(VALOBADA("MOVISTO", "sum(CantSalid)", CONDI$, "NOMESS"))
   STACT = SUINGRE - SUSALI
   'STOCKACT = XVALO(STACT) - STONODIS(CI%)  'No considerar depositos no disponibles
   STOCKACT = XVALO(STACT)
   '
End Function
'
Function STONODIS(CI%)
   '
   Static CTX%, CONDI$
   If CTX% < 1 Then
      Call CARDEPOS
      CONDI$ = ""
      For UI% = 1 To NUSU%
        If DISTO%(UI%) < 0 Then
           CONDI$ = CONDI$ + " OR DepoMovi = " & COSU%(UI%)
        End If
      Next UI%
   End If
   '
   strsql$ = "Cod_inte = " & CI% & " AND (" & Mid$(CONDI$, 4) & ")"
   SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", strsql$, "NOMESS"))
   SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", strsql$, "NOMESS"))
   '
   STONODIS = SUINGRE - SUSALI ' STAC
   '
End Function
'
Function STODEPOS(CI%, DEPO%)
   '
   STAC = 0
   If DEPO% < 1 Then
      STAC = STOCKACT(CI%)
      GoTo 99
   End If
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   strsql$ = "Cod_inte = " & CI%
   If DEPO% > 1 Then
        strsql$ = strsql$ + " And DepoMovi= " & DEPO%
      Else
        strsql$ = strsql$ + " And DepoMovi <= " & DEPO%
   End If
   '
   SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", strsql$, "NOMESS"))
   SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", strsql$, "NOMESS"))
   STAC = SUINGRE - SUSALI ' STAC
   '
99 STODEPOS = STAC
   '
End Function
'
Function STODEPOS_NUEVO(CI%, DEPO%)
   '
   STAC = 0
   If DEPO% < 1 Then
      STAC = STOCKACT(CI%)
      GoTo 99
   End If
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   strsql$ = "Cod_inte = " & CI%
   If DEPO% > 1 Then
        strsql$ = strsql$ + " And DepoMovi= " & DEPO%
      Else
        strsql$ = strsql$ + " And DepoMovi <= " & DEPO%
   End If
   '
   SQLX$ = "SELECT SUM(CANTINGRE) as SUINGRE, SUM(CANTSALID) AS SUSALI FROM MOVISTO WITH(NOLOCK) WHERE " & strsql$
   'SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", STRSQL$, "NOMESS"))
   'SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", STRSQL$, "NOMESS"))
   Dim rst As ADODB.Recordset
   Set rst = READSET(SQLX$)
   If Not (rst.BOF And rst.EOF) Then
        STAC = XVALO(rst!SUINGRE) - XVALO(rst!SUSALI) ' STAC
   End If
   rst.Close
   Set rst = Nothing
   '
99 STODEPOS_NUEVO = STAC
   '
End Function

Function STODEPFECH(CI%, DEPO%, FECH%, Optional LOTE$)
   '
   ' DEVUELVE LA EXISTENCIA EN EL DEPÓSITO DEPO% A LA FECHA FECH% - 04/04/07 (OT-07-0141)
   ' SI DEPO% = 0 DEVUELVE EXISTENCIA TOTAL DEL ITEM A ESA FECHA
   If CI% > 0 Then
      SCMA = Screen.MousePointer
      Screen.MousePointer = 11
      '
      CARRAS# = 0
      HASDAT = FETEXCOR1$(FECH%)
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT sum(CantIngre) AS SUINGRE, SUM(Cantsalid) as SUSALI FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      SQLX$ = SQLX$ + "AND FechMov <= '" & HASDAT & "'"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      
      Dim MoviTemp As ADODB.Recordset
      Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             CARRAS# = XVALO(!SUINGRE) - XVALO(!SUSALI) ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing
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
35  Call ABRESERVA
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    'SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
    '
'    Dim ReserTemp As ADODB.Recordset
'    Set ReserTemp = New ADODB.Recordset
'    ReserTemp.CursorType = adOpenDynamic
'    ReserTemp.LockType = adLockReadOnly
'25  On Error Resume Next
'    ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
    Set ReserTemp = READSET(SQLX$)
    '
    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            SARE = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If SARE < 0 Then SARE = 0
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
             SARE = SARE / 1000
          End If

    End With
    '
    STOCKRES = SARE
    '
End Function
Function STOCKRES_POR_OFABRICACION(NORFA$, CI1%)
    '
    Call ABRESERVA
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI1%
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    Set ReserTemp = READSET(SQLX$)
    '
    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            SARE = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If SARE < 0 Then SARE = 0
          If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
             SARE = SARE / 1000
          End If

    End With
    '
    STOCKRES_POR_OFABRICACION = SARE
    '
End Function

Function STOCK_RESERV_SIN_ARMADOR(CIXI%)

        'DEVUELVE EL STOCK RESERVADO CONSIDERANDO SOLO LAS ORDENES DE FABRICACION QUE 'NO' ESTEN ASIGNADAS A ARMADORES
        '
        RESERVADO = 0
        SQLX$ = "SELECT SUM(RESERVA.CANTRES)AS CANTRES  , SUM(RESERVA.CANTCONS) AS CANTCONS "
        SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) INNER JOIN ORDEFABR WITH(NOLOCK)"
        SQLX$ = SQLX$ + " ON RESERVA.IDSUBOR = ORDEFABR.IDSUBOR"
        SQLX$ = SQLX$ + " Where (OrdeFabr.ARMADOR < 1 OR OrdeFabr.ARMADOR IS NULL)"
        SQLX$ = SQLX$ + " AND RESERVA.COD_INTE = " & CIXI%
        SQLX$ = SQLX$ + " AND ORDEFABR.STATUORF < 3  "
        SQLX$ = SQLX$ + " AND RESERVA.CantCons < (RESERVA.CantRes - 0.005)"
        SQLX$ = SQLX$ + " AND RESERVA.CodiEmpr = " & CodiEmp%
        SQLX$ = SQLX$ + " AND ORDEFABR.CodiEmpr = " & CodiEmp%
        SQLX$ = SQLX$ + " GROUP BY RESERVA.COD_INTE"
        '
        Dim RESTMP As ADODB.Recordset
        Set RESTMP = FLEXCONN.Execute(SQLX$) ' SIN FILTSQL POR QUE INNER JOIN NO CORRERSPONDE
        With RESTMP
            If Not (RESTMP.EOF And RESTMP.BOF) Then
              RESERVADO = XVALO(!CANTRES) - XVALO(!CANTCONS)
            End If
            If RESERVADO < 0 Then RESERVADO = 0
            If TRIM$(UCase$(UNIMED$(CIXI%))) = "GK" Then
               RESERVADO = RESERVADO / 1000
            End If
        End With
        RESTMP.Close
        Set RESTMP = Nothing
        '
        STOCK_RESERV_SIN_ARMADOR = RESERVADO
        '
End Function

'
Function STORESOF(CI%, NORFAI&)
    '
    NBRS$ = TRIM$(Str$(NORFAI&))
    While Len(NBRS$) < 6
      NBRS$ = "0" + NBRS$
    Wend
    NBRS$ = "OF-" + NBRS$ + "-%"
    '
10  ASII = 0
    '
35  Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS  FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "AND IdSubOr LIKE '" & NBRS$ & "' "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    '
    'Dim ReserTemp As ADODB.Recordset
    'Set ReserTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    Set ReserTemp = READSET(SQLX$)

    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            ASII = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If ASII < 0 Then ASII = 0
    End With
    If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
      ASII = ASII / 1000
    End If
    Set ReserTemp = Nothing
    '
    STORESOF = ASII
    '
End Function
'
Function STOCKDIS(CI%)
    '
    AA1 = STOCKACT(CI%)
    aa2 = STOCKASI(CI%)
    aa3 = STOCKRES(CI%)
    ' AA4 = STONODIS(CI%)   ' SUPRIMIDO PORQUE YA SE CONSIDERA EN STOCKACT
    
    STOCKDIS = AA1 - aa2 - aa3 - AA4 ' STOCKACT(CI%) - STOCKASI(CI%) - STOCKRES(CI%) - STONODIS(CI%)
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
Function STORFAB(CI%)
   '
   STORF = 0
   SQLX$ = "SELECT SUM(CANPROY) AS SUCANP, SUM(CANAPRO)AS SUCANAPRO FROM ORDEFABR "
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI% & ""
   SQLX$ = SQLX$ + " AND CANAPRO < CANPROY  "
   SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
   Dim OFFFABRIC As ADODB.Recordset
   Set OFFFABRIC = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With OFFFABRIC
    If Not (OFFFABRIC.EOF And OFFFABRIC.BOF) Then
      STORF = XVALO(!SUCANP) - XVALO(!SUCANAPRO)
    End If
    If STORF < 0 Then STORF = 0
   End With
   OFFFABRIC.Close
   STORFAB = STORF
   '
End Function

Function STOCOPEN(CI%)
     '
'     CAPEN = 0
   STOCOPEN = 0
   SQLX$ = "SELECT SUM(CANT_OCOM) AS SUOCOM, SUM(CANT_REC) AS SURECI FROM OCOMDETA"
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI% & ""
   SQLX$ = SQLX$ + " AND CANT_REC < CANT_OCOM "
   SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + " AND N_OrServ < 1 "
   Dim OCOMMMDETA As ADODB.Recordset
   Set OCOMMMDETA = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With OCOMMMDETA
    If Not (OCOMMMDETA.EOF And OCOMMMDETA.BOF) Then
      STOCOPEN = XVALO(!SUOCOM) - XVALO(!SURECI)
    End If
    If STOCOPEN < 0 Then STOCOPEN = 0
   End With
   OCOMMMDETA.Close
   Set OCOMMMDETA = Nothing

'18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
'     For UK& = 1 To Len(RFF$) Step 4
'        U& = CVS(Mid$(RFF$, UK&, 4))
'        X$ = REGLEIDO$("OCOMDETA", U&)
'        CIIX% = CVI(Mid$(X$, 9, 2))
'        STAT% = Asc(Mid$(X$, 124, 1))
'        If STAT% < 9 Then      ' no esta anulada
'          NUORSE& = CVS(Mid$(X$, 117, 4))
'          If NUORSE& < 1 Then
'            COEFCAN = CVS(Mid$(X$, 63, 4))
'            If COEFCAN < 0.005 Then COEFCAN = 1
'            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
'            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
'            CAPEN = CACOM - CAREC
'          End If
'        End If
'     Next UK&
'     '
'     If CAPEN < 0 Then CAPEN = 0
'     STOCOPEN = CAPEN

End Function

Function STOLOTES(CI%)
    '
    STLOT = 0
    '
    If CI% > 0 Then
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT * FROM LOTINGRE "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
      'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
      Dim LOTINGTEMP As ADODB.Recordset
      Set LOTINGTEMP = New ADODB.Recordset
26    On Error Resume Next
      LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
      End If
      On Error GoTo 0
      '
      With LOTINGTEMP
25      On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            STLOT = STLOT + XVALO(!CANSALDO)
          .MoveNext
          Loop
        End If
        On Error GoTo 0
      End With
      LOTINGTEMP.Close
      Set LOTINGTEMP = Nothing
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

Function UBISTO1$(CI%)
    '
20  SQLX$ = "SELECT DISTINCT UBICACION FROM UBICALOT WITH(NOLOCK) WHERE CODIGO=" & CI% & " AND CANTIDAD > 0.05"
    Set rst = READSET(SQLX$)
    '
    UB$ = ""
    With rst
        Do While Not .EOF
            UB$ = UB$ + SAFETEXT$(!UBICACION) & "; "
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    '
    UBISTO1$ = UB$
    '
End Function

'
Function GRABAUBICACION(CIXI%, CAN#, UBICA$, LOTE$)
    '
    UBICATMP$ = UBICA$
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        SE$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        ES$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        FA$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% < 1 Then
        CO$ = TRIM$(Mid$(UBICATMP$, 1))
    End If
    '
    If SE$ = "" Or ES$ = "" Or FA$ = "" Or CO$ = "" Then Exit Function
    '
100 ABUS$ = RANDSTRING$(-10)
    CONDI$ = "LOTE = '" & ABUS$ & "'"
    If CantRegistros("UBICALOT", CONDI$) > 0 Then GoTo 100
    '
    Call CreaRegistro("UBICALOT", "LOTE", ABUS$, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "SECTOR", CONDI$, SE$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "ESTANTERIA", CONDI$, ES$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "FILA", CONDI$, FA$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "COLUMNA", CONDI$, CO$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "CODIGO", CONDI$, CIXI%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "CANTIDAD", CONDI$, CAN#, REGAF&, "NOMESS")
        UBICACION$ = SE$ & ES$ & FA$ & CO$
    Call ACTUREGISTRO("UBICALOT", "UBICACION", CONDI$, UBICACION$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "LOTE", CONDI$, LOTE$, REGAF&, "NOMESS")
    '
End Function

Function UBISTONUE$(CIXI%, CAN#)
    '
    TX$ = ""
    If VOLUMENCOD#(CIXI%) > 0.00005 Then
        ' UBICACION LIBRE DENTRO DE LAS ASIGNADAS AL CODIGO
        TX$ = UBICACIONLIBRE$(CIXI%, CAN#)
        
        ' UBICACION OCUPADA DE LAS ASIGNADAS AL CODIGO, PERO CON ESPACIO SUFICIENTE
        If TX$ = "" Then TX$ = UBICACIONOCUPADA$(CIXI%, CAN#)
        
        ' UBICACION PERTENECENTE A LA FAMILIA CON ESPACIO SUFICIENTE. SE CONSIDERAN TAMBIEN A LOS
        ' PULMONES QUE LE FUERON ASIGNADO UNA FAMILIA
        If TX$ = "" Then TX$ = UBICACIONFAMILIA$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE A LA ESTANTERIA DONDE ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONDEESTANTERIA$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE AL SECTOR DONDE ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONDELSECTOR$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE A UN SECTOR DONDE NO ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONOTROSECTOR$(CIXI%, CAN#)
        
        'If TX$ = "" Then TX$ = UBICACIONMANUL$(CIXI%, CAN#)
    End If
    UBISTONUE$ = TX$

End Function

Function UBICACIONLIBRE$(CIXI%, CAN#)
    '
    UBICACIONLIBRE$ = ""
    '
    ' VERIFICA SI EL CODIGO ESTA ASIGNADO A ALGUNA UBICACION
        SQLX$ = "SELECT * FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI% & " ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set rst = READSET(SQLX$)
        With rst
            Do While Not .EOF
                ' BUSCA SI LA UBICACION (SECTOR+ESTANTE+FILA+COLUMNA) ESTA VACIO.
                ' ESTO SIGNIFICA QUE NO HAYAN REGISTROS EN UBICALOT
                SE$ = SAFETEXT$(!SECTOR)
                ES$ = SAFETEXT$(!ESTANTERIA)
                FI$ = SAFETEXT$(!FILA)
                CO$ = SAFETEXT$(!COLUMNA)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
                REG0& = CantRegistros("UBICALOT", CONDI$)
                If REG0& < 1 Then
                    ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                    LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                    M3UBICA# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$)) ' VOLUMEN DE LA UBICACION
                    If M3UBICA# > M3CAJAS# Then
                        UBICACIONLIBRE$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$ ' ASIGNA UBICACION
                        Exit Do
                    End If
                End If
                .MoveNext
            Loop
        End With
        On Error Resume Next
        rst.Close
        Set rst = Nothing
        On Error GoTo 0
    '
End Function

Function UBICACIONOCUPADA$(CIXI%, CAN#)
    '
    UBICACIONOCUPADA$ = ""
    '
    ' VERIFICA SI EL CODIGO ESTA ASIGNADO A ALGUNA UBICACION
    SQLX$ = "SELECT * FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI% & " ORDER BY SECTOR, ESTANTERIA, FILA, COLUMNA ASC"
    Set rst = READSET(SQLX$)
    With rst
        Do While Not .EOF
            M3UBICA# = 0
            ' CALCULA EL VOLUMEN TOTAL OCUPADO DE LA UBICACION
            SE$ = SAFETEXT$(!SECTOR)
            ES$ = SAFETEXT$(!ESTANTERIA)
            FI$ = SAFETEXT$(!FILA)
            CO$ = SAFETEXT$(!COLUMNA)
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!CODIGO)
                    CANTI# = XVALO(!Cantidad)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                UBICACIONOCUPADA$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
        '
End Function

Function UBICACIONFAMILIA$(CIXI%, CAN#)
    '
    UBICACIONFAMILIA$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE A LA FAMILIA DEL CODIGO
    FAMI$ = SAFETEXT$(VALOBADA("MASTER", "FAMILIA", "CODINT=" & CIXI%))
    SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE FAMILIA='" & FAMI$ & "' ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    Set rst = READSET(SQLX$)
    With rst
        Do While Not .EOF
            M3UBICA# = 0
            ' CALCULA EL VOLUMEN TOTAL OCUPADO DE LA UBICACION
            SE$ = TRIM$(SAFETEXT$(!SECTOR))
            ES$ = TRIM$(SAFETEXT$(!ESTANTERIA))
            FI$ = TRIM$(SAFETEXT$(!FILA))
            CO$ = TRIM$(SAFETEXT$(!COLUMNA))
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!CODIGO)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    CANTI# = XVALO(!Cantidad)
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                UBICACIONFAMILIA$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
                Exit Do
            End If
20          .MoveNext
        Loop
        On Error Resume Next
        RSTLOTE.Close
        Set RSTLOTE = Nothing
        On Error GoTo 0
    End With
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0

End Function

Function PULMONDEESTANTERIA$(CIXI%, CAN#)
    '
    PULMONDEESTANTERIA$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE AL CODIGO
    SQLX$ = "SELECT DISTINCT SECTOR, ESTANTERIA FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    Do While Not RSTSEC.EOF
        ' OBTIENE LOS PULMONES CORRESPONDIENTES AL SECTOR Y ESTANTERIA QUE LE FUE ASIGNADA
        SE$ = SAFETEXT$(RSTSEC!SECTOR)
        ES$ = SAFETEXT$(RSTSEC!ESTANTERIA)
        SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE SECTOR='" & SE$ & "' AND ESTANTERIA='" & ES$ & "' AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set rst = READSET(SQLX$)
        With rst
            Do While Not .EOF
                M3UBICA# = 0
                ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
                FI$ = TRIM$(SAFETEXT$(!FILA))
                CO$ = TRIM$(SAFETEXT$(!COLUMNA))
                ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
                VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
                SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
                Set RSTLOTE = READSET(SQLX$)
                With RSTLOTE
                    Do While Not .EOF
                        CIXX% = XVALO(!CODIGO)
                        CANTI# = XVALO(!Cantidad)
                        LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                        VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                        '
                        M3UBICA# = M3UBICA# + VOLUMENLOTE#
                        .MoveNext
                    Loop
                End With
                ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
                LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                    PULMONDEESTANTERIA$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
                    Exit Do
                End If
20              .MoveNext
            Loop
        End With
        On Error Resume Next
        rst.Close
        Set rst = Nothing
        On Error GoTo 0
        RSTSEC.MoveNext
    Loop
    On Error Resume Next
    RSTSEC.Close
    Set rst = Nothing
    On Error GoTo 0
    '
End Function

Function PULMONDELSECTOR$(CIXI%, CAN#)
    '
    PULMONDELSECTOR$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE A LA FAMILIA DEL CODIGO
    SQLX$ = "SELECT DISTINCT SECTOR FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    Do While Not RSTSEC.EOF
        SECTOR$ = SAFETEXT$(RSTSEC!SECTOR)
        SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE SECTOR='" & SECTOR$ & "' AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set rst = READSET(SQLX$)
        With rst
            Do While Not .EOF
                M3UBICA# = 0
                ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
                SE$ = TRIM$(SAFETEXT$(!SECTOR))
                ES$ = TRIM$(SAFETEXT$(!ESTANTERIA))
                FI$ = TRIM$(SAFETEXT$(!FILA))
                CO$ = TRIM$(SAFETEXT$(!COLUMNA))
                ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
                VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
                SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
                Set RSTLOTE = READSET(SQLX$)
                With RSTLOTE
                    Do While Not .EOF
                        CIXX% = XVALO(!CODIGO)
                        CANTI# = XVALO(!Cantidad)
                        LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                        VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                        '
                        M3UBICA# = M3UBICA# + VOLUMENLOTE#
                        .MoveNext
                    Loop
                End With
                ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
                LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                    PULMONDELSECTOR$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
                    Exit Do
                End If
20              .MoveNext
            Loop
        End With
        On Error Resume Next
        rst.Close
        Set rst = Nothing
        On Error GoTo 0
        RSTSEC.MoveNext
    Loop
    On Error Resume Next
    RSTSEC.Close
    Set rst = Nothing
    On Error GoTo 0
    '
End Function

Function PULMONOTROSECTOR$(CIXI%, CAN#)
    '
    PULMONOTROSECTOR$ = ""
    '
    ' OBTENGO LOS SECTORES DONDE ESTA EL PRODUCTO PARA OBVIARLOS
    SQLX$ = "SELECT DISTINCT SECTOR FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    I% = 0: CONDI$ = ""
    Do While Not RSTSEC.EOF
        '
        SECTOR$ = TRIM$(SAFETEXT$(RSTSEC!SECTOR))
        I% = I% + 1
        If I% = 1 Then
            CONDI$ = CONDI$ & " SECTOR<>'" & SECTOR$ & "'"
        Else
            CONDI$ = CONDI$ & " AND SECTOR<>'" & SECTOR$ & "'"
        End If
        '
        RSTSEC.MoveNext
    Loop
    '
    SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE " & CONDI$ & " AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    If CONDI$ = "" Then SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    Set rst = READSET(SQLX$)
    With rst
        Do While Not .EOF
            M3UBICA# = 0
            ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
            SE$ = TRIM$(SAFETEXT$(!SECTOR))
            ES$ = TRIM$(SAFETEXT$(!ESTANTERIA))
            FI$ = TRIM$(SAFETEXT$(!FILA))
            CO$ = TRIM$(SAFETEXT$(!COLUMNA))
            ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "' AND (FAMILIA='' OR FAMILIA IS NULL)"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!CODIGO)
                    CANTI# = XVALO(!Cantidad)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                PULMONOTROSECTOR$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
                Exit Do
            End If
20          .MoveNext
        Loop
        On Error Resume Next
        RSTLOTE.Close
        Set RSTLOTE = Nothing
        On Error GoTo 0
    End With
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    '
End Function

Function CAMBIAUBICACION$()
        '
        Call SELECHO("MASTER/Maestro de Articulos")
        COD$ = TRIM$(VALACT1$("MASTER"))
        If COD$ <> "" Then
            CIXI% = CODINT%(COD$)
            LOTINDI# = LINDIVEN(CIXI%)
            If LOTINDI# < 0.00005 Then
                Call COMUNI("Imposible Continuar. Falta Lote Indivisible de Ventas")
                Exit Function
            End If
            '
            Call ARMALILOT(CIXI%, , , , 1) ' PRESENTA SOLO LOTES CON SALDO
            Call SELECHO("!LILOTSE")
            '
            LOTE$ = TRIM$(VALACT1$("!LILOTSE"))
            LOTERESTO$ = TRIM$(VALACT2$("!LILOTSE"))
            If LOTE$ = "" Then Exit Function
            '
            FAMI$ = SAFETEXT$(VALOBADA("MASTER", "FAMILIA", "CODINT=" & CIXI%))
            SECTOR$ = TRIM$(SAFETEXT$(VALOBADA("UBICAVOL", "SECTOR", "FAMILIA='" & FAMI$ & "'")))
            '
            If SECTOR$ <> "" Then
                Call CARGALISEL("UBIMANU", "UBICAVOL", "SECTOR/A4;ESTANTERIA/A4;FILA/A4;COLUMNA/A4;FAMILIA/A4", "SECTOR='" & SECTOR$ & "'")
            Else
                Call CARGALISEL("UBIMANU", "UBICAVOL", "SECTOR/A4;ESTANTERIA/A4;FILA/A4;COLUMNA/A4;FAMILIA/A4", "SECTOR<>''")
            End If
10          Call SELECHO("UBIMANU/Ubicaciones del Sector '" & SECTOR$ & "'")
            
            SE$ = TRIM$(VALACT1$("UBIMANU"))
            If SE$ = "" Then Exit Function
            
            TXTX$ = VALACT2$("UBIMANU")
            ES$ = TRIM$(Mid$(TXTX$, 1, 4))
            FI$ = TRIM$(Mid$(TXTX$, 6, 4))
            CO$ = TRIM$(Mid$(TXTX$, 11, 4))
            '
            CAN# = XVALO(VALOBADA("UBICALOT", "CANTIDAD", "LOTE='" & LOTE$ & "'"))
            UBAC$ = TRIM$(SAFETEXT$(VALOBADA("UBICALOT", "UBICACION", "LOTE='" & LOTE$ & "'")))
            '
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!CODIGO)
                    CANTI# = XVALO(!Cantidad)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 1
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXI%)
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                REUBICA$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
            End If
            If REUBICA$ = "" Then
                OPX% = COMALTER%("La Ubicación Elegida No Posee Espacio Suficiente.\\Elegir Otra Ubicacion\Cancelar")
                If OPX% < 1 Or OPX% > 2 Then Exit Function
                If OPX% = 1 Then GoTo 10
            End If
            OPX% = COMALTER%("¿Esta Seguro Cambiar la Ubicacion del Lote '" & LOTE$ & "'\De '" & UBAC$ & "' Por '" & REUBICA$ & "'?\\No, Conservar\Si, Cambiar")
            If OPX% = 2 Then
                Call BORRAREGISTROS("UBICALOT", "LOTE='" & LOTE$ & "'", REGAF&)
                Call GRABAUBICACION(CIXI%, CAN#, REUBICA$, LOTE$)
                Call COMUNI("Cambio de Ubicación Realizada")
            End If
            '
        End If
        '
End Function

Function SEPRODUC(CI%)
    '
20  SPROD = 0
    Call APERBASDAT("ORFAB")
    '
    SQLX$ = "SELECT NumeOrFa, SecciPro FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
    SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
    'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORFABTEMP As ADODB.Recordset
    Set ORFABTEMP = New ADODB.Recordset
25  On Error Resume Next
    ORFABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    On Error Resume Next
    With ORFABTEMP
      .MoveFirst
      If Err < 1 Then
        SPROD = XVALO(!SecciPro)
      End If
      On Error GoTo 0
    End With
    Set ORFABTEMP = Nothing
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

Function PRELISTA#(APREC$, CIIXXI%)
   '
   Call APERBASDAT("PEDCLI")
   '
   If UCase$(Left$(APREC$, 5)) = "LIPRE" Then
       NLIS% = XVALO(Mid$(APREC$, 6, 3))
     Else
       NLIS% = XVALO(APREC$)
   End If

   PRELISTA# = 0: VUELTA% = 0
        If CIIXXI% > 0 Then
          If CIIXXI% <= NUMAS% Then
            SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
            SQLX$ = SQLX$ + "WHERE Cod_Lista =" & NLIS% & " "
            SQLX$ = SQLX$ + "and Cod_pieza =" & CIIXXI% & ""
            Dim PRODUC As ADODB.Recordset
            Set PRODUC = FLEXCONN.Execute(FILTSQL$(SQLX$))
            IncluirIva% = XVALO(VALOBADA("LISTAPRE", "INCLUIVA", "Cod_Lista =" & NLIS% & " "))

            With PRODUC
               On Error Resume Next
               .MoveFirst
               FIN& = .RecordCount
               If Err < 1 Then
                  If IncluirIva% = 1 Then ' SI INCLUYE IVA CALCULA EL PRECIO SIN IVA
                    PRELISTA# = (XVALO(!Precio_Unitario)) / (1 + PORCEIVA(CIIXXI%) / 100)
                  Else
                    PRELISTA# = XVALO(!Precio_Unitario)
                  End If
               End If
               On Error GoTo 0
            End With
            Set PRODUC = Nothing
          End If
        End If
        
   If CONVERPRECIO% = 1 Then
     'NLIS% = XVALO(Mid$(APREC$, 6, 3))
     PRELISTA# = PRELISTA# * TICAMONE(MONELISTA%(NLIS%)) / TICAMONE(MONEMI%)
   End If
   '
End Function

Function PRELISTACT#(CIX%, TIMONE%, Optional SOLOLISVIG%)
    'SOLOLISVIG% PARAMETRO OPCIONAL 0:LISTA VIGENTE SI ES 0 TRAE COSTO REPOSICION
    '                               1:SOLO DEVUELVE LISTA VIGENTE
    'TIMONE% , PARAMETRO QUE DEVUELVE CON EL TIPO DE MONEDAD DE LA LISTA
    'SE CONSIDERA QUE SOLO HAY UNA LISTA VIGENTE.
    'SI SE QUIERE QUE TRAIGA LA LISTA VIGENTE QUE CONTENGA ESE ARTICULO VOY A TENER
    'QUE AGREGAR EN PUNTO 2 LA CONDICION TAMBIEN DEL   'AND COD_PIEZA = " & CI%'
    '
    CI% = CIX%
    '
'1.-TOMAR LA FECHA DEL DIA
    FECHAX = FETEXCOR1$(HOY(HO$))
    '----------------------------------------------------------------
    '
'2.-BUSCO LA LISTA VIGENTE
    'BUSCA QUE LA FECHA DE VIGENCIA SEA <= A LA FECHA DE HOY Y LO ORDENA DE MANERA DESCENDENTE
    'PARA QUE PRESENTE LA MAS CERCANA A HOY
    CONDI$ = "COD_LISTA > 0 AND VIGENCIA <= '" & FECHAX & "'"
    CONDI$ = CONDI$ + " ORDER BY VIGENCIA DESC"
    LISTA_ACTIVA% = XVALO(VALOBADA("LISTAPRE", "COD_LISTA", CONDI$, "NOMESS"))
    '------------------------------------------------------------------
    '
'3.-BUSCO EL CODIGO EN LA LISTA
    CONDI2$ = "COD_LISTA = " & LISTA_ACTIVA% & " AND COD_PIEZA = " & CI%
    VALORLISTA# = XVALO(VALOBADA("PREVENTA", "PRECIO_UNITARIO", CONDI2$))
    '-------------------------------------------------------------------
    '
'4.- SI ENCUENTRA EL CODIGO EN LA LISTA Y EL PRECIO ES > 0 SALE
     If VALORLISTA# > 0 Or SOLOLISVIG% > 0 Then 'SI TIENE PRECIO EN LISTA O SI ES 0 PERO SOLO DEVUELVE LISTA VIGENTE
          PRELISTACT# = VALORLISTA#
          TIMONE% = MONELISTA%(LISTA_ACTIVA%)
          If TIMONE% < 1 Then TIMONE% = 1
          Exit Function
     End If
    '----------------------------------------------------------------
    '
'6.-EN CASO QUE SEA 0 DEVUELVE EL COSTO DE REPOSICION,
    PRELISTACT# = COSUNI(CI%)
    TIMONE% = MONECOSTO%(CI%)
    If TIMONE% < 1 Then TIMONE% = 1
    '
End Function
Function MONELISTA%(ILISTA%)
    Call ABREPEDCLI
    ML% = XVALO(VALOBADA("LISTAPRE", "Moneda", "Cod_Lista = " & ILISTA%))
    If ML% < 1 Then ML% = 1
    MONELISTA% = ML%
End Function

Function DECILISTA%(ILISTA%)
    Call ABREPEDCLI
    ML% = XVALO(VALOBADA("LISTAPRE", "Decimales", "Cod_Lista = " & ILISTA%))
    If ML% < 2 Then ML% = 2
    If ML% > 5 Then ML% = 5
    DECILISTA% = ML%
End Function


Function DENOLISTA$(ILISTA%)
    Call ABREPEDCLI
    ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
    DNL$ = "Precios Concertados (Pesos)"
    If ILIS% > 0 Then
       DNL$ = TRIM$(SAFETEXT$(VALOBADA("LISTAPRE", "Denom_Lista", "Cod_Lista = " & ILISTA%)))
       MNL% = MONELISTA%(ILIS%)
       If MNL% = 2 Then
           DNL$ = DNL$ + " (Dólares)"
         Else
           DNL$ = DNL$ + " (Pesos)"
       End If
    End If
    DENOLISTA$ = DNL$
End Function
'
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
       'Master.FindFirst ("CODINT = " & COD1%)
       'If Master.NoMatch = False Then
       SQLX$ = "SELECT DepoEntra FROM MASTER "
       SQLX$ = SQLX$ + " WHERE CODINT = " & COD1% & ""
       Dim MASTTTER As New ADODB.Recordset
       Set MASTTTER = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MASTTTER
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
          DEPENART% = XVALO(!DEPOENTRA)
         End If
         On Error GoTo 0
       End With
       Set MASTTTER = Nothing
   End If
   '
End Function
'
Function DEPCONART%(COD2%) 'DEPOSITO POR OMIDION PARA CONSUMO DE O/F
   '
   Call APERBASDAT("MASTER")
   If COD2% > 0 Then
       'Master.FindFirst ("CODINT = " & COD2%)
       'If Master.NoMatch = False Then
       SQLX$ = "SELECT DepoCon FROM MASTER "
       SQLX$ = SQLX$ + " WHERE CODINT = " & COD2% & ""
       Dim MASTTTER As New ADODB.Recordset
       Set MASTTTER = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MASTTTER
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
          DEPCONART% = XVALO(!DEPOCON)
         End If
         On Error GoTo 0
       End With
       Set MASTTTER = Nothing
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
   CAIMINI% = 0
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
                 CAIMINI% = 1
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
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR AltOper IS NULL) "
       
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SECOPTEMP.MoveFirst
   With SECOPTEMP
     NHPF$ = SAFETEXT$(!HoProFab)
   End With
   Set SECOPTEMP = Nothing
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
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " OR AltOper IS NULL "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SECOPTEMP.MoveFirst
   With SECOPTEMP
     NPLX$ = SAFETEXT$(!CodMaq0)
   End With
   Set SECOPTEMP = Nothing
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
   SQLX$ = SQLX$ + "WHERE CodIConj= " & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR  AltOper IS NULL )"
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   
   CAIT% = 0
   'On Error Resume Next
   'SECOPTEMP.MoveLast
   'If Err Then GoTo 90
   'CAREC% = SECOPTEMP.RecordCount
   'If CAREC% < 1 Then GoTo 90
   '
   'SECOPTEMP.MoveFirst
   With SECOPTEMP
     If Not (.EOF And .BOF) Then
        CALOTES = 1
        CAOPS = XVALO(!CantiOps)
          If CAOPS < 0.01 Then CAOPS = 1
        CICLOHE = XVALO(!PieCiclo)
          If CICLOHE < 1 Then CICLOHE = 1
        MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
          If MODO% > 0 Then MINPREP = 0
        MINFIJO = 60 * !HorsFijo * CALOTES
        MINVARI = CANTID# * !MinuStan
        MINHERR = CANTID# * 60 * !HorsMHer / CICLOHE
        MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
        MINLIMP = 60 * !HorsLimp * CALOTES
        MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
        '
        TSTA = CAOPS * MINLOTE
     End If
     '
   End With
   Set SECOPTEMP = Nothing
   

   

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
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR AltOper is NULL)  "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With SECOPTEMP
     CALOTES = 1
     CICLOHE = XVALO(!PieCiclo)
       If CICLOHE < 1 Then CICLOHE = 1
     MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
       If MODO% > 0 Then MINPREP = 0
     MINFIJO = 60 * XVALO(!HorsFijo) * CALOTES
     MINVARI = CANTID# * XVALO(!MinuStan)
     MINHERR = CANTID# * 60 * XVALO(!HorsMHer) / CICLOHE
     MINDEMO = (MINFIJO + MINVARI + MINHERR) * XVALO(!PorDemor) / 100
     MINLIMP = 60 * XVALO(!HorsLimp) * CALOTES
     MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
     '
     TSTA = MINLOTE
     '
   End With
   Set SECOPTEMP = Nothing
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
   SQLX$ = "SELECT count(*) as CANOPE from SecOper"
   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + " AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   On Error Resume Next
   CNP% = SECOPTEMP!CANOPE    'RecordCount
   On Error GoTo 0
   Set SECOPTEMP = Nothing
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
   '
   CONDI$ = "CodIConj=" & CIXI% & " "
   CONDI$ = CONDI$ + " AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        CONDI$ = CONDI$ + " AND AltOper = " & ALTOP% & " "
      Else
        CONDI$ = CONDI$ + " AND (AltOper = " & ALTOP% & " OR AltOper IS NULL ) "
   End If
   CONDI$ = CONDI$ + " AND StatOper = 1"
   '
   Descr$ = VALOBADA("SecOper", "DescOper", CONDI$, "NOMESS")
   CAIT% = 0

   If TRIM$(Descr$) <> "" Then
      NPLX$ = Descr$
   End If
   '
   DESCRIOP$ = NPLX$
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
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   SECOPTEMP.MoveFirst
   If Err < 1 Then
      ULTX1% = XVALO(SECOPTEMP!NUMEOPER)
   End If
   On Error GoTo 0
   Set SECOPTEMP = Nothing
   '
90 ULTIMOP% = ULTX1%
   '
End Function
'
Function CANOPREP(CIXI%, NOPE%)  ' CANTIDAD DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 1
   'Set SecopTemp = Articulos.OpenRecordset("SELECT Nuoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT Nuoper_Prep FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & ""
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   '
   If CAREC% > 0 Then
     SECOPTEMP.MoveFirst
     If SECOPTEMP!NUOPER_PREP > 0 Then
       NPLY = SECOPTEMP!NUOPER_PREP
     End If
   End If
   Set SECOPTEMP = Nothing
   '
90 CANOPREP = NPLY
   '
End Function
'
Function CATOPREP%(CIXI%, NOPE%)  ' CATEGORIA DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 0
   'Set SecopTemp = Articulos.OpenRecordset("SELECT Catoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT Catoper_Prep FROM SecOper "
   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & ""
   SQLX$ = SQLX$ + " AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   CAIT% = 0
'   On Error Resume Next
'   SECOPTEMP.MoveLast
'   If Err Then GoTo 90
'   CAREC% = SECOPTEMP.RecordCount
'   '
'   If CAREC% > 0 Then
'     SECOPTEMP.MoveFirst
'     If SECOPTEMP!CATOPER_PREP > 0 Then
    If Not SECOPTEMP.EOF Then
       NPLY = XVALO(SECOPTEMP!CATOPER_PREP)
     End If
'   End If
   Set SECOPTEMP = Nothing
   '
90 CATOPREP% = NPLY
   '
End Function
'

Function OBSEROP$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   SQLX$ = "SELECT ObseOper FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   NPLX$ = SECOPTEMP!OBSEOPER
   On Error GoTo 0
   Set SECOPTEMP = Nothing
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
   'CARGA EL AJUSTE EN EL VECTOR PARA TODOS LOS MESES DE LA TABLA
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
   'LEE EL VALOR DEL MES DESDE
10 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEBAS% Then
       IAJ1# = IAJU#(KKI%)
       GoTo 20
     End If
   Next KKI%
   IAJ1# = IAJU#(CAPER%)
   '
   'LEE EL VALOR DEL MES HASTA
20 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEACT% Then
       IAJ2# = IAJU#(KKI%)
       '================================================================================================
       'SE AGREGA ESTAS 2 LINEAS PARA QUE CUANDO CORRE EL CALCULO SEGURAMENTE EL MES ACTUAL NO VA A TENER EL INDICE DE INFLACION.
       'ENTONCES COMPARA SI EL INDICE DEL MES ACTUAL ES < AL ANTERIOR SE QUEDA CON EL DEL MES ANTERIOR,
       'OCURRIA QUE EL MES ACTUAL TENIA 1 Y POR EJEMPLO EL ANTERIOR 1.8 PERO SE QUEDABA CON 1 Y NO ACTUALIZABA EL COSTO.
       IAJ3# = IAJU#(KKI% - 1)
       If IAJ2# < IAJ3# And IAJ2# = 1 Then IAJ2# = IAJ3# 'COMPARA QUE SEA = 1 POR QUE PUEDE SER MENOR EN CASO DE DEFLACION (EN ESE CASO VA A QUEDARSE CON EL ULTIMO)
       'FIN AGREGADO ====================================================================================
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
Function CODORIG$(CIXXX%, NCXXX)
   '
   ASELE$ = "REPLACOC"
   If NCXXX < 0 Then ASELE$ = "REPLACOP"
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
         If CVI(Mid$(XXYY$, 1, 2)) = Abs(NCXXX) Then
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
Function GRUCOFAB_V2%(CIJK%)
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
        ' SI EL ARTICULO TIENE PROVEEDOR HABITUAL Y A SU VEZ SECUENCIA DE OPERACIONES.
        ' SE CONSIDERA QUE SE FABRICA PERO EVENTUALMENTE SE COMPRA
        If CantRegistros("SECOPER", "CodIConj=" & CStr(CIJK%)) = 0 Then
          GRF%(CIJK%) = 1
        End If
      End If
    End If
  End If
  '
  GRUCOFAB_V2% = GRF%(CIJK%)
  '
End Function

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
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

10 COD1$ = "": COD2$ = "": CIKX1% = 0: CIKX2% = 0
   CI1% = PRICOD%
   CI2% = ULTCOD%
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
   'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With LOTINGTEMP
25   On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        FECREALOT% = CVINT%(!FECHALTA)
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
   '
   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With LOTINGTEMP
25   On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        SALNOCONS = XVALO(!CanAlta) - XVALO(!CanConsu)
     End If
     On Error GoTo 0
   End With
   Set LOTINGTEMP = Nothing
   If SALNOCONS < 0 Then SALNOCONS = 0
   '
End Function


Sub ALTACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER 'abre conexion
   '
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
10 OBX$ = OBJPLA$("ALTACODIGOSQL", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 24))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > LONCODI% Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que " & LONCODI% & " Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      '\\\OT-06-0263-WAR-05/07/06///
      CIX1% = CODINT%(CODINUE$)
      If CIX1% < 0 Then 'El Articulo esta de Baja
        RTA1% = COMALTER("Código de Baja. Puede Ahora Restaurar\ el Código de Baja.\\Restaurar\Continuar")
        If RTA1% = 1 Then 'Restaurar
          Call ABREMASTER 'abre conexion
          
          'With Master
          '  .FindFirst ("CODINT = " & Abs(CIX1%))
          '  If .NoMatch = False Then
          '    .Edit
          '    !Status = 1
          '    If Mid$(Master!CODIGO, 16, 1) = Chr$(96) Then
          '      Master!CODIGO = CODINUE$
          '    End If
          '    .Update
          'SQL
          strsql = ("Update Master set Status=1 where Codint=") & Abs(CIX1%)
          FLEXCONN.Execute strsql
          
              UPDATMASTER% = 1
              Call GENECOMAS
              OKX% = 1 'OK
              GoTo 99
            End If
          'End With
        
        End If
      'End If
      '\\\OT-06-0263-WAR-FIN///
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   GRCONTA% = CVI(Mid$(OBX$, 3, 2))
   If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
     GRCONTA% = 0
   End If

   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
   '
   TIMA% = Asc(Mid$(OBX$, 101, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   
  If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
    GRCONTA% = 0
   End If

   '
   Screen.MousePointer = 11
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER 'abreconexion
   '
ESAHP% = 0
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "AHP") > 0 Then
   ESAHP% = 1
End If

   Dim RMaster As ADODB.Recordset
   Set RMaster = New ADODB.Recordset
   RMaster.CursorType = adOpenDynamic
   RMaster.LockType = adLockPessimistic
   strsql = ("Select * from Master where codint=" & CODII%)
25 On Error Resume Next
   RMaster.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   If Not (RMaster.EOF And RMaster.BOF) Then 'si no esta vacio por lo tanto existe
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If

   With RMaster
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !CODIGO = CODINUE$
      !Descripcion = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      !Grucoven = GRCONTA%
      If ESAHP% > 0 Then
        !TRAZABLE = 1
      End If
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
   End With
   '
   RMaster.Close
   Set RMaster = Nothing
   UPDATMASTER% = 1
   '
   
   ' generar un registro para los saldos de stock
   'no se usa mas porque se remplaza por funcion stockact
'   Call ABRESTOCKS
'   Dim RSalStock As ADODB.Recordset
'   Set RSalStock = New ADODB.Recordset
'   RSalStock.CursorType = adOpenDynamic
'   RSalStock.LockType = adLockPessimistic
'
'   RSalStock.Open ("Select * from SalStock where CodiInte=" & CODII%), FLEXCONN, , , adCmdText
'
'   If (RSalStock.EOF And RSalStock.BOF) Then 'si esta vacio por lo tanto no existe
'     With RSalStock
'       .AddNew
'       !CodiInte = CODII%
'       !VeSalSto = String$(256, 0)
'       .Update
'       .Close
'     End With
'   End If
'
     '
99 Screen.MousePointer = 1
Set RMaster = Nothing
End Sub

Sub ALTARELACODIGO(VDEF$, OKX%)
   '
   Call ABREMASTER 'abre conexion
   '
   OBX$ = VDEF$
   OKX% = 0
   LONCODI% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 24))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 99
   End If
   '
   If Len(CODINUE$) > LONCODI% Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que " & LONCODI% & " Caracteres.")
      GoTo 99
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      CIX1% = CODINT%(CODINUE$)
      If CIX1% < 0 Then 'El Articulo esta de Baja
        RTA1% = COMALTER("Código de Baja. Puede Ahora Restaurar\ el Código de Baja.\\Restaurar\Continuar")
        If RTA1% = 1 Then 'Restaurar
          Call ABREMASTER 'abre conexion
          
          strsql = ("Update Master set Status=1 where Codint=") & Abs(CIX1%)
          FLEXCONN.Execute strsql
              Call GENECOMAS
              OKX% = 1 'OK
              GoTo 99
            End If
        End If
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 99
   End If
   GRCONTA% = 0
   CODIORIGINAL% = CVI(Mid$(OBX$, 67, 2))
   If CODIORIGINAL% < 1 Then
    If TRIM$(Mid$(OBX$, 3, 64)) = "" Then
       Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
       GoTo 99
    End If
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 3, 64))
   '
   TIMA% = 1
   '
   UNID$ = "U."
   '
   Screen.MousePointer = 11
   'SI TIENE  CODIGO DE RELACION RELACIONA EN LA TABLA CODICLIENTE Y SALE
   If CODIORIGINAL% > 0 Then
      '
      'VALIDA SI YA ESTA RELACIONADO
      CONDI$ = "Codigo_Cliente = '" & CODINUE$ & "' AND CODINT = " & CODIORIGINAL%
      EXISTECODNUE% = CantRegistros&("CODICLIENTE", CONDI$, "NOMESS")
      If EXISTECODNUE% > 0 Then
         Call COMUNI("Código yá Relacionado Con Item Seleccionado")
         GoTo 99
      End If
      Dim RS As New ADODB.Recordset
      SQLX$ = " SELECT * FROM CODICLIENTE "
      SQLX$ = SQLX$ + " WHERE CODINT = " & CODIORIGINAL%
26    On Error Resume Next
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
      End If
      On Error GoTo 0
      With RS
         .AddNew
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         !Codigo_Cliente = UCase$(CODINUE$)
         !Status = 1
         !CODINT = CODIORIGINAL%
         On Error Resume Next
         .Update
         OKX% = 2
         If Err Then
           Call MENSERR(24, "Relación de Código no Procesada Correctamente")
           OKX% = 0
         End If
         On Error GoTo 0
      End With
      RS.Close
      Set RS = Nothing
      GoTo 99
   End If
   'FIN RELACION CODIGO CLIENTE (SI ENTRO ACA SE VA )
   CODII% = NUMAS% + 1
   '
20 On Error Resume Next
   Call ABREMASTER 'abreconexion
   '
   Dim RMaster As ADODB.Recordset
   Set RMaster = New ADODB.Recordset
   RMaster.CursorType = adOpenKeyset
   RMaster.LockType = adLockPessimistic
   strsql = ("Select * from Master where codint=" & CODII%)
25 On Error Resume Next
   RMaster.Open FILTSQL$(strsql), FLEXCONN, , , adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   If Not (RMaster.EOF And RMaster.BOF) Then 'si no esta vacio por lo tanto existe
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 99
   End If

   With RMaster
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !CODIGO = CODINUE$
      !Descripcion = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      !Grucoven = GRCONTA%
      !TRAZABLE = 1
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
   End With
   '
   RMaster.Close
   Set RMaster = Nothing
   '
99 Screen.MousePointer = 1
On Error Resume Next
Set RMaster = Nothing
On Error GoTo 0
End Sub


Function ESTRAZABLE%(CI%)
   '
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then ESTRAZABLE% = CDbl(MasterTemp!TRAZABLE)
   '
End Function
Function MOLDE%(CI1%)
   '
   On Error Resume Next
   If STATUIT%(CI1%) <> 0 Then MOLDE% = CDbl(MasterTemp!MOLDE)
   '
End Function

Function MARCALOT$(CIXI%, IDLOT$)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT Marca FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          LTP1$ = SAFETEXT$(!MARCA)
       End If
    End With
    '
    Set LOTINGTEMP = Nothing
    MARCALOT$ = LTP1$
    '
End Function
         
Function LOTORIG$(CIXI%, IDLOT$)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT LoteCol_Prov FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          LTP1$ = SAFETEXT$(!LOTECOL_PROV)
       End If
    End With
    '
    Set LOTINGTEMP = Nothing
    LOTORIG$ = LTP1$
    '
End Function
'
Function PROXILOTE$(LOTEORI$)
    '
    NCR$ = TRIM$(LOTEORI$)
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1) ' VERIFICA LA EXISTENCIA DE LA BARRA IDENTIFICADORA DE SUBLOTE
    End If
    '
    Call APERBASDAT("STOCKS")
    For KKII% = 65 To 90
      NCR1$ = NCR$ + "/" + Chr$(KKII%) ' ARMA EL NUEVO SUBLOTE
      QTR& = CantRegistros("LOTINGRE", "IdenLote = '" & NCR1$ & "'") ' VERIFICA LA EXISTENCIA DEL SUBLOTE
      If QTR& < 1 Then Exit For
    Next KKII%
    ' OJO QUE SI SALE POR AQUÍ QUEDA EN /Z
    PROXILOTE$ = NCR1$
    '
End Function
'
Function UBILOTE$(CIXI%, IDLOT$, Optional MODO%)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT DepoUbic, Ubicfis FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    DPX% = 0: UBX$ = ""
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          DPX% = XVALO(!DepoUbic)
          UBX$ = SAFETEXT$(!UBICFIS)
       End If
    End With
    '
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
    DPXX$ = ""
    If MODO% < 1 Then
      If DPX% > 0 Then
        DPXX$ = ECOARCH$("TADEPOSI", Chr$(DPX%))
        If DPXX$ <> "" Then DPXX$ = DPXX$ + " - "
      End If
    End If
    '
    UBILOTE$ = DPXX$ + UBX$
    '
End Function
'
Function DEPOLOTE%(CIXI%, IDLOT$)
    '
    Call APERBASDAT("STOCKS")
    CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE ='" & IDLOT$ & "'"
    DEPOLOTE% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
    '
End Function
'
Function AUTONOMIASM$(CI%)
   'Leer dato en campo Autono_Proy
   PSM$ = ""
   CONDI$ = "CODINT = " & CI%
   PSM$ = TRIM$(VALOBADA("MASTER", "Autono_Proy", CONDI$, "NOMESS"))
   '
   If PSM$ <> "" Then
      AUTONOMIASM$ = PSM$
      GoTo 99
   End If
   '
   TIMAT% = TIMATE%(CI%)
   If TIMAT% > 0 Then ' FUNCION QUE DEVUELVE EL LEAD TIME SEGUN EL TIPO DE MATERIAL
    XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMAT%))
    PSM$ = Mid$(XX$, 82, 6)
   End If
99 AUTONOMIASM$ = PSM$

End Function
Function LEAD_TIME%(CI%)

   'Leer dato en campo LeTiCom y comparar para seguir
   LTME% = 0
   CONDI$ = "CODINT = " & CI%
   LCOM$ = SAFETEXT(VALOBADA("MASTER", "LeTiCom", CONDI$, "NOMESS"))
   
   If LCOM$ <> "" Then
      LTME% = XVALO(LCOM$)
      GoTo 99
   End If
   
   TIMAT% = TIMATE%(CI%)
   If TIMAT% > 0 Then ' FUNCION QUE DEVUELVE EL LEAD TIME SEGUN EL TIPO DE MATERIAL
    XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMAT%))
    LTME% = CVI(Mid$(XX$, 18, 2))
   End If
99 LEAD_TIME% = LTME%

End Function

Function VALCODPROV%(NPRO%, CIX%)  ' FUNCION QUE VALIDA QUE EL CODIGO DEL ARTICULO
    '                              SE ENCUENTRE ASOCIADO A UN PROVEEDOR
    Static CODPROVE%
    '
    If CODPROVE% = 0 Then
       CODPROVE% = -1
       If CONTROL$("OCOMPRA", "VALCOPRO") = "SI" Then CODPROVE% = 1
    End If
    VALCODPROV% = 1
    If CODPROVE% < 1 Or CIX% < 1 Then Exit Function
    '
    CONDI$ = "Num_Prov = " & NPRO% & " and Cod_Inte = " & CIX%
    If CantRegistros("PROVECOD", CONDI$) < 1 Then
      Call MENSERR(24, "Código Artículo '" & CODEXT$(CIX%) & "' no Relacionado al\Proveedor Seleccionado")
      VALCODPROV% = 0
    End If
    '
End Function

Sub CONVERUNID(USON, UNXX$)
    '
10  If Abs(USON) > 0 Then
      If Abs(USON) < 1 Then
        For KKIL& = 1 To ULTREG&("UNIMED")
          RLX$ = REGLEIDO$("UNIMED", KKIL&)
          ULEI$ = UCase$(TRIM$(Mid$(RLX$, 1, 4)))
          If ULEI$ = UCase$(UNXX$) Then
            UREP$ = UCase$(TRIM$(Mid$(RLX$, 33, 4)))
            If UREP$ <> ULEI$ Then
              FACO = CVS(Mid$(RLX$, 37, 4))
              If FACO > 1 Then
                UNXX$ = TRIM$(Mid$(RLX$, 33, 4))
                USON = USON * FACO
                GoTo 10
              End If
            End If
          End If
        Next KKIL&
      End If
    End If
End Sub

Function FACONVERT(UNISTO$, UNIFOR$)
    '
    FACOV = 1
    UNSTO$ = TRIM$(UCase$(UNISTO$))
    UNFOR$ = TRIM$(UCase$(UNIFOR$))
    '
10  If UCase$(UNSTO$) <> UCase$(UNFOR$) Then
      For KKIL& = 1 To ULTREG&("UNIMED")
         RLX$ = REGLEIDO$("UNIMED", KKIL&)
         ULEI$ = UCase$(TRIM$(Mid$(RLX$, 33, 4)))
         If ULEI$ = UNFOR$ Then
           UREP$ = UCase$(TRIM$(Mid$(RLX$, 1, 4)))
           If UREP$ <> ULEI$ Then
              FACO = CVS(Mid$(RLX$, 37, 4))
              If FACO > 1 Then
                UNFOR$ = TRIM$(UCase$(Mid$(RLX$, 1, 4)))
                FACOV = FACOV * FACO
                GoTo 10
              End If
           End If
         End If
      Next KKIL&
    End If
    '
    FACONVERT = FACOV
    '
End Function

Function REDON_INDIVIS(CI%, CANT#, Optional MODO$)
   '
   'SI MODO$ ES IGUAL A VENTA ES CALCULA EL LOTE INDIVISIBLE DE VENTAS
   If UCase$(MODO$) = "VENTA" Then LINDI = LINDIVEN(CI%): GoTo 10
   LINDI = LINDIVIS(CI%)
   '
10 If LINDI < 0.00001 Then
      REDON_INDIVIS = CANT#
      Exit Function
   End If
   CANTNUE& = Int(CANT# / LINDI)
   RESTO# = CANT# - (CANTNUE& * XVALO(LINDI))
   
   If RESTO# > 0 Then
      CANTNUE& = CANTNUE& + 1
   End If
   
   REDON_INDIVIS = LINDI * CANTNUE&
   
End Function

Function PRESENVASE$(CI%)
    
    PRVASE$ = ""
    If CI% > 0 Then
      CAN = ROUND(LINDIVIS(CI%), 0)
      If CAN >= 1 Then
        EMBA$ = TRIM$(ATRIB1$(CI%))
        If EMBA$ = "" Then EMBA$ = "Bultos"
        TIPUN$ = UNIMED$(CI%)
        '
        PRVASE$ = EMBA$ & " de " & CAN & " " & TIPUN$
        '
      End If
    End If
    '
99  PRESENVASE$ = PRVASE$

End Function

Function CANBULTOS(CI%, CANT#, Optional MODO$)
   '
   CANB = 0
   '
   ' SI MODO ES VENTA TRAE EL LOTE INDIVISIBLE DE VENTAS
   If CI% > 0 Then
     If CANT# > 0 Then
       If UCase$(MODO$) = "VENTA" Then
          CANPORBTO = ROUND(LINDIVEN(CI%), 0)
       Else
          CANPORBTO = ROUND(LINDIVIS(CI%), 0)
       End If
       '
       If CANPORBTO >= 1 Then
          CANB = Int(CANT# / CANPORBTO)
          RESTO# = CANT# - (CANB * CANPORBTO)
          If RESTO# > 0 Then
            CANB = CANB + 1
          End If
       End If
     End If
   End If
   '
   CANBULTOS = CANB
   '
End Function


Function BIENCAP%(CI%)
   '
    BICA% = 0
    If CI% < 1 Then GoTo 99
    '
    CONDI$ = "CODINT = " & CI%
    BICA% = XVALO(VALOBADA("MASTER", "BIEN_CAP", CONDI$, "NOMESS"))
   
99  BIENCAP% = BICA%
   
End Function

Function STATMOVI%(CI%)
   '
   STATMOVI% = 0
   On Error Resume Next
   If STATUIT%(CI%) <> 0 Then STATMOVI% = XVALO(MasterTemp!Stat_Movi)
   '
End Function

Function UNIMEAFIP%(CI%)

   'MODIFICADO 09 POR QUE NO DEVOLIVA POR DEFECTO LA 'UNIDAD' 7
   UNIMEAFIP% = 7 ' POR DEFECTO DEVUELVE UNIDAD
   
   If CI% < 1 Then Exit Function
   '
   UNIM$ = VALOBADA("MASTER", "UMEDIDA", "CODINT =" & CI%, "NOMESS")
   X$ = FILTERGETRECORD$("UNIMED", 1, TRIM$(UNIM$))
   UMEDAFIP% = CVI(Mid$(X$, 41, 2))
   '
   If UMEDAFIP% > 0 Then
     UNIMEAFIP% = UMEDAFIP%
   End If
   '
End Function
Sub CALCUVALRES(CODIMAQ$, Optional FECHASEL%)
    ' GRABA EL VALOR RESIDUAL EN LA TABLA
    
    'CALCULO LA DIFERENCIA DE AÑOS CALCULAR EL VALOR RESIDUAL
    CODEQUIPO$ = TRIM$(CODIMAQ$)
    '
    CALCUGENERAL% = 0
    'VALIDO SI SE PASO FECHA HASTA SI NO TOMA LA DE HOY
    If FECHASEL% > 0 Then ' FECHA DE RECALCULO TOTAL
      FE% = FECHASEL%
      CALCUGENERAL% = 1
    Else
      FE% = HOY(HO$) ' SI NO SE LE PASA FECHA TOMA LA DEL DIA
    End If
    '
    'FECHA  PARA TOMAR AÑO
    FECH = FECHATEXLAR$(FE%): ANIOACTUAL = Mid$(FECH, 7)
    'FECHA ACTIVACION PARA TOMAR EL AÑO
    Call ABREMAQUI
    CONDI$ = "CODIGO_EQU =  '" & CODEQUIPO$ & "'"
    FECHACTIV% = CVINT(VALOBADA("EQUIPOS", "FECHAACTIV_EQU", CONDI$, "NOMESS"))
    FECACTIV = FECHATEXLAR$(FECHACTIV%): ANIOACTIV = XVALO(Mid$(FECACTIV, 7))
    IMPORTEVTA = XVALO(VALOBADA("EQUIPOS", "ImpVta_Equ", CONDI$, "NOMESS"))

    '
    'CALCULO EL PORCENTAJE DE DEPECIACION EN BASE AL PLAZO DE AMORTIZACION CONTABLE
    PLAZAMORT% = 0
    PLAZAMORT% = XVALO(VALOBADA("EQUIPOS", "PLAZOAMORT_EQU", CONDI$, "NOMESS"))
    If PLAZAMORT% > 0 Then
      DEPRECANUAL = 100 / PLAZAMORT%
    End If
    '
    'CALCULO EL PORCENTAJE DE DEPECIACION EN BASE AL PLAZO DE AMORTIZACION FISCAL
    PLAZAMORTFIS% = 0
    PLAZAMORTFIS% = XVALO(VALOBADA("EQUIPOS", "PLAZOAMORTFISC_EQU", CONDI$, "NOMESS"))
    If PLAZAMORTFIS% > 0 Then
      DEPRECANUALFIS# = 100 / PLAZAMORTFIS%
    End If
    VALORORIG = XVALO(VALOBADA("EQUIPOS", "VALORIGPESOS_EQU", CONDI$, "NOMESS"))
    ANIOSTRANSC = 1 + (XVALO(ANIOACTUAL) - XVALO(ANIOACTIV))
    '
    If ANIOSTRANSC <= 0 Then
       VALORCALCULADO = VALORORIG
       ANIOSTRANSC = 0
       GoTo 50
    End If
    '
    Resultado = XVALO((VALORORIG * ANIOSTRANSC) * (DEPRECANUAL / 100))
    RESULTADOFISC = (VALORORIG * ANIOSTRANSC * ROUND(DEPRECANUALFIS#, 2))
    '
    'SI EL RESUTADO ES MAYOR AL VALOR ORIGINAL O ESTA VENDIDO
    If Resultado > VALORORIG Or IMPORTEVTA > 0 Then
         NUEVOVALOR = 0
    Else
         NUEVOVALOR = VALORORIG - XVALO(Resultado)
    End If
    '
    If RESULTADOFISC > VALORORIG Or IMPORTEVTA > 0 Then
         NUEVOVALORFISC = 0
    Else
         NUEVOVALORFISC = VALORORIG - XVALO(RESULTADOFISC)
    End If

    '
    'VALIDA QUE LA FECHA ACTUAL O LA FECHA DE CALCULO NO SEA < A LA FECHA DE ACTIVACION
    'ESTO ES PARA QUE SI EL BIEN TIENE FECHA DE ACTIVACION POSTERIOR A LA FECHA DE
    ' CALCULO O LA DE HOY EL VALOR RESIDUAL VA A SER 0 POR QUE A ESA FECHA EL BIEN NO EXISTE
    If FE% < FECHACTIV% Then NUEVOVALOR = 0
    '
    'VALIIDO SI ES POR CALCULO GENERAL ACTUALIZO EL MISMO Y LA FECHA DE CALCULO
50  If CALCUGENERAL% > 0 Then
       Call ACTUREGISTRO("EQUIPOS", "ValResCalc_Equ", CONDI$, NUEVOVALOR, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ValResCalcFisc_Equ", CONDI$, NUEVOVALORFISC, REGAF&, "NOMESS")
       FECHACALCU = FETEXCOR1$(FE%) 'FECHA DE CALCULO
       Call ACTUREGISTRO("EQUIPOS", "FechCalc_Equ", CONDI$, FECHACALCU, REGAF&, "NOMESS")
    
    Else ' SI ES ACTUALIZACION INDEPENDIENTE ACTUALIZO EL VALOR RESIDUAL COMUN
       Call ACTUREGISTRO("EQUIPOS", "ValorResid_Equ", CONDI$, NUEVOVALOR, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ANIOSTRANSC_EQU", CONDI$, ANIOSTRANSC, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ValorResidFisc_Equ", CONDI$, NUEVOVALORFISC, REGAF&, "NOMESS")
    End If
   
   'SE GRABA SIEMPRE
    Call ACTUREGISTRO("EQUIPOS", "DEPRECANUAL_EQU", CONDI$, FORMATNUM$(DEPRECANUAL, "F12.2"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("EQUIPOS", "DEPRECANUALFISC_EQU", CONDI$, FORMATNUM$(DEPRECANUALFIS, "F12.2"), REGAF&, "NOMESS")
    '
End Sub

 Function DERIMPO(CI%)
   On Error Resume Next
   TIMA% = TIMATE%(CI%)
   XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMA%))
   DERIM = CVS(Mid$(XX$, 88, 4))
   DERIMPO = DERIM
End Function

Function DEREXPO(CI%)
   On Error Resume Next
   TIMA% = TIMATE%(CI%)
   XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMA%))
   DEREX = CVS(Mid$(XX$, 92, 4))
   DEREXPO = DERIM
End Function

Function CODINTBARRA%(UPC$)
   '
   CODI% = 0
   If TRIM(UPC$) = "" Then GoTo 99
   '
   CODI% = XVALO(VALOBADA("MASTER", "CODINT", "COLOR = '" & UPC$ & "'", "NOMESS"))
   
99 CODINTBARRA% = CODI%
End Function

Function VOLUMENCOD#(CI%)
    '
    M3CAJA# = XVALO(VALOBADA("MASTER", "M3CAJA", "CODINT=" & CI%, "NOMESS"))
    VOLUMENCOD# = M3CAJA#
    '
End Function

Sub ADD_NOVEDAD(CI1%, NPROV%, PREC_ACT#, PREC_NUE#, OBSERVAC$, MONED%, MODO$)
    'NOVEDADES DE CAMBIOS DE PRECIO
    
    ' 'LISTA' SI ES ACTUALIZACION DE LISTA
    '"COSTO" SI ES ACTUALIZACION DE COSTO POR  PROVEEDOR
    User% = USNBR% Mod 100
    FE% = HOY(HO$)
    '
    'TOMA EL VALOR EN DOLARES
    If XVALO(MONED%) < 1 Then MONED% = 1
    COTIZA = TICAMONE(MONED%)
    
    '
    'ESTO ES PARA CUANDO SE INICIALIZA UN ITEM QUE NO REGISTRA IMPORTE ACTUAL
    If PREC_ACT# < 0.005 Then PREC_ACT# = PREC_NUE#
    
    If CI1% > 0 Then
       SQLX$ = "SELECT * FROM NOVEDADES "
       SQLX$ = SQLX$ + " WHERE CODINT = -1"
       '
       Set rst = New ADODB.Recordset
25     On Error Resume Next
       rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       
       With rst
          .AddNew
          !CODINT = CI1%
          !Nume_Usuar = User%
          !Fech_Mod = Now
          !Prec_Ant = PREC_ACT# * XVALO(COTIZA) ' EL PRECIO  SE GUARDA SIEMPRE EN PESOS
          !Prec_New = PREC_NUE# * XVALO(COTIZA) ' EL PRECIO  SE GUARDA SIEMPRE EN PESOS
          !NumProv = NPROV%
          !Observacion = Left$(OBSERVAC$, 512)
          !MONEDA = MONED%
          !Lista_Repos = MODO$
          !COTIZA_DOLAR = XVALO(COTIZA)
          .Update
       End With
       '
       rst.Close
       Set rst = Nothing
       '
    End If
    '
End Sub
Function COSTO_SEGUN_PROV(NPROV&, CI1%)
   'COSTO SEGUN PROVEEDOR GUARDADO EN LA TABLA RELACODPROV
   CONDI$ = "NUM_PROV = " & NPROV& & " AND COD_INTE = " & CI1%
   COSTO_SEGUN_PROV = XVALO(VALOBADA("RELACODPROV", "PRECIO", CONDI$, "NOMESS"))
End Function
Function FECHA_COSTO_POR_PROV%(NPROV&, CI1%)
 'FECHA DE COSTO SEGUN PROVEEDOR GUARDADO EN LA TABLA RELACODPROV
   CONDI$ = "NUM_PROV = " & NPROV& & " AND COD_INTE = " & CI1%
   FECHA_COSTO_POR_PROV% = CVINT(VALOBADA("RELACODPROV", "FECHA", CONDI$, "NOMESS"))
End Function

Function CODICLIENTE&(CODIGO$)
  'FUNCION CREADA PARA AHP QUE DEVUELVE EL CODIGO INTERNO PASANDOLE EL CODIGO EXTERNO DEL CLIENTE
  LONGIMAX% = XVALO(CONTROL$("MASTER", "LOMAXCOD"))
  If LONGIMAX% < 15 Then LONGIMAX% = 15
  If Len(CODIGO$) > LONGIMAX% Then
    Call COMUNI("LONGITUD EXCESIVA DE CODIGO \ Máximo Acatual: " & LONGIMAX% & " Caracteres")
    Exit Function
  End If
  '
  CODI$ = TRIM$(CODIGO$)
  If CODI$ = "" Then Exit Function
  '
  CONDI$ = "Codigo_Cliente = '" & CODIGO$ & "' AND STATUS >= 0"
  CODICLIENTE& = XVALO(VALOBADA("CODICLIENTE", "CODINT", CONDI$, "NOMESS"))
  '
End Function

Function SALDO_LOTE_PEDIDO(CI%, Optional LOTE$, Optional NPED&)
   '
   Call ABRECONEXION
   'SI SE LE PASA EL NUMERO DE PEDIDO, ESTE, NO SE TIENE EN CUENTA
   Call PoneEnCeroNull("PEDDETA", "IDENLOTE", REGAF&)
   CONDI$ = "CANPED > CantEnt AND STATUS < 8 "
   CONDI$ = CONDI$ + " AND CodiInt =  " & CI%
   If TRIM$(LOTE$) <> "" Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
   '
   ''SI SE LE PASA EL NUMERO DE PEDIDO ESTE NO SE TIENE EN CUENTA
   If NPED& > 0 Then CONDI$ = CONDI$ + "  And NROPED <> " & NPED&
   '
   RESERVADO = XVALO(VALOBADA("PEDDETA", "SUM(CANPED)", CONDI$, "NOMESS"))
   CONSUMO = XVALO(VALOBADA("PEDDETA", "SUM(CANTENT)", CONDI$, "NOMESS"))
   SALDO_LOTE_PEDIDO = RESERVADO - CONSUMO

End Function
Function RESERVAS_PEDIDAS_X_VIA_PROV(CI%, VIA%, PROV%)
   '
   Call ABRECONEXION
   'SI SE LE PASA EL NUMERO DE PEDIDO ESTE NO SE TIENE EN CUENTA
   ACUMU_ST = 0
   CONDI$ = "CANPED > CantEnt AND STATUS < 8 "
   CONDI$ = CONDI$ + " AND CodiInt =  " & CI%
   CONDI$ = CONDI$ + " AND IDENLOTE <> ''"
   '
   SQLX$ = "SELECT SUM(CANPED) AS PEDIDA, SUM(CANTENT) AS ENTREGADA"
   SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) WHERE " & CONDI$
   Set RS = READSET(SQLX$)
   With RS
      Do While Not .EOF
        LOTE$ = SAFETEXT$(!IdenLote)
        PROVEEDOR_LOTE% = PROVELOTE%(LOTE$)
        VIA_LOTE% = VIALOTE%(LOTE$)
        If PROVEEDOR_LOTE% = PROV% And VIA_LOTE% = VIA% Then
          ST = PEDIDA - ENTREGADA
          ACUMU_ST = ACUMU_ST + ST
        End If
        .MoveNext
      Loop
   End With

End Function

Sub CARGALOTE_PED(NUPED&)
    '
    'CARGA LA INFORMACION PARA ASIGNAR LOTES A LOS ITEMS DE UN PEDIDO.
    LOTESPED.List2.Clear
    For U& = 1 To ULTREG&("!CARDETPE")
      X$ = REGLEIDO$("!CARDETPE", U&)
      CAN = CVS(Mid$(X$, 47, 4))
      If CAN > 0.05 Then
        '
        CIXI% = CVI(Left$(X$, 2))
        If CIXI% > 0 Then
            TX$ = ""
            NUORIT% = CVI(Mid$(X$, 132, 2))
            LOTEX$ = Mid$(X$, 3, 16)
            DESCUTEX$ = TRIM$(Mid$(X$, 75, 14))
            PRELI = FORMATNUM$(CVS(Mid$(X$, 39, 4)), "F12.2")
            COEFCON = COEFDESCU(DESCUTEX$)
            NUPRESU& = CVS(Mid$(X$, 125, 4))
            FE% = CVI(Mid$(X$, 63, 2))
            '
            Call REPLA(TX$, CODEXT$(CIXI%), 1, 16) 'OJO-LONGITUD DE CODIGO EXTERNO
            Call REPLA(TX$, DESCRIT$(CIXI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN, "F11.1"), 50, 11)
            Call REPLA(TX$, LOTEX$, 62, 16)
            Call REPLA(TX$, FORMATNUM$(PRELI, "F12.2"), 80, 12)
            Call REPLA(TX$, SAFETEXT$(NUPRESU&), 94, 8)
            Call REPLA(TX$, FECHATEX$(FE%), 104, 8)
            Call REPLA(TX$, DESCUTEX$, 114, 14)
            LOTESPED.List2.AddItem TX$
         End If
        End If
    Next U&
    '
   'Presentar formulario para asignar lotes
   If LOTESPED.List2.ListCount > 0 Then
     LOTESPED.Label16 = NUPED&
30   LOTESPED.Show 1
     If LOTESPED.List2.ListCount < 1 Or LOTESPED.Label7 = "FALTAN_LOTES" Then
          LOTESPED.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         'PRIMERO VALIDO QUE TODOS TENGAN LOTE
         For I& = 1 To LOTESPED.List2.ListCount
           LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(I& - 1), 62, 16))
           CI% = CODINT%(Mid$(LOTESPED.List2.List(I& - 1), 1, 16))
           If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote Para el Código: " & CODEXT(CI%))
              GoTo 30 ' PRESENTA NUEVAMENTE EL LOTESPED
           End If
         Next I&

         j& = 0
         'SI EL ARCHIVO TENIA CODIGOS NO TRAZABLES LOS CARGA PRIMERO
         For YX& = 1 To ULTREG&("!CARDETPE")
          YY$ = REGLEIDO$("!CARDETPE", YX&)
          CI% = CVI(Mid$(YY$, 1, 2))
          CAN = CVD(Mid$(YY$, 19, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If ESTRAZABLE%(CI%) < 1 Then
                 j& = j& + 1
                 Call REPLA(YY$, MKI$(Int(j&)), 132, 2)
                 Call GRAREG("!CARDETPE", YY$, j&)
              End If
            End If
          End If
        Next YX&
        'FE% = HOY(HO$)
        '
        'RECORRE LOS LOTES TRAZABLES Y LOS GRABA A CONTINUACION EN CARDETPE.
        For I& = 1 To LOTESPED.List2.ListCount
             'CODIGO DE LIST DE LOTESPED
              CI% = CODINT%(TRIM$(Left$(LOTESPED.List2.List(I& - 1), 16)))
              
              CANTI = XVALO(Mid$(LOTESPED.List2.List(I& - 1), 50, 11))
              LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(I& - 1), 62, 16))
              PRELI = XVALO(Mid$(LOTESPED.List2.List(I& - 1), 80, 12))
              'USOII = XVALO(Mid$(LOTESPED.List2.List(I& - 1), 85, 12))
              NUPRESU& = XVALO(Mid$(LOTESPED.List2.List(I& - 1), 94, 8))
              FE% = FECHANUM(Mid$(LOTESPED.List2.List(I& - 1), 104, 8))
              DESCUTEX$ = Mid$(LOTESPED.List2.List(I& - 1), 114, 14)
              '
              j& = j& + 1
              ZZ$ = ""
              Call REPLA(ZZ$, MKI$(CI%), 1, 2)
              Call REPLA(ZZ$, LOTE$, 3, 16)
              Call REPLA(ZZ$, MKS$(PRELI), 39, 4)
              Call REPLA(ZZ$, MKS$(CANTI), 47, 4)   ' CANTIDAD
              Call REPLA(ZZ$, MKS$(PRELI), 51, 4)
              Call REPLA(ZZ$, MKS$(PRELI), 65, 4)
              Call REPLA(ZZ$, MKI$(FE%), 63, 2)
              Call REPLA(ZZ$, DESCUTEX$, 75, 14)
              Call REPLA(ZZ$, MKS$(NUPRESU&), 125, 4)
              Call REPLA(ZZ$, MKI$(Int(j&)), 132, 2)
              '
              Call GRAREG("!CARDETPE", ZZ$, j&)
        Next I&
        '
        Call SETULTREG("!CARDETPE", j&)
     End If
   End If
   '
   LOTESPED.List2.Clear

End Sub
Function CANTICONSU(CI1%, Optional DES%, Optional HAS%)

     CONFECHA% = 0
     If DES% > 0 And HAS% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(DES%)
        HAS1 = FETEXCOR1$(HAS%)
     End If
     '
     CANTICONSU = 0
     SQLX$ = " SELECT SUM(CANTINGRE) AS INGRE1, SUM(CANTSALID) AS SALI1   FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
     SQLX$ = SQLX$ + " AND (CODIMOVI = 'RT' OR CODIMOVI = 'CF' OR  CODIMOVI = 'TR' OR  CODIMOVI = 'VM')"
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FECHMOV >= '" & DES1 & "' "
        SQLX$ = SQLX$ + " AND FECHMOV <= '" & HAS1 & "' "
     End If
     SQLX$ = SQLX$ + " GROUP BY COD_INTE"

     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
           CANTI = XVALO(!SALI1) - XVALO(!INGRE1)
           If CANTI < 0.05 Then CANTI = 0
           .MoveNext
       Loop
     End With
     '
     RS.Close
     Set RS = Nothing
     '
     CANTICONSU = CANTI
     '
End Function

Function CANTIRECIBIDA(CI1%, Optional DES%, Optional HAS%)

     CONFECHA% = 0
     If DES% > 0 And HAS% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(DES%)
        HAS1 = FETEXCOR1$(HAS%)
     End If
     '
     CANTI = 0
     ACUMULACANTI = 0
     SQLX$ = " SELECT SUM(CanToApro) AS APRO, SUM(CanToRech) AS RECHA ,SUM(CanToFalta) AS FALTA  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FechRecep >= '" & DES1 & "' "
        SQLX$ = SQLX$ + " AND FechRecep <= '" & HAS1 & "' "
     End If
     SQLX$ = SQLX$ + " GROUP BY COD_INTE"

     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
           CANTI = XVALO(!APRO) - (XVALO(!RECHA) + XVALO(!FALTA))
           If CANTI < 0.05 Then CANTI = 0
         .MoveNext
       Loop
     End With
     RS.Close
     Set RS = Nothing
     AA = CODEXT$(CI1%)
     '
     CANTIRECIBIDA = CANTI
     '
End Function


Function CANTICOMPRADA(CI1%, TOTALCOMPRA#, Optional DES%, Optional HAS%)
     CONFECHA% = 0
     If DES% > 0 And HAS% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(DES%)
        HAS1 = FETEXCOR1$(HAS%)
     End If
     
'     Funcion que me devuelva la cantidad comprada
'    fecha >= desde and fecha <= hasta.
'    statcom < 8
     ACUMULACANTI = 0
     SQLX$ = "Select  COD_INTE, UNIDS_ocom,  PrUn_Neto_Pes from OCOMDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom >= 0 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FEMI_OCOM >= '" & DES1 & "'"
        SQLX$ = SQLX$ + " AND FEMI_OCOM <= '" & HAS1 & "'"
     End If
     SQLX$ = SQLX$ + " ORDER BY NUME_OCOM"
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
         CANTI = XVALO(!Unids_Ocom)
         PREUNIX# = XVALO(!PrUn_Neto_Pes)
         Subtotal# = CANTI * PREUNIX#
         AcumTotal# = AcumTotal# + Subtotal#
         ACUMULACANTI = ACUMULACANTI + CANTI
         .MoveNext
       Loop
     End With
     '
     RS.Close
     Set RS = Nothing

     CANTICOMPRADA = 0
     TOTALCOMPRA# = 0
     If ACUMULACANTI > 0.005 Then
        TOTALCOMPRA# = AcumTotal#
        CANTICOMPRADA = ACUMULACANTI
     End If
     '
End Function

Function CompraPendientesAcumuladas(CI1%, Optional NUOC&, Optional SIN_PEDIDO%)

15   Call ABRECOMPRAS
     CompraPendientesAcumuladas = 0
     SQLX$ = " SELECT  Coef_Unids, cant_ocom, Cant_Rec,NUME_OCOM FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "AND N_ORSERV < 1 "
     If NUOC& > 0 Then SQLX$ = SQLX$ + "AND NUME_OCOM <> " & NUOC&
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = READSET(SQLX$)
     With OcoTemp
         Do While Not .EOF
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = XVALO(!cant_ocom) * COEFCAN
             CAREC = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CACOM - CAREC
             If SIN_PEDIDO% > 0 Then
               CONDI$ = "NUME_CLI > 0 AND  NUME_OCOM = " & XVALO(!nume_ocom)
               If CantRegistros&("OCOMENCA", CONDI$, "NOMESS") > 0 Then GoTo 50
             End If
             CompraPendientesAcumuladas = CompraPendientesAcumuladas + CAPEN
50           .MoveNext
         Loop
     End With
     
End Function
Function FabricacionPendiente(CI1%)
     Call ABREORFAB
     TOFABRI = 0
     'TOTAL DE CANTIDAD PENDIENTE EN ORDENES DE FABRICACION PENDIENTES DE CIERRE.
     SQLX$ = " SELECT CanProy,  CanApro,  NUMEORFA, FechEnSol, Referen, FechProyEntre, FechRecal FROM ORDEFABR WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = READSET(SQLX$)
     With OfaTemp
         Do While Not .EOF
           CACOM = XVALO(!CANPROY)
           CAREC = XVALO(!CANAPRO)
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
     End With
     '
     FabricacionPendiente = TOFABRI
End Function

Function VIAENTRE_COMPRA(NUMERO_ORD_COMPRA&)
   'DEVUELVE LA VIA CARGADA QUE SE CARGO EN LA ORDEN DE COMPRA.
   VIAENTRE_COMPRA = XVALO(VALOBADA("OCOMENCA", "VIA_ENTREGA", "NUME_OCOM = " & NUMERO_ORD_COMPRA&, "NOMESS"))
End Function
Function CANFAPEN(CI%) 'DEVUELVE LA CANTIDAD TOTAL EN PROCESO DE FABRICACIÓN DE UN ARTÍCULO
'\\\OT-07-0386-FL-17/10/07///
    CANFA# = 0
    If CI% > 0 Then
        Call APERBASDAT("ORFAB")
        SQLX$ = "SELECT * FROM ORDEFABR WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE CANAPRO < CANPROY "
        SQLX$ = SQLX$ + "AND COD_INTE = " & CI% & " "
        SQLX$ = SQLX$ + "AND StatuOrf < 3 "
        Set OrdefabrTemp = READSET(SQLX$)
        '
        With OrdefabrTemp
             Do While Not .EOF
               CANFA# = CANFA# + (XVALO(!CANPROY) - XVALO(!CANAPRO))
             .MoveNext
             Loop
        End With
    End If
    CANFAPEN = CANFA#
    '\\\OT-07-0386-FIN///
End Function
Function CANOCPEN(CI1%) ' DEVUELVE LAS CANTIDADES COMPRADAS PENDIENTES DE RECEPCIÓN
    ' '\\\OT-07-0386-FL-16/10/07///
    CANPEN# = 0
    If CI1% > 0 Then
        Call APERBASDAT("COMPRAS")
        SQLX$ = "SELECT *  FROM OCOMDETA WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE CANT_OCOM > CANT_REC "
        SQLX$ = SQLX$ + "AND COD_INTE = " & CI1% & " "
        SQLX$ = SQLX$ + "AND STAT_OCOM < 3 "
        Set OcomdetaTemp = READSET(SQLX$)
        '
        With OcomdetaTemp
             Do While Not .EOF
               CANPEN# = CANPEN# + (XVALO(!cant_ocom) - XVALO(!Cant_Rec))
             .MoveNext
             Loop
        End With
    End If
    CANOCPEN = CANPEN#
   ' '\\\OT-07-0386-FIN///
End Function
Function CANRECEP(CI1%)
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM BOLRECEP WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + "AND Cod_Inte = " & CI1% & " "
    '
    Set COBRECEPTMP = READSET(SQLX$)
    '
    TORECEPE = 0
    With COBRECEPTMP
         JJ& = 0
         K& = 0
         Do While Not .EOF
            CARECEP = XVALO(!CANTNOM) - XVALO(!CANTOAPRO)
            If CARECEP > 0 Then
               TORECEPE = TORECEPE + CARECEP
            End If
         .MoveNext
         Loop
    End With
    CANRECEP = TORECEPE
End Function

Function CANTIDAD_PEDIDA_PENDIENTE(CI1%)

     Call ABREPEDCLI
     TOPEDID = 0
     SQLX$ = " SELECT NROPED, NroClie, CanPed,  CantEnt, FechEmis, AcuDespa, AcuPrepa, FeSolEnt FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI1% & " "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
         Do While Not .EOF
           CAPEN = XVALO(!CANPED) - XVALO(!CantEnt)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
           End If
           .MoveNext
         Loop
     End With
     CANTIDAD_PEDIDA_PENDIENTE = TOPEDID

End Function

Function CODISECC#(CI%)
    '
    CODISECC# = 0
    If ExisteArchivo%(LUDAT$, "CODSEPRO.RFS") = 1 Then
        RFF$ = FILTERGETRECORD$("CODSEPRO", 1, MKI$(CI%))
        CODISECC# = CVS(Mid$(RFF$, 3, 4))
    End If
    '
End Function
