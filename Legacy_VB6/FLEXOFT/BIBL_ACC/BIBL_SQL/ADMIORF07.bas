Attribute VB_Name = "ADMIORF07"
Public CHECOMPO%
Public PRECOLAN%
Public PRECOCIE%
Public FRALOMAX%
Public VALINDOP%
Public INCLUPRI%
Dim APROLAN%

'
Sub GENLISORF()
   CONDI$ = "IDSUBOR <> '' "
   CONDI$ = CONDI$ & " AND STATUORF < 3 "
   CONDI$ = CONDI$ & " AND CanApro < CanProy "
   CONDI$ = CONDI$ & " ORDER BY IDSUBOR ASC "
   Call CARGALISEL("ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
   Call FRESHECHO("ECORDEF")
End Sub
'
Function CODIFAB%(NORFX$)
    CICON% = 0
    '
    If TRIM$(NORFX$) <> "" Then
      Call ABREORFAB
      Dim ORDEFA As ADODB.Recordset
      SQLX$ = "SELECT Cod_Inte FROM OrdeFabr WHERE IdSubOr = '" & NORFX$ & "'"
      Set ORDEFA = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      If Not (ORDEFA.BOF And ORDEFA.EOF) Then
         CICON% = ORDEFA!Cod_Inte
      End If
      ORDEFA.Close
      Set ORDEFA = Nothing
    End If
    '
    CODIFAB% = CICON%
End Function
'
Sub ANULARES(NORFA$)
    '
    Dim CIXA%(), IDLA$()
    CARESAN% = 0
    '
    Call APERBASDAT("RESERVA")
    '
    Screen.MousePointer = 11
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    '
5   'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With RESERTMP
10     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            CARESAN% = CARESAN% + 1
            ReDim Preserve CIXA%(CARESAN%), IDLA$(CARESAN%)
            CIXA%(CARESAN%) = XVALO(!Cod_Inte)
            IDLA$(CARESAN%) = SAFETEXT$(!IdenLote)
         .MoveNext
         Loop
       End If
    End With
    RESERTMP.Close
    '
    SQLX$ = "DELETE  FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    'Stocks.Execute SQLX$
     FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    For KKI% = 1 To CARESAN%
       Call ACSALOTE(CIXA%(KKI%), IDLA$(KKI%), "NOMESS")
    Next KKI%
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub ORFLAN()      ' LANZAMIENTO DE ORDEN DE FABRICACION
    '
        HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE LANZAMIENTO DE ORDENES DE FABRICACION"

100  NROOF$ = ""
     '
1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call ABREORFAB
       SQLX$ = " SELECT NumeOrFa, Referen, StatuOrf, IdSubOr, CanProy, Cod_Inte, FechLan FROM ORDEFABR "
       SQLX$ = SQLX$ + " WHERE StatuOrf = 2 OR StatuOrf = 0 "
       SQLX$ = SQLX$ + " order by IdSubOr ASC "
       'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       Dim ORFATEMP As ADODB.Recordset
       Set ORFATEMP = New ADODB.Recordset
       ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
       On Error Resume Next
       With ORFATEMP
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No hay Ordenes de Fabricación Proyectadas\o Suspendidas que Puedan Lanzarse.")
           Screen.MousePointer = 1
           On Error GoTo 0
           Exit Sub
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While Not .EOF
           If XVALO(Mid$(SAFETEXT$(!IDSUBOR), 11)) = 1 Then
             ZZ$ = REGBLAN$("!ORFANUL")
             REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
             Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
             Call REPLA(ZZ$, REFXX$, 5, 64)
             If !Statuorf = 2 Or !Statuorf = 0 Then
              FEX = CVINT(!FechLan)
              Call REPLA(ZZ$, " (" + FECHATEX$(FEX) + ")          ", 54, 15)
             End If
             JJ& = JJ& + 1
             Call GRAREG("!ORFANUL", ZZ$, JJ&)
           End If
         .MoveNext
         Loop
       End With
       Set ORFATEMP = Nothing
       Call SETULTREG("!ORFANUL", JJ&)
       Call CIERRARCH("!ORFANUL")
       Screen.MousePointer = 1
     End If
     '
1405 Screen.MousePointer = 1
     Call FRESHECHO("!ORFANUL")
     If ULTREG&("!ORFANUL") < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         GoTo 1499
     End If
     Call SELECHO("!ORFANUL/Ordenes de Trabajo para Lanzamiento")
     '
1410 NROOFL& = XVALO(VALACT1$("!ORFANUL"))
     If NROOFL& < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         GoTo 1499
     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE NumeOrFa= " & NROOFL & ""
     Dim CARGAOBJTMP As ADODB.Recordset
     Set CARGAOBJTMP = New ADODB.Recordset
     CARGAOBJTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     '
     'Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     
     '
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
     Call REPLA(VDEF$, SAFETEXT$(!REFEREN), 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, SAFETEXT$(!Unidad), 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, SAFETEXT$(!DESTINO), 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(!Statuorf), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     For I& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '

1420 YY$ = OBJPLA$("ORFALAN/Lanzamiento de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     '
     For I& = j& + 1 To ULTREG&("!ORFANUL")
        Call GRAREG("!ORFANUL", String$(256, 0), I&)
     Next I&
     '
     FLAN% = CVI(Mid$(YY$, 113, 2))
     If FLAN% < HO% Then FLAN% = HO%
     Call LANORF(NROOFL&, FLAN%)
     '
     If APROLAN% = 1 Then
        j& = 0
        For I& = 1 To ULTREG&("!ORFANUL")
           XX$ = REGLEIDO$("!ORFANUL", I&)
           If CVS(Left$(XX$, 4)) <> NROOFL& Then
              j& = j& + 1
              Call GRAREG("!ORFANUL", XX$, j&)
           End If
        Next I&
        Call SETULTREG("!ORFANUL", j&)
     End If
      Call COMUNI("La Orden de Fabricación " & NROOFL & "  \ha sido Lanzada")
     '
     GoTo 1405
     '
'1499 Call ACTECOR(0)
'////////***ANTIGUO***//////
'    HO% = HOY(HOS$)
'    LU$ = LUDAT$
'    '
'    TITU$ = "PROCESO DE LANZAMIENTO DE ORDENES DE FABRICACION"
'
'100  NROOF$ = ""
'     '
'1400 If INDI% = 0 Then
'       '
'       Screen.MousePointer = 11
'       Call COPYSTRU("ORFANUE", "!ORFANUL")
'       J& = 0
'       '
'       For I& = 1 To ULTREG&("ORFANUE")
'         XX$ = REGLEIDO$("ORFANUE", I&)
'         ST% = Asc(Mid$(XX$, 112, 1))
'         '
'         If ST% <> 3 Then
'            If CVI(Mid$(XX$, 119, 2)) > 0 Then
'                Call REPLA(XX$, MKI$(0), 119, 2)
'                Call GRAREG("ORFANUE", XX$, I&)
'            End If
'         End If
'         '
'         If ST% = 0 Or ST% = 2 Then           ' solo proyectadas y suspendidas
'            FEX = CVI(Mid$(XX$, 113, 2))
'            Call REPLA(XX$, " (" + FECHATEX$(FEX) + ")          ", 54, 15)
'            J& = J& + 1
'            Call GRAREG("ORFANUL", XX$, J&)
'         End If
'       Next I&
'       '
'       Call SETULTREG("ORFANUL", J&)
'       Call CIERRARCH("ORFANUL")
'       INDI% = 1
'       '
'       Screen.MousePointer = 1
'       If ULTREG&("ORFANUL") < 1 Then
'          Call MENSERR(24, "No hay Ordenes de Fabricación Proyectadas\o Suspendidas que Puedan Lanzarse.")
'          Call CIERRARCH("ORFANUL")
'          Call CIERRARCH("*.*")
'          GoTo 1499
'       End If
'       '
'     End If
'
'1405 Call FRESHECHO("ORFANUL")
'     If ULTREG&("ORFANUL") < 1 Then
'         INDI% = 0
'         Call CIERRARCH("*.*")
'         GoTo 1499
'     End If
'     Call SELECHO("ORFANUL/Ordenes de Trabajo para Lanzamiento")
'     '
'1410 NROOFL& = XVALO(VALACT1$("ORFANUL"))
'     If NROOFL& < 1 Then
'         INDI% = 0
'         Call CIERRARCH("*.*")
'         GoTo 1499
'     End If
'     '
'     For I& = ULTREG&("ORFANUE") To 1 Step -1
'       XX$ = REGLEIDO$("ORFANUE", I&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
'     Next I&
'     '
'     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
'     INDI% = 0
'     Call CIERRARCH("*.*")
'     GoTo 1499
'     '
'1420 YY$ = OBJPLA$("ORFALAN/Lanzamiento de Ordenes de Fabricación", XX$)
'     If YY$ = "" Then GoTo 1405
'     '
'     For I& = J& + 1 To ULTREG&("ORFANUL")
'        Call GRAREG("ORFANUL", String$(256, 0), I&)
'     Next I&
'     '
'     FLAN% = CVI(Mid$(YY$, 113, 2))
'     If FLAN% < HO% Then FLAN% = HO%
'     Call LANORF(NROOFL&, FLAN%)
'     '
'     If APROLAN% = 1 Then
'        J& = 0
'        For I& = 1 To ULTREG&("ORFANUL")
'           XX$ = REGLEIDO$("ORFANUL", I&)
'           If CVS(Left$(XX$, 4)) <> NROOFL& Then
'              J& = J& + 1
'              Call GRAREG("ORFANUL", XX$, J&)
'           End If
'        Next I&
'        Call SETULTREG("ORFANUL", J&)
'     End If
'     '
'     GoTo 1405
'     '
1499 Call ACTECOR(0)
     '
End Sub

Sub ORFSUS()
     '
' Este proceso SUSPENDE una O/F para lo que pone el ESTADO de la O/F StatuOrf
' en '2' y el de todas las operaciones de esa O/F también en '2'. No se
' modifican las Reservas de Materiales
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "PROCESO DE SUSPENSION DE ORDENES DE FABRICACION"
 
100  Screen.MousePointer = 11
     NROOF$ = ""
     '
     Call ABREORFAB
     SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
     SQLX$ = SQLX$ + "WHERE StatuOrf <= 1 "
     SQLX$ = SQLX$ + "order by IdSubOr ASC"
     'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim ORFATEMP As ADODB.Recordset
     Set ORFATEMP = New ADODB.Recordset
     ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     '
     On Error Resume Next
     With ORFATEMP
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No Hay O/F's Pendientes\que Puedan Supenderse.")
         Screen.MousePointer = 1
         On Error GoTo 0
         Exit Sub
       End If
       On Error GoTo 0
       '
       JJ& = 0
       ULORF& = 0
       Do While Not .EOF
         If XVALO(Mid$(SAFETEXT$(!IDSUBOR), 4, 6)) <> ULORF& Then
           ZZ$ = REGBLAN$("!ORFANUL")
           REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
           Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
           Call REPLA(ZZ$, REFXX$, 5, 64)
           JJ& = JJ& + 1
           Call GRAREG("!ORFANUL", ZZ$, JJ&)
           ULORF& = XVALO(Mid$(SAFETEXT$(!IDSUBOR), 4, 6))
         End If
       .MoveNext
       Loop
     End With
     Set ORFATEMP = Nothing
     Call SETULTREG("!ORFANUL", JJ&)
     Call CIERRARCH("!ORFANUL")
     Screen.MousePointer = 1
     '
1405 Call FRESHECHO("!ORFANUL")
     If ULTREG&("!ORFANUL") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("!ORFANUL/Ordenes de Trabajo para Suspensión")
     '
1410 NROOFL& = XVALO(VALACT1$("!ORFANUL"))
     If NROOFL& < 1 Then
       INDI% = 0
       Exit Sub
     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE NumeOrFa= " & NROOFL & ""
     'Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim CARGAOBJTMP As ADODB.Recordset
     Set CARGAOBJTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
     Call REPLA(VDEF$, SAFETEXT$(!REFEREN), 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, SAFETEXT$(!Unidad), 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, SAFETEXT$(!DESTINO), 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(1), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     Set CARGAOBJTMP = Nothing
1412 YY$ = OBJPLA$("ORFALAN/Suspensión de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     Screen.MousePointer = 11
     '
     Call ABREORFAB
     SQLX$ = "SELECT Statuorf, IdSubOr FROM OrdeFabr "
     SQLX$ = SQLX$ + "where NumeOrFa =" & NROOFL & " AND StatuOrf < 2 "
     Dim ORDEFFFABR As ADODB.Recordset
     Set ORDEFFFABR = New ADODB.Recordset
     ORDEFFFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     '
1415 With ORDEFFFABR
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
            !Statuorf = 2
            .Update
            IDSUB$ = SAFETEXT$(!IDSUBOR)
            .MoveNext
            '
            SQLX$ = "SELECT StatOper FROM OperFab "
            SQLX$ = SQLX$ + "where IdSubOr = '" & IDSUB$ & "'"
            Dim OPERFFFAB As ADODB.Recordset
            Set OPERFFFAB = New ADODB.Recordset
            OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
1416        With OPERFFFAB
              On Error Resume Next
              .MoveFirst
              If Err < 1 Then
                Do While Not .EOF
                  !StatOper = 2
                  .Update
                  .MoveNext
                Loop
              End If
              On Error GoTo 0
            End With
            Set OPERFFFAB = Nothing
          Loop
       End If
     End With
     Set ORDEFFFABR = Nothing
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     Call COMUNI("La Orden de Fabricación " & NROOFL & "  \ha sido Suspendida")
     GoTo 100
     '
End Sub
'
Sub ORFPRI()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "PROCESO DE CAMBIO DE PRIORIDAD DE ORDENES DE FABRICACION"

100  NROOF$ = ""
     '
     Call ABREORFAB
     SQLX$ = " SELECT NumeOrFa, Referen, StatuOrf, IdSubOr, CanProy, Cod_Inte, FechLan FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE StatuOrf =1 "
     SQLX$ = SQLX$ + " order by IdSubOr ASC "
     'Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim ORFATEMP As ADODB.Recordset
     Set ORFATEMP = New ADODB.Recordset
     ORFATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     '
     On Error Resume Next
     With ORFATEMP
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Ordenes de Fabricación\en Curso para Cambio de Prioridad.")
         Screen.MousePointer = 1
         On Error GoTo 0
         Exit Sub
       End If
       On Error GoTo 0
       '
       JJ& = 0
       Do While Not .EOF
         If XVALO(Mid$(SAFETEXT$(!IDSUBOR), 11)) = 1 Then
           ZZ$ = REGBLAN$("!ORFANUL")
           REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
           Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
           Call REPLA(ZZ$, REFXX$, 5, 64)
           If !Statuorf = 2 Or !Statuorf = 0 Then
            FEX = CVINT(!FechLan)
            Call REPLA(ZZ$, " (" + FECHATEX$(FEX) + ")          ", 54, 15)
           End If
           JJ& = JJ& + 1
           Call GRAREG("!ORFANUL", ZZ$, JJ&)
         End If
       .MoveNext
       Loop
     End With
     Set ORFATEMP = Nothing
     '
     Call SETULTREG("!ORFANUL", JJ&)
     Call CIERRARCH("!ORFANUL")
     Screen.MousePointer = 1
     '
1405 Call FRESHECHO("!ORFANUL")
     If ULTREG&("!ORFANUL") < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         GoTo 1499
     End If
     Call SELECHO("!ORFANUL/Ordenes de Trabajo p/Cambio de Prioridad")
     '
1410 NROOFL& = XVALO(VALACT1$("!ORFANUL"))
     If NROOFL& < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         GoTo 1499
     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE NumeOrFa= " & NROOFL & ""
     SQLX$ = SQLX$ + " order by IdSubOr ASC "
     'Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim CARGAOBJTMP As ADODB.Recordset
     Set CARGAOBJTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

     VDEF$ = REGBLAN$("ORFANUE")
     With CARGAOBJTMP
       On Error Resume Next
       .MoveFirst
       On Error GoTo 0
       Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
       Call REPLA(VDEF$, SAFETEXT$(!REFEREN), 5, 64)
       Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
       Call REPLA(VDEF$, SAFETEXT$(!Unidad), 73, 4)
       Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
       Call REPLA(VDEF$, MKS$(!CanProy), 97, 4)
       Call REPLA(VDEF$, SAFETEXT$(!DESTINO), 101, 8)
         ENTSOL% = CVINT(!FechEnSol)
         If ENTSOL% < 1 Then ENTSOL% = HOY(HOS$)
       Call REPLA(VDEF$, MKI$(ENTSOL%), 109, 2)
       Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
       Call REPLA(VDEF$, Chr$(!Statuorf), 112, 1)
         LANPRO% = CVINT(!FechLan)
         If LANPRO% < 1 Then LANPRO% = HOY(HOS$)
       Call REPLA(VDEF$, MKI$(LANPRO%), 113, 2)
       Call REPLA(VDEF$, MKS$(!SecciPro), 115, 4)
       Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     Set CARGAOBJTMP = Nothing
     XX$ = VDEF$
     '
1420 YY$ = OBJPLA$("ORFAPRI/Cambio Prioridad o Sección de O/F's", XX$)
     If YY$ = "" Then GoTo 1405
     '
     PRIO% = Asc(Mid$(YY$, 111, 1))
     If PRIO% > 9 Then
       PRIO% = PRIO% Mod 10
       Call REPLA(YY$, Chr$(PRIO%), 111, 1)
     End If
     SEPROD = CVS(Mid$(YY$, 115, 4))
     ENTSOL% = CVI(Mid$(YY$, 109, 2))
     LANPRO% = CVI(Mid$(YY$, 113, 2))
     '
     Screen.MousePointer = 11
     Call ABREORFAB
     'With OrdeFabr
       '.FindFirst "NumeOrfa = " & NROOFL&
       SQLX$ = "SELECT * FROM OrdeFabr "
       SQLX$ = SQLX$ + " WHERE NumeOrfa = " & NROOFL& & ""
       Dim NUMMMORFA As ADODB.Recordset
       Set NUMMMORFA = New ADODB.Recordset
       NUMMMORFA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       On Error Resume Next
       With NUMMMORFA
          .MoveFirst
          If Err < 1 Then
1425      'If .NoMatch = False Then
            '.Edit
            !Priorid = PRIO%
            !SecciPro = SEPROD
            !FechEnSol = CVDAT(ENTSOL%)
            !FechLan = CVDAT(LANPRO%)
            .Update
            '.FindNext "NumeOrfa = " & NROOFL&
            'GoTo 1425
            .MoveNext
          End If
          On Error GoTo 0
     End With
     Set NUMMMORFA = Nothing
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     Call COMUNI("Prioridad Cambiada")
     GoTo 1405
     '

1499 End Sub
'
Sub ORFCIE(Optional NUMERORF$, Optional CANTAPRO, Optional FECHINFO$)
     '
     MODOGABAL% = 0: MODOCIE% = 0
     If InStr(1, EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
     MODOMEDITEA% = 0
     If InStr(1, EMPREAC$, "MEDITEA") > 0 Then MODOMEDITEA% = 1
     '
     TIPOTRAN% = 0: CIERRAOF% = 0
90   If MODOGABAL% = 1 Then
        '
        OPXX% = 1
        MODOCIE% = OPXX% ' GUARDA LA MODALIDAD DE CIERRE EN EL CASO DE GABAL
        If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
        If OPXX% = 2 Then
            If DERACCE%("CIEMAOF", "Cierre Manual de O-Fabricacion") >= 2 Then GoTo 95
            GoTo 90
        End If
        '
91      CIERORFA.Show 1
        '
        On Error GoTo msglafe
        '
        QCAJAS% = CIERORFA.List1.ListCount
        If QCAJAS% < 1 Then
            Unload CIERORFA
            Exit Sub
        End If
        NORFA$ = TRIM$(CIERORFA.Label3)
        CIXI% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
        LOTEVEN = LINDIVEN(CIXI%)
        '
        'CANTIDAD FABRICADA CALCULADA A PARTIR DEL LOTE DE VENTAS Y LA CANTIDAD DE CAJAS LEIDAS
        CANTINFO# = QCAJAS% * LOTEVEN
' ****** CAMBIE ACA
        'VERIFICA SI LA ULTIMA CAJA TIENE UNA CANTIDAD MENOR
        TXTX$ = CIERORFA.List1.List(CIERORFA.List1.ListCount - 1)
        ULTCAJA# = XVALO(Mid$(TXTX$, 64, 14))
        If ULTCAJA# > 0.05 Then
            CANTINFO# = CANTINFO# - LOTEVEN + ULTCAJA#
            CIERRAOF% = 1
        End If
        '
        CI% = CIXI%
        CODCON$ = CODEXT$(CI%)
        DESCON$ = DESCRIT$(CI%)
        LOTE$ = NORFA$
        CMOV$ = "OF"
        DOPRI$ = NORFA$
        DOSEC$ = "IF-" + TRIM$(Mid$(NORFA$, 4))
        REFE$ = "Fabricación Propia"
        VUNI# = 0
        MONE$ = "$"
        NC% = 0
        DP2% = 0
            ' TOMA EL DEPOSITO ASIGNADO AL ARMADOR
            ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IdSubOr = '" & NORFA$ & "'", "NOMESS"))
            If ARMAD% > 0 Then
                TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
                DP2% = Asc(Mid$(TX$, 71, 1))
            Else
                DP2% = XVALO(Control$("", "DEPODEST"))
            End If
        If DP2% < 1 Then
            Call COMUNI("Imposible Continuar. Falta Definir Deposito de Consumo")
            Unload CIERORFA
            Exit Sub
        End If
        DEPO% = XVALO(Control("", "DEPOCIE"))
            If DEPO% < 1 Then
                Call COMUNI("Imposible Continuar. Falta Definir Deposito de Cierre")
                Unload CIERORFA
                Exit Sub
            End If
' ****** CAMBIE ACA
        'ESTA VARIABLE DEJA AL MOVIMIENTO DE ENTRADA LISTO PARA QUE SEA RECIBIDO.
        'SE APLICA SOLO PARA ARMADORES. EN LAS OF INTERNAS NO HAY RECEPCION
        If ARMAD% > 0 Then TIPOTRAN% = 1
        '
        CANTI = CANTINFO#
        SCRAP = 0
        CANPRO# = VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
        CANAPR# = VALOBADA("ORDEFABR", "CanApro", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
        CAPENDE# = CANPRO# - CANAPR#: If CAPENDE# < 0 Then CAPENDE# = 0
        FECHA% = HOY(HOS$)
        '
        On Error GoTo 0
        '
        GoTo 1428

     End If
     '
95   HO% = HOY(HOS$)
     LU$ = LUDAT$
       CANTAPROX = 0
       On Error Resume Next
       CANTAPROX = CANTAPRO + 0
       On Error GoTo 0
     NUMERORFX$ = TRIM$(NUMERORF$)
     FECHCIE% = FECHANUM(FECHINFO$)
     '
     If PRECOCIE% = 0 Then
        PRECOCIE% = (-1)
        If Control$("", "PRECOCIE") = "SI" Then
           PRECOCIE% = 1
        End If
     End If
     '
     If NUMERORFX$ <> "" Then
        NROOF$ = NUMERORFX$
        GoTo 104
     End If
     '
     Call ACTECOR(2)     ' parametro 2 para seleccionar solo las que estan en status 1
     '
     TITU$ = "PROCESO DE CIERRE DE ORDENES DE FABRICACION"

100  Screen.MousePointer = 1
     NROOF$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
104  If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
        ACONEC$ = "TRAZAB07"
        If EXISTE%("TRAZAB07.EXE") < 1 Then ACONEC$ = "TRAZAB08"
        Call CONECRUN(ACONEC$ + "/OFABRIC" + NROOF$, "Registro de Trazabilidad de O/F's")
        Exit Sub
     End If
     '
     
     Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
'CODIGO ANTERIOR A CAMBIO PARA ARCON.
     '*************************** MEDITEA
     CIN% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
     If ESTRAZABLE(CIN%) = 2 Then
        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CIN% & " AND DEPONUMSE < 0 AND Dic_Ensa LIKE 'A%'"
        CANRE& = CantRegistros&("TANUMSE", CONDI$)
        If CANRE& < 1 Then ' si tiene menos de un registro
           Call MENSERR(24, "No Existen Números de Serie Generados y Aprobados Para Esta Orden")
           GoTo 100 ' vuleve al objeto
        Else
           SELNUMSE.Label34 = NORFA$
           Screen.MousePointer = 1
           If SELNUMSE.Visible = True Then SELNUMSE.Hide
           SELNUMSE.Show 1
           GoTo 98
        End If

     End If
''CODIGO NUEVO CON VARIACION POR ARCON
'     '*************************** MEDITEA
'     CIN% = VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'", "NOMESS")
'     If ESTRAZABLE(CIN%) = 2 Then
'        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CIN% & " AND DEPONUMSE < 0 "
'        'SI NO ES ARCON FILTRA QUE ESTE APROBADO
''        If UCase$(InStr(1, EMPREAC$, "ARCON")) < 1 Then CONDI$ = CONDI$ + " AND Dic_Ensa LIKE 'A%'"
'        '
'        CANRE& = CantRegistros&("TANUMSE", CONDI$)
'        If CANRE& < 1 Then ' si tiene menos de un registro
'           Call MENSERR(24, "No Existen Números de Serie Generados y Aprobados Para Esta Orden")
'           GoTo 100 ' vuleve al objeto
'        Else
'                SELNUMSE.Label34 = NORFA$
'                Screen.MousePointer = 1
'                If SELNUMSE.Visible = True Then SELNUMSE.Hide
'                SELNUMSE.Show 1
'                GoTo 98
'        End If
'
'     End If
     Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "'"
     'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim ORDEFABTMP As ADODB.Recordset
     Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

     With ORDEFABTMP
       On Error Resume Next
       .MoveFirst
       If Err Then GoTo 1410
       On Error GoTo 0
         CI% = XVALO(!Cod_Inte)
         JJ& = 0
         Y$ = REGBLAN("ORDEFABR")
         '
1405     Call REPLA(Y$, MKS$(!NUMEORFA), 1, 4)
         Call REPLA(Y$, SAFETEXT$(!REFEREN), 5, 64)
         Call REPLA(Y$, SAFETEXT$(!IDSUBOR), 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechGen)), 81, 2)
         Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
         Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
         Call REPLA(Y$, MKD$(!CanProy), 89, 8)
         '
         On Error Resume Next
         Call REPLA(Y$, SAFETEXT$(!DESTINO), 97, 8)
         On Error GoTo 0
         '
         Call REPLA(Y$, MKI$(CVINT(!FechEnSol)), 105, 2)
         Call REPLA(Y$, Chr$(!Priorid), 107, 1)
         Call REPLA(Y$, Chr$(!Statuorf), 108, 1)
         Call REPLA(Y$, MKI$(CVINT(!FechLan)), 109, 2)
         Call REPLA(Y$, MKS$(!SecciPro), 113, 4)
           FEXIE% = HO%
           If FECHCIE% > 0 Then FEXIE% = FECHCIE%
         Call REPLA(Y$, MKI$(FEXIE%), 117, 2)
           CAPENDE# = XVALO(!CanProy) - XVALO(!CanApro): If CAPENDE# < 0 Then CAPENDE# = 0
           CAPEN1# = CAPENDE#
           If CANTAPROX > 0 Then CAPEN1# = CANTAPROX
         Call REPLA(Y$, MKD$(CAPEN1#), 119, 8)
         Call REPLA(Y$, MKD$(0), 127, 8)
         Call REPLA(Y$, SAFETEXT$(!IdenRetra), 135, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechProyEntre)), 147, 2)
         Call REPLA(Y$, MKI$(CVINT(!FechTempEntre)), 149, 2)
         Call REPLA(Y$, MKI$(CVINT(!FechRecal)), 151, 2)
         Call REPLA(Y$, SAFETEXT$(!OBSERVA), 155, 64)
     End With
     Set ORDEFABTMP = Nothing
     '
     CODI1% = CVI(Mid(Y$, 83, 2)) 'PARA MOSTRAR EL DEPOSITO DE CONSUMO Y
     DEPENT1% = DEPENART%(CODI1%) 'DE ENTRADA QUE TIENE POR DEFAULT
     
     ' Si existe deposito por default para la entrada de materiales
     ' presenta este ultimo, aunque tenga uno el articulo
     DEPOENTRA% = XVALO(Control("", "DEPOCIE"))
     If DEPOENTRA% > 0 Then DEPENT1% = DEPOENTRA%
     '
     DEPCONS1% = DEPCONART%(CODI1%) 'EL ARTICULO CUANDO HACEMOS EL CIERRE
     XX$ = Y$ + Chr$(DEPENT1%) + Chr$(DEPCONS1%)
     GoTo 1420
     '
1410 Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     GoTo 100
     '
1420 YY$ = OBJPLA$("ORFACIE/Cierre Parcial de Ordenes de Fabricación", XX$ + Chr$(0) + Chr$(0))
     If YY$ = "" Then GoTo 98
     '
     FECHA% = CVI(Mid$(YY$, 117, 2))
     FEX = FECHA%
     FECHACIOF$ = FECHATEX$(FEX)
     If TRIM$(FECHACIOF$) = "" Then
       Call MENSERR(24, "Fecha No Válida")
       GoTo 1420
     End If
     '
     CI% = CVI(Mid$(YY$, 83, 2))
     CODCON$ = CODEXT$(CI%)
     DESCON$ = DESCRIT$(CI%)
     LOTE$ = NORFA$
     CMOV$ = "OF"
     DOPRI$ = NORFA$
     DOSEC$ = "IF-" + TRIM$(Mid$(NORFA$, 4))
     REFE$ = "Fabricación Propia"
     VUNI# = 0
     MONE$ = "$"
     NC% = 0
     DEPO% = Asc(Mid$(YY$, 257, 1))
     DP2% = Asc(Mid$(YY$, 258, 1))
     CANTI = CVD(Mid$(YY$, 119, 8))
     SCRAP = CVD(Mid$(YY$, 127, 8))
     '
'     MODOGABAL% = 0
'     EPAC$ = TRIM$(UCase$(EMPREAC$))
'     If InStr(EPAC$, "GABAL") > 0 Then MODOGABAL% = 1
'     If MODOGABAL% > 0 Then
'       'VALIDO QUE POSEA ASIGNACION DE ARMADOR EXTERNO
'       CONDI$ = "Doc_Orig =  '" & NORFA$ & "' AND N_OrServ  > 0 "
'       If CantRegistros&("OCOMDETA", CONDI$, "NOMESS") <= 0 Then
'          OPX% = COMALTER("Atención !!!  Orden de Fabricación Sin Asignación de Armador Externo\\Cancelar Cierre\Cerrar Igual")
'          If OPX% = 2 Then
'               GoTo 1422
'             Else
'               GOTO 100
'          End If
'       End If
'       XANTI = LINDIVEN(CI%)
'       If CANTI Mod XANTI <> 0 Then ' SI NO ES MULTIPLO NO DEJA CONTINUAR
'          Call MENSERR(24, "Imposible Cerrar Orden De Fabricación \ Cantidad Debe ser Multiplo de Lote (" & XANTI & ")")
'          GoTo 1420
'       End If
'       '
'       If CANTI > CAPENDE# Then
'          Call MENSERR(24, "Imposible Cerrar Orden De Fabricación \ Cantidad Supera Cantidad Pendiente")
'          GoTo 1420
'       End If
'     End If
'     'FIN BLOQUE GABAL
     '
1422 If CANTAPROX >= 0.05 Then
        If Abs(CANTI - CANTAPROX) > 0.05 Or Abs(SCRAP) >= 0.05 Then
           Call MENSERR(24, "En Cierre por Parte de Producción\no Válido Cambiar Cantidades.")
           GoTo 1420
        End If
        If FECHA% <> FECHCIE% Then
           Call MENSERR(24, "En Cierre por Parte de Producción\no Válido Cambiar Fecha.")
           GoTo 1420
        End If
     End If
     '
     If CANTI < 0 Or SCRAP < 0 Then
        Call MENSERR(24, "Cantidades Incorrectas.")
        GoTo 100
     End If
     '
     XX$ = YY$
     If DEPO% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DEPO%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 1420
     End If
     '
     If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 1420
     End If
     '
     If SCRAP >= 0.05 Then
       If EXISTE%("RECHLI07.EXE") > 0 Then
          DSC% = DEPO%
          GoTo 1428 ' SALTEA DEPOSITO ENTRADA RECHAZOS
       End If
       '
       EPAC$ = UCase$(EMPREAC$)
1425   Call SELECHO("TADEPOSI/Depósito Entrada Rechazos")
       DSC% = XVALO(VALACT1$("TADEPOSI"))
       If DSC% < 1 Then
         'If InStr(EPAC$, "DOSIVA") > 0 Then
         Call MENSERR(24, "Debe Indicar Depósito\de Entrada de Rechazos.")
         GoTo 1425
         'End If
       End If
     End If
     '
1428 CODICONJ$ = CODEXT$(CI%)
     CODPRO$ = CODICONJ$
     CICONJ% = CI%
     NORFABRI$ = NORFA$
     CACONJU# = CDbl(CANTI + SCRAP)
     'ESTO SE HACE PARA QUE SOLO DESCARGUE MATERIALES EN EL CASO DE SCRAP
     'EL RESTO SE DESCARGA EN LA APROBACION
     'If InStr(1, EMPREAC$, "ARCON") > 0 Then CACONJU# = 0
     VACONTROL% = 0: If PRECOCIE% = 1 Then VACONTROL% = 1
     '
     Call ABRESTOCKS
     NOCLOSE% = 1: NOARRAS% = 1
     '
1426 Call BLOQARCH("RESERVOF")
     Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
     If VACONTROL% < 0 Then
        GoTo 149           ' cancelacion
     End If
     '
     CIERRA% = 0
     If CANTI > CAPENDE# * 0.99 Then
          CIERRA% = 1
        Else
          If MODOGABAL% = 0 Then
            MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
            If MDXX% = 2 Then CIERRA% = 1
          Else
            CIERRA% = CIERRAOF%
          End If
     End If
     '
     Screen.MousePointer = 11
     '
     '\\\OT-05-0730-WAR-10/11/05///
     CAITMO% = 0
     For YX& = 1 To ULTREG&("!MATEROF")
          YY$ = REGLEIDO$("!MATEROF", YX&)
          CI% = CVI(Mid$(YY$, 1, 2))
          CAN# = CVS(Mid$(YY$, 5, 4))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                CAITMO% = CAITMO% + 1
                CIMOVIM%(CAITMO%) = CI%
                CAMOVIM(CAITMO) = (-1) * CAN#
              End If
            End If
          End If
     Next YX&
     If PUEDEMOVI%(DP2%) < 0 Then
        If MODOGABAL% = 1 Then GoTo 91
        GoTo 1420 ' CANCELA SALIDA
     End If
     '\\\OT-05-0730-WAR-FIN///
     
     '\\\OT-08-0995-OD-30/10/08///
     'SE CAMBIA LA FORMA DE GUARDAR EN EL ANOTADOR (OT-8-631) AL USO DE LA RUTINA AGREGAR_ANOTAORFAB (OT-8-995).
     'Y ADEMAS SE AGREGA EL REGISTRO DE LAS CANTIDADES APROBADAS Y RECHAZADAS (OT-8-995)
140  TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Cierre de O/Fabricación", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
     '
     TTXYZ$ = Space$(76)
     Call REPLA(TTXYZ$, "     - Cantidad Aprob.=  " & FORMATNUM(CANTI, "F12.1"), 11, 56)
     Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
     '
     If SCRAP > 0 Then
       TTXYZ$ = Space$(76)
       Call REPLA(TTXYZ$, "     - Cantidad Scrap=   " & FORMATNUM(SCRAP, "F12.1"), 11, 56)
       Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
     End If
     Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
     '\\\OT-08-995-OD-FIN///
     '
'1.- Actualiza saldo pendiente de la Orden de Fabricacion
     Call ABREORFAB
     'With OrdeFabr
    '.FindFirst "IdSubOr = '" & NORFA$ & "'"
     '
     ' ACTUALIZA SOLO PARA SCRAP. LA CANTIDAD OK PASA A LA INSTANCIA DE APROBACION
     If InStr(1, EMPREAC$, "ARCON") > 0 Then
        CANTIAPRO = CANTI
        CANTI = 0
     End If
     '
     SQLX$ = " SELECT * FROM  ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
     Dim ABREORFFFAB As ADODB.Recordset
     Set ABREORFFFAB = New ADODB.Recordset
     ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     On Error Resume Next
     With ABREORFFFAB
        .MoveFirst
        'If .NoMatch = False Then
        If Err < 1 Then
           On Error GoTo 0
          '.Edit
          !CanApro = XVALO(!CanApro) + CANTI
          !CanRech = XVALO(!CanRech) + SCRAP
          If CIERRA% = 1 Then
            !Statuorf = 3         ' MARCA COMO CERRADA
          End If
          '\\\OT-06-0044-WAR-07/02/06/// 'TOTAL O PARCIAL
          !FechCierr = CVDAT(FECHANUM(FECHACIOF$))
          '\\\OT-06-0044-WAR-FIN///
          .Update
        End If
        On Error GoTo 0
     End With
     Set ABREORFFFAB = Nothing
     '
     ' ACTUALIZA 'CANTI' CON LA CANTIDAD APROBADA
     If InStr(1, EMPREAC$, "ARCON") > 0 Then CANTI = CANTIAPRO
     '
     'ASIGNA UBICACION AL PRODUCTO FABRICADO
     If MODOGABAL% = 1 Then
        CAN# = CANTI
'        UB$ = UBISTONUE$(CICONJ%, CAN#)
'        If UB$ <> "" Then
'            OPX% = COMALTER%("Se Ha Asignado Automáticamente La Ubicación '" & UB$ & "'" & "\¿Desea Conservar La Misma?\\Si, Conservar\No, Cancelar")
'            If OPX% = 1 Then
'                Call GRABAUBICACION(CICONJ%, CAN#, UB$, NORFA$)
'            Else
'                UB$ = ""
'            End If
'        Else
'            Call COMUNI("No Se Ha Encontrado Una Ubicación Libre Para El Lote '" & NORFA$ & "'.\Asigne Una En Forma Manual.")
'        End If
     End If
     '
     'CREA LOTE DEL PRODUCTO FABRICADO
     If MODOGABAL% = 1 And MODOMEDITEA% = 1 And ESTRAZABLE%(CICONJ%) = 1 Then
        If CantRegistros&("LOTINGRE", "IDENLOTE='" & NORFA$ & "'") < 1 Then
            Call CREALOTIN(FECHA%, FECHA%, FECHA%, CICONJ%, NORFA$, "", "", CDbl(CANTI))
        End If
     End If
     '
'2.- Genera Movimiento de entrada P.Terminado
143  If CANTI >= 0.05 Then
        DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
        NUORIT% = 0
        If InStr(1, EMPREAC$, "ARCON") > 0 Then
            ' GRABA EN TABLA DE APROBACION
            Call MOVISTOCK_APROF(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, "PESOS", NC%, DEPX%, CANTI, DP2%, "")
        Else
            Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI, , , , , , , TIPOTRAN%)
            CIMOVI%(CICONJ%) = 1
        End If
     End If
     '
     If CICONJ% > 0 Then
        'Master.FindFirst ("CODINT = " & CICONJ%)
        SQLX$ = "UPDATE  MASTER "
        SQLX$ = SQLX$ + " SET DepoEntra = " & DEPX% & ", "
        SQLX$ = SQLX$ + " DepoCon = " & DP2% & " "
        SQLX$ = SQLX$ + " WHERE CODINT = " & CICONJ% & ""
         FLEXCONN.Execute (FILTSQL$(SQLX$))
'        If Master.NoMatch = False Then
'           Master.Edit
'           Master!DepoEntra = DEPX%
'           Master!DepoCon = DP2%
'           Master.Update
'        End If
     End If
     '
'3.- Genera Movimientos de Consumo de Componentes
'    y Cancela Materiales Reservados
     Call ABRESERVA
     FIMATER& = ULTREG&("!MATEROF")
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     For U& = 1 To FIMATER&
       Call TRACE(20, U&, FIMATER&)
       ZZ$ = REGLEIDO$("!MATEROF", U&)
       CI2% = CVI(Left$(ZZ$, 2))
       CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
       '
       If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
          CAN = CAN / 1000
       End If
       '
       USOUNI# = CVD(Mid$(ZZ$, 9, 8))
       RERES& = CVS(Mid$(ZZ$, 17, 4))
       If CI2% > 0 Then
         If CI2% <= NUMAS% Then
           If Abs(CAN) >= 0.05 Then
             'LOTE$ = ""
             ' determina el lote reservado
             CONDIX$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2% & ""
             LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
             '
             CMOV$ = "CF"
             '
             DOPRI$ = NORFA$
             DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
             REFE$ = "CONSUMO " + CODICONJ$
             If CAN > 0 Then
                REFE$ = "RECUPERO " + CODICONJ$
             End If
             If Abs(CAN) >= 0.05 Then
                DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                  If InStr(EPAC$, "YALTRE") > 0 Then
                    Call CARDEPOS
                    For KKII% = 1 To NUSU%
                      If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                        DEPX% = COSU%(KKII%)
                        GoTo 144
                      End If
                    Next KKII%
                  End If
                End If
144             NUORIT% = NUORIT% + 1
                If InStr(1, EMPREAC$, "ARCON") > 0 Then
                    Call MOVISTOCK_APROF(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, "PESOS", NC%, DEPX%, CAN, 0)
                    GoTo 146
                End If
                '
'                If MODOGABAL% = 1 And ESTRAZABLE%(CI2%) = 1 Then
'                    SALDOPEND = (-1) * CAN
'                    Call ARMALILOT(CI2%, DEPX%, , , 1)
'                    For UU& = 1 To ULTREG("!LILOTSE")
'                       If SALDOPEND > 0.05 Then
'                            XX$ = REGLEIDO$("!LILOTSE", UU&)
'                            IDLO$ = TRIM$(Left$(XX$, 16))
'                            CANTAA# = SaldoLote(CI2%, IDLO$, DEPX%)
'                            'CANTAA# = REDON_INDIVIS(CI2%, CANTAA#, "COMPRAS")
'                            If CANTAA# > SALDOPEND Then CANTAA = SALDOPEND
'                            Call MOVISTOK(FECHA%, CI2%, IDLO$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, (-1) * CANTAA)
'                            SALDOPEND = SALDOPEND - CANTAA
'                        End If
'                    Next UU&
'                    GoTo 1440
'                End If
                '
                Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
             End If
             REGMOVI& = ULTREG&("MOVISTO")
             '
1440         SALCON# = (-1) * CAN
             'With Reserva
               '.FindFirst "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2%
             SQLX$ = "SELECT * FROM Reserva "
             SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
             SQLX$ = SQLX$ + " AND Cod_Inte = " & CI2% & " "
             Dim RESETEMP As ADODB.Recordset
             Set RESETEMP = New ADODB.Recordset
             RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
             On Error Resume Next
             With RESETEMP
145            'If .NoMatch = False Then
               .MoveFirst
               If Err < 1 Then
                 On Error GoTo 0
                 '.Edit
                 Do While Not .EOF
                    CARESE# = XVALO(!CantRes)
                    CACOANT# = XVALO(!CantCons)
                    If CARESE# - CACOANT# >= SALCON# Then
                         !CantCons = CACOANT# + SALCON#
                         SALCON# = 0
                       Else
                         SALCON# = SALCON# - (CARESE# - CACOANT#)
                         !CantCons = XVALO(!CantRes)
                    End If
                    .Update
                    .MoveNext
                 Loop
'                 If SALCON# >= 0.05 Then
'                   .FindNext "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2%
'                   GoTo 145
'                 End If
               End If
               On Error GoTo 0
             End With
             Set RESETEMP = Nothing
           End If
         End If
       End If
146  Next U&
     '
     Call CIERRARCH("*.*")
     '
     ' ACA COMIENZA LA GENERACION DE NUMEROS DE SERIE PARA ARCON
     If InStr(1, EMPREAC$, "ARCON") > 0 Then
         If CANTIAPRO > 0.05 Then
           FAMI$ = CODFAMIL$(CICONJ%)
           If FAMI$ <> "" Then
                CANUMSE = CANTIAPRO
                FILA& = 0
                With ASIGNUMSE.MSFlexGrid1
                    .Clear
                    .Row = FILA&: .Col = 0: .ColWidth(.Col) = 1500: .TEXT = "Código"
                    .Row = FILA&: .Col = 1: .ColWidth(.Col) = 3000: .TEXT = "Descripción"
                    .Row = FILA&: .Col = 2: .ColWidth(.Col) = 1500: .TEXT = "Número de Serie"
                    .Row = FILA&: .Col = 3: .ColWidth(.Col) = 1500: .TEXT = "Familia"
                    .Row = FILA&: .Col = 4: .ColWidth(.Col) = 1500: .TEXT = "Orden"
                    .Row = FILA&: .Col = 5: .ColWidth(.Col) = 1500: .TEXT = "Lote"
                    .Row = FILA&: .Col = 6: .ColWidth(.Col) = 1500: .TEXT = "Control N-Serie"
                    .Row = FILA&: .Col = 7: .ColWidth(.Col) = 1500: .TEXT = "Codigo Origen"
                    .Row = FILA&: .Col = 8: .ColWidth(.Col) = 1500: .TEXT = "Lote Origen"
                End With
                '
                CANTREG& = 0
                For U& = 1 To FIMATER&
                  ZZ$ = REGLEIDO$("!MATEROF", U&)
                  CI2% = CVI(Left$(ZZ$, 2))
                  If CantRegistros("TRAZARECE", "COD_INTE=" & CStr(CI2%)) > 0 Then CANTREG& = CANTREG& + 1
                Next U&
                
                If CANTREG& > 1 Or CANTREG& = 0 Then
                    ULTNUMSE& = XVALO(VALOBADA("TRAZARECE", "max(NumeroSerie)", "CFAMILIA='" & FAMI$ & "'"))
                    For III& = 1 To CANUMSE
                        ULTNUMSE& = ULTNUMSE& + 1
                        NSERIE$ = CStr(ULTNUMSE&)
                        '
                        FILA& = FILA& + 1
                        With ASIGNUMSE.MSFlexGrid1
                            On Error Resume Next
                            .Row = FILA&:
                            If Err Then .AddItem "", FILA&
                            On Error GoTo 0
                            .Row = FILA&
                            .Col = 0: .TEXT = CODEXT$(CICONJ%)
                            .Col = 1: .TEXT = DESCRIT0$(CICONJ%)
                            .Col = 2: .TEXT = NSERIE$
                            .Col = 3: .TEXT = FAMI$
                            .Col = 4: .TEXT = "0"
                            .Col = 5: .TEXT = NORFA$
                            .Col = 6: .TEXT = NSERIE$
                            .Col = 7: .TEXT = 0
                            .Col = 8: .TEXT = ""
                        End With
                    Next III&
                Else
                    QTRAZA& = 0: EXCLUYENUMSE$ = ""
                    SQLX$ = "SELECT TOP " & CStr(CANUMSE) & " * FROM FORMULAS INNER JOIN TRAZARECE WITH(NOLOCK) ON CODIELEM=COD_INTE WHERE CODICONJ=" & CStr(CICONJ%) & " AND DESTINO='' AND DESTINOCONSUMO='' AND IDENLOTE<>'' AND NUMEROSERIE>0"
                    If EXCLUYENUMSE$ <> "" Then SQLX$ = SQLX$ & Mid$(EXCLUYENUMSE$, 1, Len(EXCLUYENUMSE$) - 4)
                    SQLX$ = SQLX$ & " ORDER BY NUMEROSERIE ASC"
                    Dim RST01 As ADODB.Recordset
                    Set RST01 = READSET(SQLX$)
                    With RST01
                        Do While Not .EOF
                            CICOMPO% = XVALO(!Cod_Inte)
                            NSERIE$ = SAFETEXT$(!NUMEROSERIE)
                            ORIGEN$ = TRIM$(SAFETEXT$(!IdenLote))
                            ILOTE$ = SAFETEXT$(!IdenLote)
                            FILA& = FILA& + 1
                            With ASIGNUMSE.MSFlexGrid1
                                On Error Resume Next
                                .Row = FILA&:
                                If Err Then .AddItem "", FILA&
                                On Error GoTo 0
                                .Row = FILA&
                                .Col = 0: .TEXT = CODEXT$(CICONJ%)
                                .Col = 1: .TEXT = DESCRIT0$(CICONJ%)
                                .Col = 2: .TEXT = NSERIE$
                                .Col = 3: .TEXT = FAMI$
                                .Col = 4: .TEXT = "0"
                                .Col = 5: .TEXT = NORFA$
                                .Col = 6: .TEXT = NSERIE$
                                .Col = 7: .TEXT = CStr(CICOMPO%)
                                .Col = 8: .TEXT = ILOTE$
                            End With
                            QTRAZA& = QTRAZA& + 1
                            EXCLUYENUMSE$ = EXCLUYENUMSE$ & " NUMEROSERIE<>" & NSERIE$ & " AND "
                            .MoveNext
                        Loop
                    End With
                    On Error Resume Next
                    RST01.Close
                    Set RST01 = Nothing
                    On Error GoTo 0
                    '
                    ULTNUMSE& = XVALO(VALOBADA("TRAZARECE", "max(NumeroSerie)", "CFAMILIA='" & FAMI$ & "'"))
                    For IIX& = (QTRAZA& + 1) To CANUMSE
                        FILA& = FILA& + 1
                        With ASIGNUMSE.MSFlexGrid1
                            On Error Resume Next
                            .Row = FILA&:
                            If Err Then .AddItem "", FILA&
                            On Error GoTo 0
                            .Row = FILA&
                            .Col = 0: .TEXT = CODEXT$(CICONJ%)
                            .Col = 1: .TEXT = DESCRIT0$(CICONJ%)
                            .Col = 2: .TEXT = ""
                            .Col = 3: .TEXT = FAMI$
                            .Col = 4: .TEXT = "0"
                            .Col = 5: .TEXT = NORFA$
                            .Col = 6: .TEXT = ""
                            .Col = 7: .TEXT = 0
                            .Col = 8: .TEXT = ""
                        End With
                     Next IIX&
                End If
           End If
           '
           Screen.MousePointer = 1
           ASIGNUMSE.Show 1
           '
        End If
    End If
    '
    ' FIN GENERACION DE NUMEROS DE SERIE PARA ARCON
    '
'4.- Genera Movimientos de Consumo Correspondientes al SCRAP
     If SCRAP >= 0.05 Then
       '\\\OT-5-739-FG-13/12/05///
       NUORIT% = 0
       '\\\OT-5-739-FG-FIN///
       If DSC% > 0 Then
         CACOSCRA# = SCRAP
         '
         'aqui procesar el caso de que el scrap se declare como PRODUCTO
         'YENDO LUEGO A 125
         'HAY QUE HACER EL MOVISTOK CON CANTIDAD = SCRAP
         '
         '\\\OT-5-739-FG-13/12/05///
         ' PROCESA EL CASO DE QUE EL SCRAP SE DECLARE COMO PRODUCTO
         If Control$("ORDEFABR", "MODOSCRA") = "PRODUCTO" Then
             CI2% = CODINT%(CODICONJ$)
             LOTE$ = NORFA$
             CMOV$ = "SC"
             DOPRI$ = NORFA$
             DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
             REFE$ = "SCRAP " + CODICONJ$
             DEPX% = DSC%: If DSC% = 99 Then DEPX% = DEPOPRE%(CI2%)
             CAN = SCRAP
             '
             If EXISTE%("RECHLI07.EXE") > 0 Then
                  CARECHAZA = SCRAP
                  DOPRI12$ = NORFA$
                  Call PROCERECHA(CI2%, FECHA%, CARECHAZA, DOPRI12$) '\\\OT-06-0191-WAR-09/06/06///
                  HUBORECHAZOS% = 1
                Else
                  Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
             End If
             CIMOVI%(CI2%) = 1
             GoTo 125
         End If
         '\\\OT-5-739-FG-FIN///
         '
         VACONTROL% = 0
         Call BLOQARCH("RESERVOF")
         Call CARCOMPO(CODICONJ$, NORFABRI$, CACOSCRA#, VACONTROL%, 2)
         '
         Call BLOQARCH("MOVISTO")
         FIMATER& = ULTREG&("!MATEROF")
         For U& = 1 To FIMATER&
           Call TRACE(20, U&, FIMATER&)
           ZZ$ = REGLEIDO$("!MATEROF", U&)
           CI2% = CVI(Left$(ZZ$, 2))
           CAN = CVS(Mid$(ZZ$, 5, 4))
           '
           If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
              CAN = CAN / 1000
           End If
           '
           USOUNI# = CVD(Mid$(ZZ$, 9, 8))
           RERES& = CVS(Mid$(ZZ$, 17, 4))
           If CI2% > 0 Then
             If CI2% <= NUMAS% Then
               If Abs(CAN) >= 0.05 Then
                 LOTE$ = NORFA$
                 CMOV$ = "SC"
                 '
                 DOPRI$ = NORFA$
                 DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
                 REFE$ = "SCRAP " + CODICONJ$
                 If Abs(CAN) >= 0.05 Then
                    If EXISTE%("RECHLI07.EXE") > 0 Then
                         CARECHAZA = CAN
                         DOPRI12$ = NORFA$
                         Call PROCERECHA(CI2%, FECHA%, CARECHAZA, DOPRI12$) '\\\OT-06-0191-WAR-09/06/06///
                         HUBORECHAZOS% = 1
                       Else
                         DEPX% = DSC%: If DSC% = 99 Then DEPX% = DEPOPRE%(CI2%)
                         NUORIT% = NUORIT% + 1
                         Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                    End If
                    CIMOVI%(CI2%) = 1
                 End If
                 '
               End If
             End If
           End If
         Next U&
       End If
     End If
     '
125  NOCLOSE% = 0: NOARRAS% = 0
     '
     Call CIERRARCH("MOVISTO")
     If CIERRA% = 1 Then
        ' Cancela Operaciones Programadas
        Call ABREORFAB
        SQLX$ = "DELETE FROM OperFab WHERE IdSubOr = '" & NORFA$ & "'"
        FLEXCONN.Execute SQLX$
        '
        ' Cancela Reservas de Componentes
        Call ABRESERVA
        SQLX$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
        FLEXCONN.Execute SQLX$
        '
        ' Cancela Orden en Indice de Ordenes Pendientes
        Call FILTERDELETE("ECORDEP", 1, NORFA$)
        Call FRESHECHO("ECORDEP")
        '
    End If
    '
    ' GABAL - Marca las Etiquetas Leidas e Imprime Formulario de Cierre
    If MODOGABAL% = 1 And MODOCIE% = 1 Then
         For II% = 0 To CIERORFA.List1.ListCount - 1
             TX$ = CIERORFA.List1.List(II%)
             ORD% = XVALO(Mid$(TX$, 19, 5))
             Call ACTUREGISTRO("ORFAETI", "STATUS", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & ORD%, "3", REGAF&)
             Call ACTUREGISTRO("ORFAETI", "FECHACIERRE", "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & ORD%, FETEXCOR1$(HOY(HOS$)), REGAF&)
         Next II%
         '
         'IMPRIME COMPROBANTE DE CIERRE
         FORIPRE$ = Control$("", "VALCIEOF")
         If FORIPRE$ <> "" Then
            FECHDOC$ = FECHATEX$(HOY(HOS$))
            NUMEDOC$ = ""
            TIPODOC$ = "Cierre de O-Fabricacion"
            DOCUORIG$ = NORFA$
            DESTIDOC% = 0
            '
            CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = NORFA$
            CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHDOC$
                CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
            CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = CODEXT$(CIXI%)
            CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DESCRIT0$(CIXI%)
            CODPARAM$(5) = "UBI-MPRI": DOCPARAM$(5) = UB$
            CAPARDOC% = 5
            '
            CODTABU1$(1) = "REF-MAT1"
            CACOLTAB1% = 1
            CAITAB1% = 0
            '
            For U& = 0 To CIERORFA.List1.ListCount - 1
              XX$ = CIERORFA.List1.List(U&)
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = XX$
            Next U&
            '
            Call PRINTDOC("@" & FORIPRE$)
         End If
         Unload CIERORFA
    End If
    '
149 Screen.MousePointer = 1
    Call LIBARCH("RESERVOF")
    Call CIERRARCH("*.*")
    Call LIBARCH("ECORDEP")
    '
    ' 8.- Imprime Comprobante de cierre de O/F
    LU$ = LUDAT$
    If EXISTE%(LU$ + "OFAP-ARC.FRM") > 0 Then
       FECHDOC$ = FECHACIOF$
       NUMEDOC$ = ""
       TIPODOC$ = "Vale Aprobación Fabricación"
       DOCUORIG$ = NORFA$
       '
       CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODCON$
       CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = DESCON$
       CODPARAM$(3) = "CANTIDAD": DOCPARAM$(3) = TRIM$(Str$(CANTI))
       CAPARDOC% = 3
       '
       DESTIDOC% = 0
       CAITAB1% = 0
       '
       Call CIERRARCH("!OBSERVOF")
       Call BLANARCH("!OBSERVOF")
       VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Cierre O-F Numero " + NORFA$)
       '
       Call PRINTDOC("@OFAP-ARC")   ' Remito de Facturacion
       '
    End If
    '
    '
98  If CANTAPROX < 0.05 Then     ' SOLO EN CASO DE CIERRE MANUAL - NO POR PARTE DIARIO
       If MODOGABAL% = 1 Then GoTo 90
       GoTo 100
    End If
    '
    Exit Sub
    
msglafe:
    Call COMUNI("Error en Seccion " & NOP%)
    Exit Sub
End Sub
Sub ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FFACT%, Optional DEPOENTRA%)
    '
    CAN = CANTI
    If DEPOENTRA% < 1 Then
      DEPOENTRA% = XVALO(Control("", "DEPOCIE"))
      If DEPOENTRA% < 1 Then DEPOENTRA% = 1
    End If
    DEPX% = DEPOENTRA%   ' DEPOSITO DE DESTINO
    FECHA% = FFACT%
    CIXI% = CI1%
    LOTE$ = ""
    CMOV$ = "OF"
    DOPRI$ = NORFA$ ' NUMERO DE ORDEN DE FABRICACION
    DOSEC$ = "IF-" + TRIM$(Mid$(DOPRI$, 4))
    REFE$ = "Fabricación Propia - Aut_x_Fact."
    VUNI# = 0
    MONE$ = "$"
    Call MOVISTOK(FECHA%, CIXI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
    '
End Sub

Sub CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FFACT%, Optional DEPOCONSU%)
    '
    NORFAX$ = NORFA$
    If DEPOCONSU% < 1 Then
      DP2% = DEPENART%(CODI1%)
      If DP2% < 1 Then DP2% = 1
      DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CICOMPO%)
    End If
    
    FIMATER& = ULTREG&("!MATEROF")
    '
    For U& = 1 To FIMATER&
       Call TRACE(20, U&, FIMATER&)
       ZZ$ = REGLEIDO$("!MATEROF", U&)
       CICOMPO% = CVI(Left$(ZZ$, 2))
       CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
       If TRIM$(UCase$(UNIMED$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
       '
       FECHA% = FFACT%
       Call CONSUMECOMPOFAC(CICOMPO%, NORFAX$, CANCONSUM, FECHA%, DEPX%)
    Next U&
    '
End Sub
Sub CONSUMECOMPOFAC(CI2%, NORFAX$, CAN, FECONSUM%, Optional DP2%)
    '
       If CI2% > 0 Then
         If CI2% <= NUMAS% Then
           If Abs(CAN) >= 0.005 Then
             'LOTE$ = ""
             ' determina el lote reservado
             CONDIX$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2% & ""
             LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
             CODICONJ$ = CODEXT$(CODIFAB%(NORFAX$))
             '
             CMOV$ = "CF"
             '
             DOPRI$ = NORFAX$
             DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
             REFE$ = "CONSUMO " + CODICONJ$
             If Abs(CAN) >= 0.005 Then
                DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                NUORIT% = NUORIT% + 1
                '
                Call MOVISTOK(FECONSUM%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DP2%, CAN * (-1))
             End If
           End If
         End If
       End If
End Sub

Sub ACTURESERVA(CI2%, NORFA$, CANTI)

             SQLX$ = "SELECT * FROM Reserva "
             SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
             SQLX$ = SQLX$ + " AND Cod_Inte = " & CI2% & " "
             Dim RESETEMP As ADODB.Recordset
             Set RESETEMP = New ADODB.Recordset
             RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             On Error Resume Next
             With RESETEMP
                 Do While Not .EOF
                    CARESE# = XVALO(!CantRes)
                    CACOANT# = XVALO(!CantCons)
                    If CARESE# - CACOANT# >= SALCON# Then
                         !CantCons = CACOANT# + SALCON#
                         SALCON# = 0
                       Else
                         SALCON# = SALCON# - (CARESE# - CACOANT#)
                         !CantCons = XVALO(!CantRes)
                    End If
                    .Update
                    .MoveNext
                 Loop
               
               On Error GoTo 0
             End With
             RESETEMP.Close
             Set RESETEMP = Nothing
 

End Sub
Sub ANOTAENORFA_FACTURADO(CI1%, NORFA$, CANTI, FFACT%)
    '

    NORFAX$ = TRIM$(NORFA$)
    TTXYZ$ = Space$(76)
    FE% = FFACT%
    Call REPLA(TTXYZ$, FECHATEX$(FE%), 1, 8)
    Call REPLA(TTXYZ$, "Cierre de O/Fabricación Items Facturados", 11, 56)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call AGREGA_ANOTAORFAB(NORFAX$, TTXYZ$)
    '
    TTXYZ$ = Space$(76)
    Call REPLA(TTXYZ$, " - Cantidad Aprob.= " & CANTI, 11, 56)
    Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
    Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
    '
End Sub
Sub ACTUALIZAORFA(CI1%, NORFAX$, CANTI, FFACT%)

          CONDI$ = "CODIMOVI = 'OF' AND DOPRILAR = '" & NORFAX$ & "'"
          TOFABRI = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE)", CONDI$, "NOMESS"))
          TOFABRI = TOFABRI - XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
          If TOFABRI < 0 Then TOFABRI = 0
          Call ACTUREGISTRO("ORDEFABR", "CanApro", "IdSubOr = '" & NORFAX$ & "'", TOFABRI, REGAFE&, "NOMESS")
       '  CIERRE DE LA ORDEN DE FABRICACION
          CANPROYE& = XVALO(VALOBADA("ORDEFABR", "CanProy", "IdSubOr = '" & NORFAX$ & "'"))
          CANREALE& = XVALO(VALOBADA("ORDEFABR", "CanApro", "IdSubOr = '" & NORFAX$ & "'"))
          '
          If CANREALE& >= CANPROYE& Then
              CONDI$ = "IdSubOr = '" & NORFAX$ & "'"
              FECHAHOY = FETEXCOR1$(FFACT%)
              Call ACTUREGISTRO("ORDEFABR", "StatuOrf", CONDI$, 3, REGAF&, "NOMESS")
              Call ACTUREGISTRO("ORDEFABR", "FechCierr", CONDI$, FECHAHOY, REGAF&, "NOMESS")
              ' Cancela Operaciones Programadas
              Call BORRAREGISTROS("OperFab", CONDI$, RAFE&, "NOMESS")
              ' Cancela Reservas de Componentes
              Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
              ' Cancela Orden en Indice de Ordenes Pendientes
              NORFAX$ = AJUSTI$(NORFAX$, 12)
              Call FILTERDELETE("ECORDEP", 1, NORFAX$)
              Call FRESHECHO("ECORDEP")
          End If
          '
          Call CIERRARCH("*.*")
          Call LIBARCH("ECORDEP")
          Call CIERRARCH("*.*")

End Sub
Sub MOVISTOCK_APROF(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI, DEPOCON%, Optional CLA0$)
   '
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   Call COPYSTRU("MOVISTO", "ORFAPEN")
   '
   SMP = Screen.MousePointer
   UREMAS& = ULTREG&("MASTER")
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   Entrada# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      Entrada# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   DepoMovi% = DEPO%
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(DepoMovi%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(Entrada#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   CLA1$ = CLA0$
   If CLA1$ = "" Then
30      CLA1$ = RANDSTRING$(-10)
        RFF$ = RECFILT$("ORFAPEN", 17, CLA1$)
        If RFF$ <> "" Then GoTo 30
   End If
   Call REPLA(RECORD$, CLA1$, 112, 10)
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, MKI$(DEPOCON%), 127, 2)
   '
   Dim BUGRA As String * 140
   ' ESTO ES PARA QUE NO SE CHOQUE CON LA APROBACION DE CIERRE DE OF.
   Call BLOQARCH("ORFAPEN")
   Call REGAPP("ORFAPEN", RECORD$)
   Call LIBARCH("ORFAPEN")
   REMOVIST& = ULTREG&("ORFAPEN")
   Call CIERRARCH("ORFAPEN")
   CLA0$ = CLA1$
   '
End Sub

'
Sub ORFANU()
    '
    MODOGABAL% = 0
    If InStr(1, EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
    If MODOGABAL% = 1 Then
        OPXX% = COMALTER%("Elija la Modalidad de Anulación\\Código de Barras\Manual")
        If OPXX% < 1 Or OPXX% > 2 Then Exit Sub
        If OPXX% = 2 Then GoTo 50
        '
        ANUORFA.Caption = "Anulación de O-Fabricación"
        ANUORFA.Label2 = "0"
        ANUORFA.Show 1
        Exit Sub
    End If
    '
50  HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE FABRICACION"

100  NROOF$ = ""

1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("ECORDEF", "ECORANU")
       '
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa, Referen, IdSubOr, CanProy, Cod_Inte FROM ORDEFABR "
       SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "
       SQLX$ = SQLX$ + "AND CanApro < CanProy "
       SQLX$ = SQLX$ + "order by IdSubOr ASC"
       'Set ORFANUTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       Dim ORFANUTMP As ADODB.Recordset
       Set ORFANUTMP = New ADODB.Recordset
       ORFANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
       '
       On Error Resume Next
       With ORFANUTMP
         .MoveFirst
         If Err Then
           Call MENSERR(24, "No Hay O/F's que Puedan Anularse.")
           Screen.MousePointer = 1
           On Error GoTo 0
           Exit Sub
         End If
         On Error GoTo 0
         '
         JJ& = 0
         Do While Not .EOF
             ZZ$ = REGBLAN$("!ECORANU")
             Call REPLA(YY$, SAFETEXT$(!IDSUBOR), 1, 12)
             Call REPLA(YY$, SAFETEXT$(!REFEREN), 13, 52)
             JJ& = JJ& + 1
             Call GRAREG("!ECORANU", YY$, JJ&)
         .MoveNext
         Loop
       End With
       Set ORFANUTMP = Nothing
       Call SETULTREG("!ECORANU", JJ&)
       Call CIERRARCH("!ECORANU")
       Screen.MousePointer = 1
     End If
     '
     Screen.MousePointer = 1
     '
1405 Call FRESHECHO("!ECORANU")
     If ULTREG&("!ECORANU") < 1 Then INDI% = 0: Exit Sub
1407 Call SELECHO("!ECORANU/Ordenes de Trabajo para Anulación")
     '
1410 NORFA$ = AJUSTI$(VALACT1$("!ECORANU"), 12)
     If TRIM$(NORFA$) = "" Then
         INDI% = 0
         Exit Sub
     End If
     '
     If MODOGABAL% < 1 Then
        Res% = COMALTER%("¿Esta Seguro que Desea Anular '" & TRIM$(NORFA$) & "'?\\No, Cancelar\Si, Confirmar")
        If Res% <> 2 Then Exit Sub
        '
     Else
        ANUORFA.Hide
        ANUORFA.Label2 = "0"
        ANUORFA.Text2 = TRIM$(REPLACAR(NORFA$, "-", ""))
        Call ANUORFA.Text2_KeyPress(13)
        Unload ANUORFA
        Call COMUNI("Proceso de Anulación Finalizado")
        Exit Sub
        '
     End If
     '
     ' VERIFICA SI EXISTE NOTA DE PEDIDO GENERADA PARA LA S/F A ANULAR
     CONDI$ = "ANEX_OCOM='" & NORFA$ & "' AND STAT_OCOM < 9"
     If CantRegistros("OCOMENCA", CONDI$) > 0 Then
         Call COMUNI("Imposible Anular S/F, Posee Nota de Pedido de Materiales En Curso\\Anule la Nota de Pedido y Vuelva a Intentarlo")
         GoTo 1407 ' VUELVE A MOSTRAR LA LISTA DE S/F
     End If
        
     ' VERIFICA SI EXISTE MATERIAL RECIBIDO Y ASIGNADO PARA LA S/F A ANULAR
     CONDI$ = "DESTINO='" & NORFA$ & "' AND Kilo_Asig > 0"
     If CantRegistros("BALAREQMO", CONDI$) > 0 Then
         Call COMUNI("Imposible Anular S/F, Posee Material Asignado.\Desasigne el Material desde la Opcion Des-Asignar Lotes Asignados a S/F\y Vuelva a Intentarlo")
         GoTo 1407 ' VUELVE A MOSTRAR LA LISTA DE S/F
     End If
    
     ' VERIFICA SI EXISTE MATERIAL RECIBIDO Y ASIGNADO PARA LA S/F EN LOTINGRE
     CONDI$ = "COD_DESTI='" & NORFA$ & " ' "
     If CantRegistros("LOTINGRE", CONDI$) > 0 Then
         Call COMUNI("Imposible Anular S/F, Hay Lotes / Bultos \Recibidos y Pre-asignados a esta  S/F.\  \Desasigne el Material desde la Opcion Des-Asignacion de Materiales \Y Vuelva a Intentarlo")
         GoTo 1407 ' VUELVE A MOSTRAR LA LISTA DE S/F
     End If
     '
     
     'VALIDACION DE NUMEROS DE SERIE
     BORRARNUSE% = 0
     CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
     CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
     If ESTRAZABLE%(CIXI%) = 2 Then
          CONDI1$ = "DEPONUMSE >= 0 AND IDSUBOR = '" & NORFA$ & "' And Cod_Inte = " & CIXI%
          CNTEQUI = CantRegistros&("TANUMSE", CONDI1$, "NOMESS")
          CANAPROB = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$, "NOMESS"))
          If CNTEQUI <> CANAPROB Then
            Screen.MousePointer = 1
            Call COMUNI("Imposible Anular O/F Cantidad de Equipos Ingresados (Deponumse <> -1)\Difiere de Cantidad de Equipos Aprobados por Control Final")
             GoTo 1407 ' VUELVE A MOSTRAR LA LISTA DE S/F
          End If
        '
        'VALIDO SI HAY EQUIPOS EN DEPOSITOS EN - 1  PRESENTA MENSAJE PARA CONFIRMAR ANULACION
        CONDI2$ = "DEPONUMSE < 0 AND IDSUBOR = '" & NORFA$ & "' And Cod_Inte = " & CIXI%
        EQMENUNO = CantRegistros&("TANUMSE", CONDI2$, "NOMESS")
        If EQMENUNO > 0 Then
           Screen.MousePointer = 1
           OPX% = COMALTER%("Se han Encontrado Equipos (N/Serie) Pendientes de Cierre.\Desea Eliminarlos ?\\ Sí, Eliminar\No, Continuar")
           If OPX% = 1 Then
              BORRARNUSE% = 1
           End If
        End If
     End If

     'FIN DE VALIDACION DE NUMEROS DE SERIE
     '
     'VALIDAR SI HAY MATERIAL TRANSFERIDO SIN CONSUMIR
     DepOrig% = XVALO(Control("", "DEPODEST")) ' DEPOSITO TALLER
     If DepOrig% > 0 Then   ' SI TIENE DEPOSITO DE TALLER CONFIGURADO
           Screen.MousePointer = 11
           ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD INGRESADA POR ARTICULO SEGUN LA O.F.
           Dim MATRANSFE#(32767)
           For II% = 1 To NUMAS%
             MATRANSFE#(II%) = 0
           Next II%
           
           SQLX$ = "SELECT * FROM MOVISTO  WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NORFA$ & "' "
           SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
           SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
           Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
           With MoviTemp
             Do While Not .EOF
               CI% = XVALO(!Cod_Inte)
               CAN = XVALO(!CantIngre) - XVALO(!Cantsalid)
               '
               MATRANSFE#(CI%) = MATRANSFE#(CI%) + CAN
               .MoveNext
             Loop
           End With
           MoviTemp.Close
           Set MoviTemp = Nothing
           '
           ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD SALIDA POR ARTICULO SEGUN LA O.F.
           SQLX$ = "SELECT * FROM MOVISTO  WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF'"
           SQLX$ = SQLX$ + " AND DOPRICOR = '" & NORFA$ & "'"
           SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
           Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
           With MoviTemp
             Do While Not .EOF
               CI% = XVALO(!Cod_Inte)
               CAN = XVALO(!Cantsalid) - XVALO(!CantIngre)
               '
               MATRANSFE#(CI%) = MATRANSFE#(CI%) - CAN
              .MoveNext
             Loop
           End With
           MoviTemp.Close
           Set MoviTemp = Nothing
           
           For I% = 1 To NUMAS%
                CI% = I%
                UM1$ = UNIMED$(CI%)
                CAN1 = MATRANSFE#(CI%)
                If MATRANSFE#(CI%) > 0.000005 Then
                    Call COMUNI("La O/F No puede ser Anulada porque tiene Materiales en Deposito Taller")
                    Screen.MousePointer = 1
                    GoTo 1407
                End If
           Next I%

     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
     '
     'Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim CARGAOBJTMP As ADODB.Recordset
     Set CARGAOBJTMP = New ADODB.Recordset
     CARGAOBJTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     NROOFL = XVALO(!NUMEORFA)
     Call REPLA(VDEF$, MKS$(NROOFL), 1, 4)
     Call REPLA(VDEF$, SAFETEXT$(!REFEREN), 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, SAFETEXT$(!Unidad), 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, SAFETEXT$(!DESTINO), 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(!Statuorf), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     Set CARGAOBJTMP = Nothing
     '
1412 YY$ = OBJPLA$("ORFALAN/Anulacion de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     '
     'ANULACION OF
     '
     Call ABREORFAB
     Call ABRESERVA

     SQLX$ = "SELECT Statuorf, IdSubOr FROM OrdeFabr "
     SQLX$ = SQLX$ + "WHERE NumeOrFa = " & NROOFL & " "
     SQLX$ = SQLX$ + " and IDSUBOR= '" & NORFA$ & "'"
     Dim ESTADORDEN As ADODB.Recordset
     Set ESTADORDEN = New ADODB.Recordset
     ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     'With OrdeFabr
     With ESTADORDEN
          '.Edit
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
             Do While Not .EOF
              !Statuorf = 9
              .Update
              IDSUB$ = SAFETEXT$(!IDSUBOR)
              '
              SQLX$ = "SELECT * FROM Reserva "
              SQLX$ = SQLX$ + "WHERE IDSUBOR= '" & IDSUB$ & "' "
              Dim CANTIRES As ADODB.Recordset
              Set CANTIRES = New ADODB.Recordset
1420          On Error Resume Next
              CANTIRES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              '
              If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 1420
              End If
              On Error GoTo 0
              '
              With CANTIRES
                  On Error Resume Next
                  .MoveFirst
                  If Err < 1 Then
                    Do While Not .EOF
                      !CantRes = XVALO(!CantCons)
                      !Saldo_Entre = XVALO(!CantRes) - XVALO(!CantCons)
                      .Update
                      '
                      'ACTUALIZO EL SALDO DEL LOTE
                      CI% = XVALO(!Cod_Inte)
                      LOTE$ = SAFETEXT$(!IdenLote)
                      Call ACSALOTE(CI%, LOTE$, "NOMESS")
                      '
                      .MoveNext
                    Loop
                  End If
                  On Error GoTo 0
              End With
              Set CANTIRES = Nothing
              'ANULACION DE OPERACIONES DE LA O-FABRICACION
              SQLX$ = "DELETE FROM OperFab "
              SQLX$ = SQLX$ + "WHERE IdSubOr = '" & IDSUB$ & "' "
              FLEXCONN.Execute (FILTSQL$(SQLX$))
              .MoveNext
             Loop
          End If
          On Error GoTo 0
     End With
     Set ESTADORDEN = Nothing
     '
     'Borra Registros de BALAREQMO
     NORFA$ = IDSUB$
     Dim Reque As ADODB.Recordset
     Set Reque = New ADODB.Recordset
     SQLX$ = "SELECT * FROM BALAREQMO WHERE DESTINO ='" & NORFA$ & "'"
     Set Reque = FLEXCONN.Execute(FILTSQL$(SQLX$))
     If Reque.BOF = True And Reque.EOF = True Then GoTo 1500
     'SE MODIFICA EL GOTO PARA QUE CONTINUE CON EL FINAL DE LA ANULACION
     With Reque
        .MoveFirst
        Do While Not (.EOF)
           CINTER% = XVALO(!Cod_Inte)
           LOTE$ = SAFETEXT(!IdenLote)
           RESER = XVALO(!Kilo_Asig)
           If RESER > 0 Then
              NUESQLX$ = "DESTINO ='" & NORFA$ & "' "
              NUESQLX$ = NUESQLX$ & " AND COD_INTE = " & CINTER%
              NUESQLX$ = NUESQLX$ & " AND IDENLOTE='" & LOTE$ & "'"
              Call BORRAREGISTROS("BALAREQMO", NUESQLX$, REG&)
              'Actualizo el LOTINGRE segun lo desasignado del balareqmo
              CONDI$ = "COD_INTE = " & CINTER% & " AND IDENLOTE = '" & LOTE$ & "' AND COD_DESTI = '" & NORFA$ & "'"
              Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI$, "", REGG&, "NOMESS")
              Call ACSALOTE(CINTER%, LOTE$, "NOMESS")
           End If
        .MoveNext
        Loop
     End With
     '
     'ACTUALIZO LOS LOTES QUE INGRESARON CON DESTINO DE LA ORDEN DE PRODUCCION
     'PERO QUE EL BALANCE DE REQUERIMENTOS NO LOS RESERVO
     CONDI$ = "COD_DESTI = '" & NORFA$ & "'"
     Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI$, "", REGG&, "NOMESS")
     '
1500 If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
        Call ANULARES(NORFA$)
     End If
     '
     'BANDERA PARA BORRAR LOS REGISTROS EN -1 DE LOS EQUIPOS
     If BORRARNUSE% > 0 Then
        CONDI$ = "DEPONUMSE < 0 AND IDSUBOR = '" & NORFA$ & "' And Cod_Inte = " & CIXI%
        Call BORRAREGISTROS("TANUMSE", CONDI$, REGA&, "NOMESS")
     End If
     '
     'SI LA ORDEN DE COMPRA TIENE ASIGNADO EL NUMERO DE LA ORDEN A ANULAR LA ANULA TAMBIEN.
     CONDI$ = "Doc_Orig =  '" & NORFA$ & "' AND N_OrServ  > 0 "
     Call ACTUREGISTRO("OCOMDETA", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
     '
     JJ& = 0
     For U& = 1 To ULTREG&("!ECORANU")
       ZZ$ = REGLEIDO$("!ECORANU", U&)
       If TRIM$(Left$(ZZ$, 12)) <> TRIM$(NORFA$) Then
         JJ& = JJ& + 1
         Call GRAREG("!ECORANU", ZZ$, JJ&)
       End If
     Next U&
     Call SETULTREG("!ECORANU", JJ&)
     '
     Call CIERRARCH("*.*")
     Call COMUNI("La Orden de Fabricación " & NORFA$ & "  \ha sido Anulada.")
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     GoTo 1405
     '
End Sub
'

Sub LANORF(NROOFL&, FLAN%)
 '
 Screen.MousePointer = 11
 APROLAN% = 0
  Call ABREORFAB
     'With OrdeFabr
       '.FindFirst "NumeOrFa = " & NROOFL
        SQLX$ = "SELECT * FROM OrdeFabr "
        SQLX$ = SQLX$ + " WHERE NumeOrFa = " & NROOFL & ""
        Dim ESTADORDEN As ADODB.Recordset
        Set ESTADORDEN = New ADODB.Recordset
        ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
1415   'If .NoMatch = False Then
         '.Edit
     With ESTADORDEN
          On Error Resume Next
            'If .NoMatch = False Then
              '.Edit
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              !Statuorf = 1
              !FechLan = CVDAT(FLAN%)
              .Update
               IDSUB$ = SAFETEXT$(!IDSUBOR)
        
              SQLX$ = "SELECT StatOper FROM OperFab "
              SQLX$ = SQLX$ + " WHERE IdSubOr = '" & IDSUB$ & "'"
              Dim ESTADOPER As ADODB.Recordset
              Set ESTADOPER = New ADODB.Recordset
              ESTADOPER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              'With OperFab
              With ESTADOPER
                 '.FindFirst "IdSubOr = '" & IDSUB$ & "'"
1416             'If .NoMatch = False Then
                   '.Edit
                   On Error Resume Next
                   .MoveFirst
                   If Err < 1 Then
                     Do While Not .EOF
                      !StatOper = 1
                      .Update
                      '.FindNext "IdSubOr = '" & IDSUB$ & "'"
                      'GoTo 1416
                      .MoveNext
                     Loop
                   End If
                   On Error GoTo 0
              End With
              Set ESTADOPER = Nothing
              '.FindNext "NumeOrFa = " & NROOFL
              .MoveNext
              'GoTo 1415
            Loop
          End If
          On Error GoTo 0
     End With
     Set ESTADORDEN = Nothing
     '
     Screen.MousePointer = 1
     APROLAN% = 1
     '
End Sub
'
Function HAYDISTO%(CIX%, NORFAI&, CACONJU)
     '
     HDS% = 0
     If CHECOMPO% = 0 Then
        CHECOMPO% = (-1)
        If Control$("", "CHECOMPO") = "SI" Then
          If Control$("", "PRECOLAN") <> "SI" Then
            CHECOMPO% = 1
            ' PRESENTA MENSAJE DE ADVERTENCIA SI
            '   VERIFICA COMPONENTES Y NO LOS PRESENTA AL LANZAMIENTO
          End If
        End If
     End If
     '
     If CHECOMPO% <> 1 Then
        HAYDISTO% = 1
        Exit Function
     End If
     '
     If CIX% < 1 Or CIX% > NUMAS% Then
        HAYDISTO% = 1
        Exit Function
     End If
     '
     COD$ = CODEXT$(CIX%)
     CODCON$ = COD$
     Call LEEEXPLOMUL(COD$, 1, 1)
     For II% = 1 To CAIT%
        CI2% = CIJ%(II%)
        CAN = CACONJU * USOI(II%)
        STDIS = STOCKDIS(CI2%)
        STROF = STORESOF(CI2%, NORFAI&)
        If CAN > STDIS + STROF Then
           MENSA$ = "El Stock Disponible (" + TRIM$(CODEXT$(CI2%)) + ") es Insuficiente\para Lanzar Esta Orden de Fabricación."
           If COMALTER%(MENSA$ + "\\Lanzar Igual\Cancelar") = 1 Then
              HDS% = 1
           End If
           GoTo 99
        End If
     Next II%
     HDS% = 1
     '
99   HAYDISTO% = HDS%
     '
End Function
'

Sub GENRESER(NROFLLA$)
     '
10   NROOF$ = NROFLLA$
     HO% = HOY(HOS$)
     '
     If PRECOLAN% = 0 Then
        PRECOLAN% = (-1)
        If Control$("", "PRECOLAN") = "SI" Then
           PRECOLAN% = 1
        End If
     End If
     '
     If TRIM$(NROOF$) <> "" Then
        NORFA$ = AJUSTI$(NROOF$, 12)
        If TRIM$(ECOARCH$("ECORDEP", NORFA$)) <> "" Then
            GoTo 120
        End If
     End If
     NROOF$ = ""
     '
100  Call ACTECOR(1)     ' parametro 1 para seleccionar solo llamzadas y suspendidas
     Screen.MousePointer = 1
     NROOF$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
110  Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
120  Call ABREORFAB
     CONDI$ = "IdSubOr = '" & NROOF$ & "'"
     CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
     CANPROYE# = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$, "NOMESS"))
     CANAPROB# = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$, "NOMESS"))
     CAPEN# = CANPROYE# - CANAPROB#
     If CAPEN# < 0 Then CAPEN# = 0
     '
     ' MODIFICADA LA RE-GENERACION DE RESERVAS POR O/F PARA REALIZAR EL CÁLCULO POR
     ' EL TOTAL PROYECTADO DE LA O/F Y LUEGO DESCONTAR LO CONSUMIDO DE ESA MISMA
     ' ORDEN DE FABRIVCACION - EPB - 16 JUNIO 2009
          '
1410 If CAPEN# < 0.1 Then
       Call MENSERR(24, "No hay Pendiente para la O/F '" + TRIM$(NORFA$) + "'.")
     End If
     '
     CICONJ% = CIXI%
     CODICONJ$ = CODEXT$(CIXI%)
     NORFABRI$ = NORFA$
     CACONJU# = CANPROYE#    ' CAPEN#
     VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
     '
1412 Call BLOQARCH("RESERVOF")
     Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 3)
     If VACONTROL% < 0 Then
        GoTo 1499
     End If
     '
     Call ABRESERVA
     SQLX$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
     FLEXCONN.Execute SQLX$
     '
1415 Dim RESERVVVA As ADODB.Recordset
     Set RESERVVVA = New ADODB.Recordset
     SQLX$ = " SELECT * FROM Reserva where IdSubOr = '" & NORFA$ & "' "
     RESERVVVA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     NORIT% = 0
     '
     For U& = 1 To ULTREG&("!MATEROF")
        ZZ$ = REGLEIDO$("!MATEROF", U&)
        CIIX% = CVI(Left$(ZZ$, 2))
        CANTX# = CVS(Mid$(ZZ$, 5, 4))
        USOII = 0: If Abs(CANPROYE#) >= 0.05 Then USOII = CANTX# / CANPROYE#
        '
        If Abs(CANTX#) >= 0.05 Then
          With RESERVVVA
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !Cod_Inte = CIIX%
            !COD_EXTE = CODEXT$(CIIX%)
            !DESCRIP = DESCRIT0$(CIIX%)
            !IDSUBOR = NORFA$
            !FechRes = Now
            !CantRes = CANTX#
               CANCONSUM = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", "Cod_Inte = " & CIIX% & " AND CodiMovi = 'CF' AND DoPriLar = '" & NORFA$ & "'", "NOMESS"))
            !CantCons = CANCONSUM
               SALDOEN = CANTX# - CANCONSUM
               If SALDOEN < 0 Then SALDOEN = 0
            !Saldo_Entre = SALDOEN
            !IdenLote = " "
            !Deposito = 0
            !UsoUnit = USOII
               NORIT% = NORIT% + 1
            !NuOrItem = NORIT%
            !NUME_SOLINT = 0
            .Update
          End With
        End If
     Next U&
     RESERVVVA.Close
     Set RESERVVVA = Nothing
     '
1499 Call LIBARCH("RESERVOF")
     If TRIM$(NROFLLA$) = "" Then
        GoTo 100
     End If
     '
End Sub
'

Sub REGENRESER()
   '
   If DERACCE%("REGERESE", "Re-Generación de Reservas según O/F") < 2 Then
      Exit Sub
   End If
   '
   Dim NOFACTI$(4096)
   CANOF% = 0: j& = 0
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("ORDEFABR")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", I&)
      ST% = Asc(Mid$(XX$, 108, 1))
      If ST% < 3 Then
         CATOT# = CVD(Mid$(XX$, 89, 8))
         CAPRA# = CVD(Mid$(XX$, 119, 8))
         CAPEN# = CATOT# - CAPRA#
         If CAPEN# > 0.05 Then
            NOFANT$ = TRIM$(Mid$(XX$, 69, 12))
            If TRIM$(NOFANT$) <> "" Then
               CANOF% = CANOF% + 1
               NOFACTI$(CANOF%) = NOFANT$
               X$ = REGLEIDO$("ECORDEF", I&)
               j& = j& + 1
               Call GRAREG("ECORDEP", X$, j&)
            End If
         End If
      End If
   Next I&
   Call SETULTREG("RESERVA", 0)
   Call CIERRARCH("RESERVA")
   Call SETULTREG("ECORDEP", j&)
   Call CIERRARCH("ECORDEP")
   '
   FIN& = CANOF%
   j& = 0
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      '
      For I1& = I& + 1 To FIN&                   ' ORDENAR
         If NOFACTI$(I&) > NOFACTI$(I1&) Then
            NOFAC$ = NOFACTI$(I&)
            NOFACTI$(I&) = NOFACTI$(I1&)
            NOFACTI$(I1&) = NOFAC$
         End If
      Next I1&
      '
      Call GENRESER(NOFACTI$(I&))
   Next I&
   Call CIERRARCH("RESERVA")
   Call CIERRARCH("*.*")
   '
   Call ACTECOR(0)
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
End Sub
'

Sub CHECKRES()
   '
   If DERACCE%("CHECKRES", "Control de Reservas según O/Fs en Curso") < 2 Then
      Exit Sub
   End If
   '
   If COMALTER%("Para Ejecutar Este Proceso\TODOS LOS USUARIOS deben SALIR DEL SISTEMA\\Cancelar\Continuar") <> 2 Then
      Exit Sub
   End If
   '
   Call ACTECOR(0)    ' SELECCIONA PROYECTADAS, LANZADAS Y SUSPENDIDAS
   '
   Screen.MousePointer = 11
   Call ABRESERVA
   '
   SQLX$ = "SELECT * FROM Reserva "
   Dim CANRES As ADODB.Recordset
   Set CANRES = New ADODB.Recordset
   CANRES.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   NOFAN$ = ""
   'With Reserva
  With CANRES
     On Error Resume Next
     .MoveFirst
25   If Err < 1 Then
       NOFARES$ = SAFETEXT$(!IDSUBOR)
       CARES# = XVALO(!CantRes)
       CACON# = XVALO(!CantCons)
       CASAL# = CARES# - CACON#
       !Saldo_Entre = CASAL#
       .Update
       '
       If CASAL# < 0.005 Or ECOARCH$("ECORDEP", NOFARES$) = "" Then
         .Delete
         If NOFARES$ <> NOFAN$ Then
           NOFAN$ = NOFARES$
           Call COMUNI("Canceladas Reservas para " + NOFARES$)
         End If
       End If
       If Not .EOF Then
         .MoveNext
         On Error Resume Next
         GoTo 25
       End If
     End If
     On Error GoTo 0
   End With
   'Reserva.Close
   Set CANRES = Nothing
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub ANURESER(NROFLLA$)
     '
     NROOF$ = NROFLLA$
     HO% = HOY(HOS$)
     '
     If TRIM$(NROOF$) = "" Then
       If DERACCE%("ANURESE", "Anulación de Reservas de Materiales") < 2 Then
         Exit Sub
       End If
     End If
     '
     If TRIM$(NROOF$) <> "" Then
        NORFA$ = AJUSTI$(NROOF$, 12)
        If TRIM$(ECOARCH$("ECORDEF", NORFA$)) <> "" Then
            GoTo 120
        End If
     End If
     NROOF$ = ""
     '
100  Screen.MousePointer = 1
     Call SELECHO("ECORDEF/Anulacion de Reservas segun O/F")
     NROOF$ = TRIM$(VALACT1$("ECORDEF"))
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
110  NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEF", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
     If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
        Call ANULARES(NORFA$)
        GoTo 130
     End If
     '
120  Call ABRESERVA
     'With Reserva
125    '.FindFirst "IdSubOr = '" & NORFA$ & "'"
       'If .NoMatch = False Then
         '.Delete
         SQLX$ = "DELETE FROM RESERVA "
         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
          FLEXCONN.Execute (FILTSQL$(SQLX$))
         
         'GoTo 125
       'End If
     'End With
     '
130  Screen.MousePointer = 1
     '
     If TRIM$(NROFLLA$) = "" Then
        Call COMUNI("Anulación de Reservas Correspondientes a '" + TRIM$(NORFA$) + "' Completada.")
        GoTo 100
     End If
     '
End Sub
'

Sub ACTECOR(MODOI%)
    '
    ' con MODOI% = 2 selecciona solo activas lanzadas (status =1
    Screen.MousePointer = 11
    Call COPYSTRU("ECORDEF", "ECORDEP")
    j& = 0
    '
    Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr, StatuOrf, CanApro, CanProy, REFEREN FROM ORDEFABR WITH(NOLOCK)"
    If MODOI% = 2 Then
        ' SQLX$ = SQLX$ + " WHERE StatuOrf =1 "
        SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "   ' para que no este anulada ni cumplida
        SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
      Else
        SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
    End If
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "NO Existen Ordenes de Fabricación\Correspondientes a Criterio de Selección.\  \Imposible Continuar la Carga.")
        GoTo 90
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REGBLAN("ECORDEP")
        '
        Call REPLA(Y$, SAFETEXT$(!IDSUBOR), 1, 12)
        Call REPLA(Y$, SAFETEXT$(!REFEREN), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
    Set ORDEFABTMP = Nothing
    
    '
90  Call SETULTREG("ECORDEP", JJ&)
    Call CIERRARCH("ECORDEP")
    Call FRESHECHO("ECORDEP")
    Screen.MousePointer = 1
    '
End Sub


Sub TOTEQUIP()
    '
    Dim EQQ$(512), TPREP(512), TOPER(512)
    CAEQUI& = ULTREG&("PADMAQ")
    If CAEQUI& > 512 Then
        Call COMUNI("Demasiados Equipos.\Consulte Soporte de Sistemas.")
        Call FINAL("")
    End If
    '
    Call GENEOPERFAP
    '
    Screen.MousePointer = 11
    For I& = 1 To CAEQUI&
      XX$ = REGLEIDO$("PADMAQ", I&)
      EQQ$(I&) = Left$(XX$, 6)
    Next I&
    '
    ULTRE& = ULTREG&("OPERFAP")
    For TTL& = 1 To ULTRE&
      SCP$ = REGLEIDO$("OPERFAP", TTL&)
      NORFA$ = TRIM$(Mid$(SCP$, 4, 9))
      PPXX% = InStr(NORFA$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NORFA$)
      NOFAX = XVALO(Left$(NORFA$, PPXX%))
      If NOFAX > 0 Then
        If NOFAX <= 999999 Then
          EQ1$ = Mid$(SCP$, 73, 6)
          EQ2$ = Mid$(SCP$, 203, 6)
          EQ3$ = Mid$(SCP$, 209, 6)
          EQ4$ = Mid$(SCP$, 215, 6)
          '
          MINUPREP = CVS(Mid$(SCP$, 103, 4))
          MINLOTE = 0: For UU% = 107 To 130 Step 4
                         MINLOTE = MINLOTE + CVS(Mid$(SCP$, UU%, 4))
                       Next UU%
          CAOPS = 1    '  CVS(MID$(SCP$,83,4))
          '
          For j% = 1 To CAEQUI&
            If TRIM$(EQQ$(j%)) <> "" Then
              If EQQ$(j%) = EQ1$ Then
                 TPREP(j%) = TPREP(j%) + MINUPREP
                 TOPER(j%) = TOPER(j%) + MINLOTE
                ElseIf EQQ$(j%) = EQ2$ Then
                 TPREP(j%) = TPREP(j%) + MINUPREP
                 TOPER(j%) = TOPER(j%) + MINLOTE
                ElseIf EQQ$(j%) = EQ3$ Then
                 TPREP(j%) = TPREP(j%) + MINUPREP
                 TOPER(j%) = TOPER(j%) + MINLOTE
                ElseIf EQQ$(j%) = EQ4$ Then
                 TPREP(j%) = TPREP(j%) + MINUPREP
                 TOPER(j%) = TOPER(j%) + MINLOTE
              End If
            End If
          Next j%
        End If
      End If
    Next TTL&
    '
    For I& = 1 To CAEQUI&
      XX$ = REGLEIDO$("PADMAQ", I&)
      Call REPLA(XX$, MKS$(60 * TPREP(I&)), 81, 4)
      Call REPLA(XX$, MKS$(60 * TOPER(I&)), 85, 4)
      Call REPLA(XX$, MKS$(60 * (TPREP(I&) + TOPER(I&))), 89, 4)
      Call GRAREG("PADMAQ", XX$, I&)
    Next I&
    '
    Call CIERRARCH("PADMAQ")
    Screen.MousePointer = 1
    '
    Call FRESHALL
    Call FINAL("*TOTIEPRO")
    '
End Sub

Function ValidaArmador%(NORFA$)
       '
       'Valida que la orden de fabricacion tenga armador externo asociado y
       'que el mismo sea Externo
       '0 = No es Armador Externo
       '1 = Armador Externo
       Vali% = 0
       CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
       ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
       If ARMAD% > 0 Then
            TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
            TIPOARM$ = TRIM$(UCase$(Mid$(TX$, 67, 1)))
            If TIPOARM$ = "E" Then Vali% = 1
       End If
       '
       ValidaArmador% = Vali%
End Function

Sub VERIDEST()
    '
    STO1% = 0: STO2% = 0
    For I& = 1 To ULTREG&("DESTFAB")
       X$ = REGLEIDO$("DESTFAB", I&)
       CODEST$ = TRIM$(UCase$(Left$(X$, 8)))
       If CODEST$ = "STOCK" Then STO1% = 1
       If CODEST$ = "SEMIELA" Then STO2% = 1
    Next I&
    '
    If STO1% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "STOCK", 1, 8)
       Call REPLA(X$, "Stock de Producto Terminado", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    If STO2% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "SEMIELA", 1, 8)
       Call REPLA(X$, "Stock de Semielaborados", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    Call CIERRARCH("DESTFAB")
    '
End Sub

Sub REVIORFLAN()
   '
   Screen.MousePointer = 11
   Dim NUORF(16384), REORF&(16384), STORF%(16384)
   j& = 0
   FIN& = ULTREG&("ORFANUE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORFANUE", U&)
      If Asc(Mid$(XX$, 112, 1)) = 1 Then
         j& = j& + 1
         NUORF(j&) = CVS(Left$(XX$, 4))
         REORF&(j&) = U&
      End If
   Next U&
   '
   FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         NOFA = CVS(Left$(XX$, 4))
         For K% = 1 To j&
            If NUORF(K%) = NOFA Then
               STORF%(K%) = 1
            End If
         Next K%
      End If
   Next U&
   '
   For U& = 1 To j&
      Call TRACE(20, U&, j&)
      If STORF%(U&) <> 1 Then
         XX$ = REGLEIDO$("ORFANUE", REORF&(U&))
         Call REPLA(XX$, Chr$(3), 112, 1)
         Call GRAREG("ORFANUE", XX$, REORF&(U&))
      End If
   Next U&
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub GENEOPERFAP()
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
    Exit Sub
    
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
    'Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
        CIXII% = XVALO(!Cod_Inte)
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
            Call REPLA(YY$, SAFETEXT$(!IDSUBOR), 1, 12)
            Call REPLA(YY$, MKI$(XVALO(!Cod_Inte)), 13, 2)
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
            Call REPLA(YY$, SAFETEXT$(!OBSERVA), 151, 48)
            On Error GoTo 0
            '
            JJ& = JJ& + 1
            Call GRAREG("OPERFAP", YY$, JJ&)
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
    Call SETULTREG("OPERFAP", JJ&)
    Call SETULTREG("OPERFAH", JJ&)
    Call FILESORT("OPERFAP", "", 1, 4, "ASC")
    Call CIERRARCH("OPERFAP")
    Call FILESORT("OPERFAH", "", 1, 4, "ASC")
    Call CIERRARCH("OPERFAH")
    '
99  Screen.MousePointer = SCMA
    '
End Sub

Sub GENEOPERORF(NORFAX$)
    '
    SCMA = Screen.MousePointer
    Call ABREORFAB
    '
    Screen.MousePointer = 11
    SQLX$ = " SELECT * FROM OPERFAB "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFAX$ & "' "
    SQLX$ = SQLX$ + " AND (StatOper = 0 OR StatOper = 1) "
    SQLX$ = SQLX$ + " AND CantReal < CantProg "
    SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
    'Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)

    Dim OPERFAPTMP As ADODB.Recordset
    Set OPERFAPTMP = New ADODB.Recordset
    OPERFAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With OPERFAPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        ' Call MENSERR(24, "No Existen Operaciones Pendientes.")
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
        CIXII% = XVALO(!Cod_Inte)
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
            YY$ = REGBLAN("OPERFAB")
            Call REPLA(YY$, SAFETEXT$(!IDSUBOR), 1, 12)
            Call REPLA(YY$, MKI$(!Cod_Inte), 13, 2)
            Call REPLA(YY$, MKI$(!NUMEOPER), 15, 2)
            Call REPLA(YY$, SAFETEXT$(!DESCOPER), 17, 48)
            'Call REPLA(YY$, !Destino, 65, 8)
            Call REPLA(YY$, SAFETEXT$(!MAQUINA), 73, 6)
            Call REPLA(YY$, MKD$(!CANTPROG), 79, 8)
            Call REPLA(YY$, MKD$(!CANTREAL), 87, 8)
            Call REPLA(YY$, MKD$(!CANTPEND), 95, 8)
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
            Call REPLA(YY$, SAFETEXT$(!OBSERVA), 151, 48)
            '
            JJ& = JJ& + 1
            Call GRAREG("!OPERFAB", YY$, JJ&)
            '
          End If
        End If
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!OPERFAB", JJ&)
    Call CIERRARCH("!OPERFAB")
    '
99  Screen.MousePointer = SCMA
    '
End Sub
'
Sub GENORDEFDAT(MODO%, Optional SECSEL)
  '
  SECSEL1 = XVALO(SECSEL)
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM ORDEFABR "
  If MODO% = 1 Then SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
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
    ASELE$ = "ORDEFABR":    ASEL2$ = "ORDEFABP":
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

      Y$ = REGBLAN("ORDEFABR")
      '
      On Error Resume Next
      Call REPLA(Y$, MKS$(XVALO(!NUMEORFA)), 1, 4)
      Call REPLA(Y$, SAFETEXT$(!REFEREN), 5, 64)
      Call REPLA(Y$, SAFETEXT$(!IDSUBOR), 69, 12)
      Call REPLA(Y$, MKI$(CVINT(!FechGen)), 81, 2)
      Call REPLA(Y$, MKI$(XVALO(!Cod_Inte)), 83, 2)
      Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
      Call REPLA(Y$, MKD$(XVALO(!CanProy)), 89, 8)
      '
      Call REPLA(Y$, SAFETEXT$(!DESTINO), 97, 8)
      '
      Call REPLA(Y$, MKI$(CVINT(!FechEnSol)), 105, 2)
      Call REPLA(Y$, Chr$(XVALO(!Priorid)), 107, 1)
      Call REPLA(Y$, Chr$(XVALO(!Statuorf)), 108, 1)
      Call REPLA(Y$, MKI$(CVINT(!FechLan)), 109, 2)
      Call REPLA(Y$, MKS$(XVALO(!SecciPro)), 113, 4)
      Call REPLA(Y$, MKI$(CVINT(!FechCierr)), 117, 2)
      Call REPLA(Y$, MKD$(XVALO(!CanApro)), 119, 8)
      Call REPLA(Y$, MKD$(XVALO(!CanRech)), 127, 8)
      Call REPLA(Y$, SAFETEXT$(!IdenRetra), 135, 8)
      Call REPLA(Y$, MKI$(CVINT(!FechProyEntre)), 147, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechTempEntre)), 149, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechRecal)), 151, 2)
      Call REPLA(Y$, SAFETEXT$(!OBSERVA), 155, 64)
      On Error GoTo 0
      '
      JJ& = JJ& + 1
      Call GRAREG(ASELE$, Y$, JJ&)
      Call TRACE(20, JJ&, FIN&)
      '
      If XVALO(!Statuorf) < 2 Then
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
End Sub

Sub CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, MODOXX%)
    '
    ' VACONTROL% = 1  presenta componentes para edicion
    '              0  no presenta componentes
    '             -1  devuelto como señal de cancelacion
    '
 ARCHIVOX1$ = "MATEROF"
' EPAC$ = TRIM$(UCase$(EMPREAC$))
' If InStr(EPAC$, "MEDITEA") > 0 Then ARCHIVOX1$ = "MATEROF9"
    
    Screen.MousePointer = 11
    Call ABRESTOCKS
    Call BLANARCH("!" & ARCHIVOX1$)
    CAIT% = 0: COEF = 1
    Dim RERES&(1024)
    '
    VERISARRAS% = 0
    If Control$("", "CHECOMPO") = "SI" Then
      VERISARRAS% = 1
    End If
    '
    CIXI% = CODINT%(CODICONJ$)
    If CIXI% < 1 Then
       Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Código de Artículo '" + TRIM$(CODICONJ$) + "' Inexistente\Consulte a su Soporte de Sistemas.")
       'Call FINAL("")' se cambio esta linea por el exit sub para que  luego que muestra el mensaje continue con el proceso
       Exit Sub '
    End If
    CODPRO$ = CODICONJ$
    '
50  If TRIM$(NORFABRI$) <> "" Then
         TITUPA$ = "Consumo según Reservas"
         GoTo 100                ' lee de reservas
       ElseIf CODINT%(CODICONJ$) > 0 Then
         GoTo 200                ' lee de formula standard
       Else
         Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Consulte a su Soporte de Sistemas.")
         Call FINAL("")
    End If
    '
100 NORFA$ = AJUSTI$(NORFABRI$, 12)
    NFX1$ = Mid$(NORFA$, 4)
    PPXX% = InStr(NFX1$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NFX1$)
    NUORFA& = XVALO(Left$(NFX1$, PPXX% - 1))
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then GoTo 200   ' No Existen Reservas efectuadas.
      On Error GoTo 0
      CAIT% = 0: JJ& = 0
      Do While (.EOF = False)
        If !Cod_Inte <> CIJ%(CAIT%) Or !UsoUnit <> USOI(CAIT%) Then
          CAIT% = CAIT% + 1
          CIJ%(CAIT%) = XVALO(!Cod_Inte)
          USOI(CAIT%) = XVALO(!UsoUnit)
        End If
      .MoveNext
      Loop
    End With
    Set RESERTMP = Nothing
    If CAIT% > 0 Or Left$(NORFABRI$, 2) = "OR" Then GoTo 300
    '
    '
200 ' arma composicion segun formula standard
    TITUPA$ = "Consumo por Fórmula Standard"
    Call LEEEXPLOMUL(CODPRO$, 1, 1)
    If CAIT% < 1 Then             ' producto sin estructura
      Call MENSERR(24, "No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
      TITUPA$ = "Consumo Informado"
    End If

300 If VACONTROL% = 1 Then
      If VERISARRAS% = 1 Then
        For KU& = 1 To NUMAS%
          CIMOVI%(KU&) = 0
        Next KU&
        '
        For I1% = 1 To CAIT%
          CIMOVI%(CIJ%(I1%)) = 1
        Next I1%
        '
        'Call VESARRGLO
      End If
    End If
    '
    '
    ' GENERA UN VECTOR STACDIS(32767) y STACRES(32767)
'    Dim STACDIS(32767), STACRES(32767), STOCKPORDEPO(32767)
    Dim STACDIS(32767), STACRES(32767), STOCKTRANSFERIDO(32767), CONSUMOX(32767)
    '
    'SI EL MODO ES 2 O 3 NO CALCULA LA SITUACION DE STOCK EN ESTE MOMENTO.
    'LO REALIZA DE OTRA FORMA AL RECORRER EL MATEROF (VER MAS ABAJO)
    If MODOXX% = 2 Or MODOXX% = 3 Or MODOXX% = 5 Then GoTo 400
    '
    Call CARDEPOS
    CONDIX$ = ""
    'AQUI  SELECCIONAR ARTICULOS DE LA ESTRUCTURA PARA SOLO RECORRER LOS NECESARIOS
    For I1% = 1 To CAIT%
       UI1& = I1%
       If CIJ%(I1%) > 0 Then
              If I1% = 1 Then
                INGRECONDI% = 1
                CONDIARTI$ = " (Cod_Inte = " & CIJ%(I1%)
              Else
                CONDIARTI$ = CONDIARTI$ + " or Cod_Inte = " & CIJ%(I1%)
              End If
       End If
       
    Next I1%
    If INGRECONDI% = 1 Then
           CONDIARTI$ = CONDIARTI$ + ")"
    End If
    '
    'SI ES SOLO PAÑOL FILTRA POR EL DEPOSITO PAÑOL
    SOLOPAÑOL% = 0
    If Control("", "STOPAÑOL") = "SI" Then
        DP% = XVALO(Control("", "DEPOCEN"))
        CONDIX$ = " DepoMovi = " & DP%
        SOLOPAÑOL% = 1
    Else ' CASO CONTRARIO TOMA TODOS LOS DEPOSITOS DISPONIBLES
        For UI% = 1 To NUSU%
           If DISTO%(UI%) >= 0 Then
             CONDIX$ = CONDIX$ + " OR DepoMovi = " & COSU%(UI%)
           End If
        Next UI%
        CONDIX$ = "(" & Mid$(CONDIX$, 4) & ")"
    End If
    '
    
    
    strsql$ = "Select COD_INTE, sum(CantIngre) AS SUINGRE, SUM(CantSalid) as SUSALI "
    strsql$ = strsql$ + " from movisto WITH(NOLOCK) where " & CONDIX$
    If INGRECONDI% > 0 Then strsql$ = strsql$ + " And " & CONDIARTI$
    strsql$ = strsql$ + " AND COD_INTE >0"
    strsql$ = strsql$ + " GROUP BY COD_INTE"
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         STACDIS(CIXI%) = XVALO(!SUINGRE) - XVALO(!SUSALI)
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
35  Call ABRESERVA
    SQLX$ = " SELECT COD_INTE, SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS "
    SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CantCons < (CantRes - 0.005) AND  COD_INTE >0"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         STACRES(CIXI%) = XVALO(!SURESE) - XVALO(!SUCONS)
         CONSUMOX(CIXI%) = XVALO(!SUCONS)
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    ' AGREGA MATERIAL PENDIENTE DE CONSIGNAR
    Call ABRECOMPRAS
    SQLX$ = " SELECT COD_INTE, SUM(CANT_ASIG) AS SUASI, SUM(CANT_ENTREG) AS SUENT "
    SQLX$ = SQLX$ + " FROM MACONSIG WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Cant_ENTREG < (Cant_ASIG - 0.005) AND COD_INTE > 0 "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE "
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         STACRES(CIXI%) = STACRES(CIXI%) + (XVALO(!SUASI) - XVALO(!SUENT))
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    '
    'SE CARGA EN UN VECTOR EL STOCK POR DEPOSITO POR QUE ES MUCHO MAS RAPIDO QUE LA FUNCION STODEPOS
    'EN PRUEBA EN MEDITEA UNA CARGA DE COMPONETNES POR ORDEN DE FABRICACION QUE DEMORABA 145 SEGUNDOS BAJO A 20
''''''    SOLOPAÑOL% = 0
''''''    If CONTROL("", "STOPAÑOL") = "SI" Or MODOXX% = 4 Then
''''''        DP% = XVALO(CONTROL("", "DEPOCEN"))
''''''        SOLOPAÑOL% = 1
''''''        STRSQL$ = "Select COD_INTE, sum(CantIngre) AS SUINGRE, SUM(CantSalid) as SUSALI "
''''''        STRSQL$ = STRSQL$ + " from movisto WITH(NOLOCK) where DEPOMOVI  = " & DP%
''''''        STRSQL$ = STRSQL$ + "  GROUP BY COD_INTE "
''''''        Set MIRS = READSET(STRSQL$)
''''''        With MIRS
''''''           Do While Not .EOF
''''''             CIXI% = XVALO(!COD_INTE)
''''''             If CIXI% > 0 Then STOCKPORDEPO(CIXI%) = XVALO(!SUINGRE) - XVALO(!SUSALI)
''''''           .MoveNext
''''''           Loop
''''''        End With
''''''        MIRS.Close
''''''        Set MIRS = Nothing
''''''        '
        ''''''''''''''''''''''''''''''''''
''''''        For I1% = 1 To CAIT%
''''''           UI1& = I1%
''''''           If CIJ%(I1%) > 0 Then
''''''                  If I1% = 1 Then
''''''                    INGRECONDI% = 1
''''''                    CONDI$ = " (Cod_Inte = " & CIJ%(I1%)
''''''                  Else
''''''                    CONDI$ = CONDI$ + " or Cod_Inte = " & CIJ%(I1%)
''''''                  End If
''''''           End If
''''''
''''''        Next I1%
''''''        If INGRECONDI% = 1 Then
''''''               CONDI$ = "WHERE " & CONDI$ + ")"
''''''        End If
    If SOLOPAÑOL% > 0 Then
        SQLX$ = "SELECT IDSUBOR, COD_INTE FROM RESERVA WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE COD_INTE > 0 "
        If INGRECONDI% = 1 Then SQLX$ = SQLX$ + " AND " & CONDIARTI$
        SQLX$ = SQLX$ + "GROUP BY COD_INTE, IDSUBOR "
        Set rst = READSET(SQLX$)
        With rst
          Do While Not .EOF
               NORFAX1$ = SAFETEXT$(!IDSUBOR)
               CIXI% = XVALO(!Cod_Inte)
               SQLX$ = " Select COD_INTE, sum(CantIngre) AS SUINGRE , SUM(CantSalid) as SUSALI"
               SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) "
               SQLX$ = SQLX$ + " Where DEPOMOVI = " & DP%
               SQLX$ = SQLX$ + " AND CodiMovi = 'TR'"
               SQLX$ = SQLX$ + " AND DOSECCOR = '" & TRIM$(NORFAX1$) & "'"
               SQLX$ = SQLX$ + " AND  COD_INTE > 0 "
               SQLX$ = SQLX$ + "  GROUP BY COD_INTE "
               Set MIRS = READSET(SQLX$)
               With MIRS
                  Do While Not .EOF
                     CIXI% = XVALO(!Cod_Inte)
                     STOCKTRANSFERIDO(CIXI%) = XVALO(!SUINGRE) - XVALO(!SUSALI)
                    .MoveNext
                  Loop
               End With
               MIRS.Close
               Set MIRS = Nothing
        
            .MoveNext
          Loop
        End With
        Set RESERTMP = Nothing
    ''''''''''''''''''''''''''''''''''
    End If

400 FIN& = CAIT%
    DP% = XVALO(Control("", "DEPOCEN"))

    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
         If Abs(USOI(I1%)) > 0 Then
           '
           ZZ$ = REGBLAN$(ARCHIVOX1$)
           Call REPLA(ZZ$, MKI$(CIJ%(I1%)), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CIJ%(I1%)), 3, 2)
             NEBRU = CACONJU# * USOI(I1%) * COEF
           Call REPLA(ZZ$, MKS$(NEBRU), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(USOI(I1%) * COEF)), 9, 8)
           Call REPLA(ZZ$, MKS$(RERES&(I1%)), 17, 4)
             If MODOXX% = 2 Then           ' CASO DE CIERRE DE O/F
                  STDIS = STOCKACT(CIJ%(I1%))
                ElseIf MODOXX% = 3 Then    ' CASO DE RE-GENERAR RESERVAS
                  STDIS = STOCKDIS(CIJ%(I1%)) + STORESOF(CIJ%(I1%), NUORFA&)
                ElseIf MODOXX% = 4 Then    ' CASO DE RE-GENERAR RESERVAS
'                  STDIS = STODEPOS(CIJ%(I1%), DP%)
                   STDIS = STACDIS(CIJ%(I1%))
                ElseIf MODOXX% = 5 Then
                  STDIS = 0
                Else
                  ' CASO DE GENERACION DE UNA O/F
                  If SOLOPAÑOL% > 0 Then 'CONSIDERA SOLO EL DEPOSITO DE PAÑOL, DESCONTANDO LAS RESERVAS
                     'EN ESTE CASO DESCUENTA LA CANTIDAD QUE YA FUE TRANSFERIDA DEL DEPOSITO 50
                     'DE ESTE MODO SE DESCUENTA DE LA CANTIDAD RESERVADA LA MERCADERIA QUE YA SE TRANSFIRIO
                     'SE SUMA AL STOCK RESERVADO EL CONSUMIDO PARA QUE MANTENGA LA CANTIDAD ORIGINAL DEL RESERVADO P'Q' SI NO NO DABA.
                     STDIS = STACDIS(CIJ%(I1%)) - ((STACRES(CIJ%(I1%)) + CONSUMOX(CIJ%(I1%))) - STOCKTRANSFERIDO(CIJ%(I1%)))
                  Else ' CONSIDERA TODOS LOS DEPOSITOS , DESCONTANDO EL STOCK RESERVADO
                     STDIS = STACDIS(CIJ%(I1%)) - STACRES(CIJ%(I1%))   ' AQUI FALTA DESCONTAR EL PENDIENTE DE CONSIGNACION
                  End If
                  '
             End If
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
             FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
             If MODOXX% = 5 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!" & ARCHIVOX1$, ZZ$)
           '
         End If
       End If
    Next I1%
    Call CIERRARCH("!" & ARCHIVOX1$)
    Screen.MousePointer = 1
    '
    If VACONTROL% = 1 Then
       CIXI% = CODINT%(CODICONJ$) ' LO VUELVO A LEER POR QUE SE ESTA QUEDANDO CON OTRA ASINGACION POSTERIOR
       TITUPA$ = TITUPA$ + ": "
       If TRIM$(NORFABRI$) <> "" Then TITUPA$ = TITUPA$ + TRIM$(NORFABRI$)
       If TRIM$(CODPRO$) <> "" Then TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
       '
       VTB% = VENTABU%(ARCHIVOX1$ & "/TITUPAN=" + TITUPA$)
'       VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
       '
       If VTB% < 0 Then
            VACONTROL% = (-1)
          Else
            If Abs(CACONJU#) >= 0.05 Then
               For I& = 1 To ULTREG&("!" & ARCHIVOX1$)
                  ZZ$ = REGLEIDO$("!" & ARCHIVOX1$, I&)
                  CANUE# = CDbl(CVS(Mid$(ZZ$, 5, 4)))
                  USNUE# = CANUE# / CACONJU#
                  Call REPLA(ZZ$, MKD$(USNUE#), 9, 8)
                  Call GRAREG("!" & ARCHIVOX1$, ZZ$, I&)
               Next I&
            End If
       End If
    End If
    '
End Sub
Sub CARCOMPO_ORIGINAL(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, MODOXX%)
    '
    ' VACONTROL% = 1  presenta componentes para edicion
    '              0  no presenta componentes
    '             -1  devuelto como señal de cancelacion
    '
 ARCHIVOX1$ = "MATEROF"
' EPAC$ = TRIM$(UCase$(EMPREAC$))
' If InStr(EPAC$, "MEDITEA") > 0 Then ARCHIVOX1$ = "MATEROF9"
    
    Screen.MousePointer = 11
    Call ABRESTOCKS
    Call BLANARCH("!" & ARCHIVOX1$)
    CAIT% = 0: COEF = 1
    Dim RERES&(1024)
    '
    VERISARRAS% = 0
    If Control$("", "CHECOMPO") = "SI" Then
      VERISARRAS% = 1
    End If
    '
    CIXI% = CODINT%(CODICONJ$)
    If CIXI% < 1 Then
       Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Código de Artículo '" + TRIM$(CODICONJ$) + "' Inexistente\Consulte a su Soporte de Sistemas.")
       'Call FINAL("")' se cambio esta linea por el exit sub para que  luego que muestra el mensaje continue con el proceso
       Exit Sub '
    End If
    CODPRO$ = CODICONJ$
    '
50  If TRIM$(NORFABRI$) <> "" Then
         TITUPA$ = "Consumo según Reservas"
         GoTo 100                ' lee de reservas
       ElseIf CODINT%(CODICONJ$) > 0 Then
         GoTo 200                ' lee de formula standard
       Else
         Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Consulte a su Soporte de Sistemas.")
         Call FINAL("")
    End If
    '
100 NORFA$ = AJUSTI$(NORFABRI$, 12)
    NFX1$ = Mid$(NORFA$, 4)
    PPXX% = InStr(NFX1$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NFX1$)
    NUORFA& = XVALO(Left$(NFX1$, PPXX% - 1))
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
    '
    'Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    RESERTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then GoTo 200   ' No Existen Reservas efectuadas.
      On Error GoTo 0
      CAIT% = 0: JJ& = 0
      Do While (.EOF = False)
        If !Cod_Inte <> CIJ%(CAIT%) Or !UsoUnit <> USOI(CAIT%) Then
          CAIT% = CAIT% + 1
          CIJ%(CAIT%) = XVALO(!Cod_Inte)
          USOI(CAIT%) = XVALO(!UsoUnit)
        End If
      .MoveNext
      Loop
    End With
    Set RESERTMP = Nothing
    If CAIT% > 0 Or Left$(NORFABRI$, 2) = "OR" Then GoTo 300
    '
    '
200 ' arma composicion segun formula standard
    TITUPA$ = "Consumo por Fórmula Standard"
    Call LEEEXPLOMUL(CODPRO$, 1, 1)
    If CAIT% < 1 Then             ' producto sin estructura
      Call MENSERR(24, "No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").")
      TITUPA$ = "Consumo Informado"
    End If

300 If VACONTROL% = 1 Then
      If VERISARRAS% = 1 Then
        For KU& = 1 To NUMAS%
          CIMOVI%(KU&) = 0
        Next KU&
        '
        For I1% = 1 To CAIT%
          CIMOVI%(CIJ%(I1%)) = 1
        Next I1%
        '
        'Call VESARRGLO
      End If
    End If
    '
    '
    ' GENERA UN VECTOR STACDIS(32767) y STACRES(32767)
    Dim STACDIS(32767), STACRES(32767)
    '
    'SI EL MODO ES 2 O 3 NO CALCULA LA SITUACION DE STOCK EN ESTE MOMENTO.
    'LO REALIZA DE OTRA FORMA AL RECORRER EL MATEROF (VER MAS ABAJO)
    If MODOXX% = 2 Or MODOXX% = 3 And MODOXX% = 5 Then GoTo 400
    '
    Call CARDEPOS
    CONDIX$ = ""
    For UI% = 1 To NUSU%
       If DISTO%(UI%) >= 0 Then
         CONDIX$ = CONDIX$ + " OR DepoMovi = " & COSU%(UI%)
       End If
    Next UI%
    CONDIX$ = "(" & Mid$(CONDIX$, 4) & ")"
    strsql$ = "Select COD_INTE, sum(CantIngre) AS SUINGRE, SUM(CantSalid) as SUSALI "
    strsql$ = strsql$ + " from movisto WITH(NOLOCK) where " & CONDIX$ & " GROUP BY COD_INTE"
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         If CIXI% > 0 Then STACDIS(CIXI%) = XVALO(!SUINGRE) - XVALO(!SUSALI)
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
35  Call ABRESERVA
    SQLX$ = " SELECT COD_INTE, SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS "
    SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CantCons < (CantRes - 0.005) GROUP BY COD_INTE"
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         STACRES(CIXI%) = XVALO(!SURESE) - XVALO(!SUCONS)
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    ' AGREGA MATERIAL PENDIENTE DE CONSIGNAR
    Call ABRECOMPRAS
    SQLX$ = " SELECT COD_INTE, SUM(CANT_ASIG) AS SUASI, SUM(CANT_ENTREG) AS SUENT "
    SQLX$ = SQLX$ + " FROM MACONSIG WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Cant_ENTREG < (Cant_ASIG - 0.005) GROUP BY COD_INTE"
    Set MIRS = New ADODB.Recordset
    MIRS.CursorLocation = adUseServer
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
       Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         STACRES(CIXI%) = STACRES(CIXI%) + (XVALO(!SUASI) - XVALO(!SUENT))
       .MoveNext
       Loop
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    '
400 FIN& = CAIT%
    DP% = XVALO(Control("", "DEPOCEN"))

    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
         If Abs(USOI(I1%)) > 0 Then
           '
           ZZ$ = REGBLAN$(ARCHIVOX1$)
           Call REPLA(ZZ$, MKI$(CIJ%(I1%)), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CIJ%(I1%)), 3, 2)
             NEBRU = CACONJU# * USOI(I1%) * COEF
           Call REPLA(ZZ$, MKS$(NEBRU), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(USOI(I1%) * COEF)), 9, 8)
           Call REPLA(ZZ$, MKS$(RERES&(I1%)), 17, 4)
             If MODOXX% = 2 Then           ' CASO DE CIERRE DE O/F
                  STDIS = STOCKACT(CIJ%(I1%))
                ElseIf MODOXX% = 3 Then    ' CASO DE RE-GENERAR RESERVAS
                  STDIS = STOCKDIS(CIJ%(I1%)) + STORESOF(CIJ%(I1%), NUORFA&)
                ElseIf MODOXX% = 4 Then    ' CASO DE RE-GENERAR RESERVAS
                  STDIS = STODEPOS(CIJ%(I1%), DP%)
                ElseIf MODOXX% = 5 Then
                  STDIS = 0
                Else
                  ' CASO DE GENERACION DE UNA O/F
                  If Control("", "STOPAÑOL") = "SI" Then 'CONSIDERA SOLO EL DEPOSITO DE PAÑOL, DESCONTANDO LAS RESERVAS
                     STDIS = STODEPOS(CIJ%(I1%), DP%) - STACRES(CIJ%(I1%))
                  Else ' CONSIDERA TODOS LOS DEPOSITOS , DESCONTANDO EL STOCK RESERVADO
                     STDIS = STACDIS(CIJ%(I1%)) - STACRES(CIJ%(I1%))   ' AQUI FALTA DESCONTAR EL PENDIENTE DE CONSIGNACION
                  End If
                  '
             End If
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
             FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
             If MODOXX% = 5 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!" & ARCHIVOX1$, ZZ$)
           '
         End If
       End If
    Next I1%
    Call CIERRARCH("!" & ARCHIVOX1$)
    Screen.MousePointer = 1
    '
    If VACONTROL% = 1 Then
       CIXI% = CODINT%(CODICONJ$) ' LO VUELVO A LEER POR QUE SE ESTA QUEDANDO CON OTRA ASINGACION POSTERIOR
       TITUPA$ = TITUPA$ + ": "
       If TRIM$(NORFABRI$) <> "" Then TITUPA$ = TITUPA$ + TRIM$(NORFABRI$)
       If TRIM$(CODPRO$) <> "" Then TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
       '
       VTB% = VENTABU%(ARCHIVOX1$ & "/TITUPAN=" + TITUPA$)
'       VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
       '
       If VTB% < 0 Then
            VACONTROL% = (-1)
          Else
            If Abs(CACONJU#) >= 0.05 Then
               For I& = 1 To ULTREG&("!" & ARCHIVOX1$)
                  ZZ$ = REGLEIDO$("!" & ARCHIVOX1$, I&)
                  CANUE# = CDbl(CVS(Mid$(ZZ$, 5, 4)))
                  USNUE# = CANUE# / CACONJU#
                  Call REPLA(ZZ$, MKD$(USNUE#), 9, 8)
                  Call GRAREG("!" & ARCHIVOX1$, ZZ$, I&)
               Next I&
            End If
       End If
    End If
    '
End Sub

'
Function CANREALI(NUMORFA$, NUMOPER%)
   '
   Call ABREORFAB
   '
   SQLX$ = " SELECT * FROM INFOFAB "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMORFA$ & "' "
   SQLX$ = SQLX$ + " AND NumeOper = " & NUMOPER% & " "
   '
   'Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim INFOFABTMP As ADODB.Recordset
   Set INFOFABTMP = New ADODB.Recordset
   INFOFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   CANREALI = 0
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If IsNull(!ReTraba) Or XVALO(!ReTraba) < 1 Then
           CANREALI = CANREALI + XVALO(!CANTREAL)
         End If
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
End Function

Function STATORFA%(NUMORFA$)
   '
   Static NORFAN$, STATAN%
   '
   If NUMORFA$ = NORFAN$ Then
      STATORFA% = STATAN%
      Exit Function
   End If
   '
   Call ABREORFAB
   SQLX$ = " SELECT * FROM ORDEFABR "
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMORFA$ & "'"
   '
   'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim ORDEFABTMP As ADODB.Recordset
   Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   STATORFA% = 99 ' no existe
   On Error Resume Next
   ORDEFABTMP.MoveFirst
   If Err < 1 Then
      STATORFA% = ORDEFABTMP!Statuorf
   End If
   ORDEFABTMP.Close
   Set ORDEFABTMP = Nothing
   '
   NORFAN$ = NUMORFA$
   STATAN% = STATORFA%
   '
End Function

Sub PROCERECHA(CODI1%, FERECH%, CARECHA, DOCORX$) '\\\OT-06-0191-WAR-09/06/06///
    '
    MODOFERVI% = 0
    If InStr(1, EMPREAC$, "FERVI") > 0 Then MODOFERVI% = 1
    '
    If CARECHA <= 0 Then
      Call MENSERR(24, "Cantidad de Rechazo no Válida.")
      Exit Sub
    End If
    '
    'CARGO EL DEFAULT DEL PASO 2
    VDEF$ = Space$(256)
    FIN1& = ULTREG&("RECHLIN")
    ULTRECH& = FIN1& + 1
    If FIN1& > 0 Then
       XX1$ = REGLEIDO$("RECHLIN", FIN1&)
       ULTRECH& = 1 + CVS(Mid$(XX1$, 1, 4))
    End If
    Call REPLA(VDEF$, MKS$(ULTRECH&), 199, 4)
    PROHACOD% = PROVHABI%(CODI1%)
    Call REPLA(VDEF$, MKI$(CODI1%), 1, 2)
    Call REPLA(VDEF$, MKI$(FERECH%), 3, 2)
    Call REPLA(VDEF$, MKI$(PROHACOD%), 5, 2)
    Call REPLA(VDEF$, Chr$(0), 12, 1)
    Call REPLA(VDEF$, Chr$(0), 11, 1)
    Call REPLA(VDEF$, Chr$(0), 16, 1)
    Call REPLA(VDEF$, DOCORX$, 17, 12) '\\\OT-06-0191-WAR-09/06/06///
    '
15  OBJ$ = OBJPLA$("DETARECHA", VDEF$)
    PA2$ = OBJ$
    If TRIM$(PA2$) <> "" Then
      'VALIDACIONES
      CODI2% = CVI(Mid$(PA2$, 1, 2))
      FEMI1$ = FECHATEX$(CVI(Mid$(PA2$, 3, 2)))
      NPRO% = CVI(Mid$(PA2$, 5, 2))
      DESTI% = Asc(Mid$(PA2$, 11, 1))
      SECT& = Asc(Mid$(PA2$, 12, 1)) 'AHORA DEPOSITO
      MOTI% = Asc(Mid$(PA2$, 16, 1))
      CAN3 = CARECHA
      REFE1$ = TRIM$(Mid$(PA2$, 17, 48))
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      If TRIM$(ECOARCH$("PROVED1", MKI$(NPRO%))) = "" Then
        If CARPROVE% > 0 Then
          Call MENSERR(24, "Proveedor no Válido.")
          VDEF$ = PA2$
          GoTo 15
        End If
      End If
      If TRIM$(ECOARCH$("TADESTI", Chr(DESTI%))) = "" Then
        Call MENSERR(24, "Destino de Rechazo no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(ECOARCH$("TADEPOSI", Chr$(SECT&))) = "" Then
        Call MENSERR(24, "Depósito no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%))) = "" Then
        Call MENSERR(24, "Motivo de Rechazo no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(REFE1$) = "" Then
        Call MENSERR(24, "Referencia de Rechazo no Válida.")
        VDEF$ = PA2$
        GoTo 15
      End If
      '
      ZZZ$ = FILTERGETRECORD$("TADESTI", 1, Chr$(DESTI%))
      CARPROVE% = 0
      If TRIM$(UCase$(Mid$(ZZZ$, 53, 2))) = "SI" Then
        CARPROVE% = 1
      End If
      '
      'CALCULO DE PRECIO E IMPORTE
      PREUNI1# = ULTIPRE#(NPRO%, CODI2%)
      IMPORT1# = PREUNI1# * CAN3
      If CARPROVE% < 1 Then
         PREUNI1# = 0: IMPORT1# = 0
         If MODOFERVI% = 0 Then NPRO% = 0
      End If
      'REGISTRO DE RECHAZO EN BDD
      FIN1& = ULTREG&("RECHLIN")
      ULTRECH& = FIN1& + 1
      If FIN1& > 0 Then
         XX1$ = REGLEIDO$("RECHLIN", FIN1&)
         ULTRECH& = 1 + CVS(Mid$(XX1$, 1, 4))
      End If
      Y$ = Space(128)
      Call REPLA(Y$, MKS$(ULTRECH&), 1, 4)
      Call REPLA(Y$, MKI$(CODI2%), 5, 2)
      Call REPLA(Y$, MKI$(FECHANUM(FEMI1$)), 7, 2)
      Call REPLA(Y$, MKI$(NPRO%), 9, 2)
      Call REPLA(Y$, MKS$(CAN3), 11, 4)
      Call REPLA(Y$, Chr$(DESTI%), 15, 1)
      Call REPLA(Y$, Chr$(MOTI%), 16, 1)
      Call REPLA(Y$, USERNAME$, 17, 32)
      Call REPLA(Y$, REFE1$, 49, 48) 'FALTA REFERENCIA
      Call REPLA(Y$, MKD$(PREUNI1#), 97, 8) 'FALTA VALOR UNITARIO
      Call REPLA(Y$, MKD$(IMPORT1#), 105, 8) 'FALTA IMPORTE
      Call REPLA(Y$, MKI$(0), 113, 2) 'FALTA PROCESADO FECHA
      Call REPLA(Y$, MKI$(0), 115, 2) 'FALTA NUMERO USUARIO PROCESO
      Call REGAPP("RECHLIN", Y$)
      Call CIERRARCH("RECHLIN")
      'AGREGO A ARCHIVO INDICE DE RECHAZO
      REFERE2$ = FEMI1$
         If CARPROVE% > 0 Then
              REFERE2$ = REFERE2$ + " - " + RASOCLI$(NPRO% * (-1))
            Else
              REFERE2$ = REFERE2$ + " - Cargo Interno"
         End If
      ZZ$ = Space(64)
      Call REPLA(ZZ$, MKS$(ULTRECH&), 1, 4)
      Call REPLA(ZZ$, REFERE2$, 5, 60)
      Call REGAPP("INDIRECH", ZZ$)
      Call CIERRARCH("INDIRECH")
      'ACA HAGO LOS MOVIMIENTOS DE STOCK SEGUN DEPOSITO
      DepOrig% = Val(SECT&) 'Asc(Mid$(XXX$, 81, 1)) 'DEP. ORIGEN
      ZZZ$ = FILTERGETRECORD$("TADESTI", 1, Chr$(DESTI%))
      DepoDes% = Asc(Mid$(ZZZ$, 32, 1)) 'DEP. DESTINO
      '
      Screen.MousePointer = 11
      '
      Call ABRESTOCKS
      SQLX$ = " SELECT max (CAST(SUBSTRING(DoPriCor,4,6) AS INTEGER)) as Maximo FROM MOVISTO"
      SQLX$ = SQLX$ + " WHERE CodiMovi = 'TR'"
      
      UNUM& = 0
    
      Dim MIRS As ADODB.Recordset
      Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If Not IsNull(MIRS!maximo) Then
        UNUM& = XVALO(MIRS!maximo)
      End If
    
      NRO& = 1 + UNUM&
      MIRS.Close
      Set MIRS = Nothing
      '\\\OT-06-0650-OD-FIN///
110   Screen.MousePointer = 1
      REMI$ = "TR." + TRIM$(Str$(NRO&))
      '
      FEMOVI% = FECHANUM(FEMI1$)
      CI% = CODI2%
      LOTE$ = ""
      DOSEC$ = "RECH-" + TRIM$(Str$(NRO&))
      REFMO$ = REFE1$
      CANTI = (-1) * CAN3
      '
      '\\\OT-06-0650-OD-28/08/08///
      Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, 0, PREUNI1#, 1, (-1) * NPRO%, DepOrig%, CANTI)
      Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", REMI$, REMI$, DOSEC$, DOSEC$, REFMO$, 0, PREUNI1#, 1, (-1) * NPRO%, DepoDes%, (-1) * CANTI)
     '
      'PREPARACION PARA IMPRESIÓN
      CI% = CODI2%
      NUDO$ = FORMATNUM$(ULTRECH&, "F11")
      Call PRIRECHALI(1, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION LOCAL
      '
      If MODOFERVI% = 1 Then GoTo 25 ' NO IMPRIME EN TESORERIA Y MAIL PROVEEDOR
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      Call PRIRECHALI(2, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION REMOTA
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      EMAI$ = TRIM$(EMAILCLI$((-1) * NPRO%))
      If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
          Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Dirección de Correo del Proveedor es\Inexistente o no Válida.")
        Else
           '\\\OT-08-0802-OD-06/09/08///
           'OPCION DE CONSULTAR PRIMERO SI DESEA ENVIAR POR MAIL EL RECHAZO DE LINEA
           OPX% = COMALTER%("¿Desea Enviar Rechazo de Linea Por Email?\\No\Si")
           If OPX% = 2 Then
                SENDBYMAIL$ = EMAI$
                Call PRIRECHALI(1, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#)
           End If
           '\\\OT-08-0802-OD-FIN///
      End If
      '
25  End If
    '
End Sub
Sub PRIRECHALI(DONDE%, FETE$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFE1$, PREUNI1#, IMPORT1#)
    ' IMPRESION DEL RECHAZO EN LÍNEA
    'DONDE%=1 IMPRESORA LOCAL (IMPRESORA DEFAULT)
    'DONDE%=2 IMPRESORA REMOTA (IMPRESORA CUENTAS A PAGAR)
    '
    FORIPRE$ = Control$("", "RECHALIN")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Printer.TrackDefault = False
      If DONDE% = 2 Then 'IMPRESION REMOTA
        PORTUFAC$ = TRIM$(Control$("", "RUIMPREC"))
        PRINTERPORT$ = PORTUFAC$
      End If
      '
      Screen.MousePointer = 11
      'FECHEMII% = FECHANUM(Label2)
      FECHEMII% = FECHANUM(FETE$)
      FECHDOC$ = FECHATEX$(FECHEMII%)
      'NUMEDOC$ = TRIM$(Label10)
      NUMEDOC$ = TRIM$(NUDO$)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Rechazo en Línea"
      '
      CODPARAM$(1) = "COD-MAT"
      CODPARAM$(2) = "DES-MAT"
      CODPARAM$(3) = "CANTIDAD"
      CODPARAM$(4) = "DESTINO"
      CODPARAM$(5) = "RASO-CLI"
      CODPARAM$(6) = "REF-MAT1"
      CODPARAM$(7) = "PRE-NETO"
      CODPARAM$(8) = "NET-ITEM"
      CODPARAM$(9) = "REF-MAT2"
      CODPARAM$(10) = "USUARIO"
      CODPARAM$(11) = "REF-MAT3"
      CODPARAM$(12) = "DOCU-APL"
      CAPARDOC% = 12
      '
'      CI% = Val(Label3)
'      CAN3 = Val(Label4)
'      DESTI% = Val(Label5)
'      NPRO% = Val(Label7)
'      MOTI% = Val(Label8)
'      PREUNI1# = Val(Label10)
'      REFE1$ = Label9
'      PREUNI1# = Val(Label15)
'      IMPORT1# = Val(Label16)
      '
      DOCPARAM$(1) = CODEXT$(CI%)
      DOCPARAM$(2) = DESCRIT0$(CI%)
      DOCPARAM$(3) = FORMATNUM$(CAN3, "F9.2")
      DOCPARAM$(4) = TRIM$(ECOARCH$("TADESTI", Chr(DESTI%)))
      DOCPARAM$(5) = ""
        If NPRO% > 0 Then DOCPARAM$(5) = RASOCLI$(NPRO% * (-1))
      DOCPARAM$(6) = TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%)))
      DOCPARAM$(7) = FORMATNUM$(PREUNI1#, "F11.4") 'PRUEBA PARA VALOR UNITARIO
      DOCPARAM$(8) = FORMATNUM$(IMPORT1#, "F11.2") 'PRUEBA PARA IMPORTE
      DOCPARAM$(9) = REFE1$ 'PRUEBA DE REFERENCIA
      DOCPARAM$(10) = ""
      DOCPARAM$(11) = "" 'NADA EN E TESORERIA
      DOCPARAM$(12) = "" 'NADA EN  N/D
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      '
      Call PRINTDOC("RECHALIN")   ' RECHAZO EN LÍNEA
      If DONDE% = 2 Then 'IMPRESION REMOTA
        PRINTERPORT$ = ""
        Call SETPRINPORT("") 'RESETEA EL CAMBIO DE PUERTO DE IMPRESIÓN.
      End If
      '
    End If
  '
End Sub
Sub AGREGA_ANOTAORFAB(NORXX$, TXTX$)
    '
    '\\\OT-08-0995-OD-30/10/08///
    If NORXX$ <> "" Then
     SQLX$ = " SELECT * FROM  ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORXX$ & "'"
     Dim ABREORFFFAB As ADODB.Recordset
     Set ABREORFFFAB = New ADODB.Recordset
     ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     On Error Resume Next
     '
     With ABREORFFFAB
      If Not (.BOF And .EOF) Then
         ANOTANTE$ = SAFETEXT$(!AnotaOrf)
         ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
         !AnotaOrf = ANOTA$
         .Update
      End If
     End With
     ABREORFFFAB.Close
     Set ABREORFFFAB = Nothing
     '
    End If
    '
    '\\\OT-08-0995-OD-FIN///
End Sub

Sub AGREGA_ANOTAORFAB_NEW(NORXX$, TXTX$, Optional FF%)
    '
    If NORXX$ <> "" Then
        HOII% = HOY(HOS$)
        FECH$ = HOS$
        If FF% > 0 Then
              FECH$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
        End If
        '
        Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
    
        For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
            TTYY$ = Space$(96)
            If KKI% = 1 Then
               Call REPLA(TTYY$, FECH$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
               Call REPLA(TTYY$, USERTER$, 69, 24)
            End If
            Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
            ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
        Next KKI%
        '
        ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf

        SQLX$ = " SELECT AnotaOrf FROM  ORDEFABR "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORXX$ & "'"
        Dim ABREORFFFAB As ADODB.Recordset
        Set ABREORFFFAB = New ADODB.Recordset
25      On Error Resume Next
        ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With ABREORFFFAB
         If Not (.BOF And .EOF) Then
            NOTAORIG$ = SAFETEXT$(!AnotaOrf)
            ANOTA$ = NOTAORIG$ + ANOTANTE$
            !AnotaOrf = ANOTA$
            .Update
         End If
       End With
       ABREORFFFAB.Close
       Set ABREORFFFAB = Nothing
       '
    End If
    '
    
End Sub
Sub REEMPLAZACOMPRAS(NORFA$, DEPORI%)
    Dim ASIGNO As ADODB.Recordset
    Set ASIGNO = New ADODB.Recordset
    Dim VALRS(256)
    '
    ' DEPURA LOS REGISTROS VACIOS
    CONDIX$ = "IDENREQUE LIKE '" & NORFA$ & "%' "
    CONDIX$ = CONDIX$ + " AND (Kilo_Reque = 0 OR KILO_REQUE IS NULL) "
    CONDIX$ = CONDIX$ + " AND (Kilo_Asig =0 OR KILO_ASIG IS NULL)"
    Call BORRAREGISTROS("BALAREQMO", CONDIX$, RAFE&)
    '
    ' REEMPLAZA LOS ASIGNADOS POR O/COMPRA Y NOTA DE PEDIDO
    SQLX$ = "SELECT * FROM BALAREQMO "
    SQLX$ = SQLX$ + " WHERE IDENREQUE LIKE '" & NORFA$ & "%' "
    SQLX$ = SQLX$ + " AND KILO_ASIG > 0 "
    SQLX$ = SQLX$ + " AND (IdenLote LIKE 'OC%' or IdenLote LIKE 'NP%' OR IDENLOTE LIKE 'SF%')"
    ASIGNO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With ASIGNO
       Do While Not .EOF
          LOCOMP$ = SAFETEXT$(!IdenLote)
          CIXI% = XVALO(!Cod_Inte)
          CANASI = XVALO(!Kilo_Asig)
          '
          ' BUSCAR EN LOTINGRE POR DOCUORIG Y CODIGO
          CONDI$ = "Cod_Inte = " & CIXI% & " AND DocuOrig = '" & LOCOMP$ & "'"
          SQLZ$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) WHERE " & CONDI$
          Dim LOSEL As ADODB.Recordset
          Set LOSEL = FLEXCONN.Execute(FILTSQL$(SQLZ$))
          '
          Do While Not LOSEL.EOF
             IDLO$ = LOSEL!IdenLote
             Call ACSALOTE(CIXI%, IDLO$)
             CONDX$ = "Cod_Inte = " & CIXI% & " AND IdenLote = '" & IDLO$ & "'"
             SALOT = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDX$))
             '
             If SALOT > 0 Then
                If SALOT >= CANASI Then
                    ' REEMPLAZA NOTA DE PEDIDO POR BOLETA DE RECEPCION
                    !IdenLote = IDLO$
                    ' PONE DOCUCOMPRA = DOCUMENTO DE COMPRA
                    !DOCOMPRA = LOCOMP$
                    .Update
                  Else
                    ' PONE KILO_ASIG = CANASI-SALOT
                    On Error Resume Next
                    !Kilo_Asig = CANASI - SALOT
                    On Error GoTo 0
                    .Update
                    For KKI% = 1 To .Fields.Count
                       VALRS(KKI%) = .Fields(KKI% - 1).Value
                    Next KKI%
                    '
                    ' AGREGA UN REGISTRO IGUAL CON kILO_Asig =salot
                    .AddNew   ' YA CONVERTIDO
                    On Error Resume Next
                    !CodiEmpr = CodiEmp%
                    On Error GoTo 0
                    For KKI% = 1 To .Fields.Count
                       .Fields(KKI% - 1).Value = VALRS(KKI%)
                    Next KKI%
                    !Kilo_Asig = SALOT
                    !IdenLote = IDLO$
                    !DOCOMPRA = LOCOMP$
                    .Update
                End If
                Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
             End If
          LOSEL.MoveNext
          Loop
          LOSEL.Close
          Set LOSEL = Nothing
          '
       .MoveNext
       Loop
    End With
    ASIGNO.Close
    Set ASIGNO = Nothing
    '
End Sub

Sub REEMPLAZALOTE(CIXI%, SOLFA$, LOTPRO$, DEPORI%)
    '
    Dim ASIGNO As ADODB.Recordset
    Set ASIGNO = New ADODB.Recordset
    Dim VALRS(256)
    '
    ' DETERMINA SALDO SEGUN LOTINGRE Y BALAREQMO
    Call ACSALOTE(CIXI%, LOTPRO$, "NOMESS")
    SALOT = XVALO(VALOBADA("LOTINGRE", "CANALTA - CANCONSU", "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTPRO$ & "'"))
    SALOT = SALOT - XVALO(VALOBADA("BALAREQMO", "SUM(KILO_ASIG)", "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTPRO$ & "'"))
    If SALOT <= 0 Then Exit Sub
    '
    ' REEMPLAZA LOS ASIGNADOS POR O/COMPRA Y NOTA DE PEDIDO
    SQLX$ = "SELECT * FROM BALAREQMO "
    SQLX$ = SQLX$ + " WHERE KILO_ASIG > 0 "
    SQLX$ = SQLX$ + " AND COD_INTE = " & CIXI% & " AND IdenLote = '" & SOLFA$ & "'"
    ASIGNO.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With ASIGNO
       Do While Not .EOF
          CIXI% = XVALO(!Cod_Inte)
          CANASI = XVALO(!Kilo_Asig)
          '
          If SALOT > 0 Then
             If SALOT >= CANASI Then
                ' REEMPLAZA NOTA DE PEDIDO POR BOLETA DE RECEPCION
                !IdenLote = LOTPRO$
                ' PONE DOCUCOMPRA = DOCUMENTO DE COMPRA
                !DOCOMPRA = SOLFA$
                .Update
                SALOT = SALOT - CANASI
              Else
                ' PONE KILO_ASIG = CANASI-SALOT
                On Error Resume Next
                !Kilo_Asig = CANASI - SALOT
                On Error GoTo 0
                .Update
                For KKI% = 1 To .Fields.Count
                   VALRS(KKI%) = .Fields(KKI% - 1).Value
                Next KKI%
                '
                ' AGREGA UN REGISTRO IGUAL CON kILO_Asig =salot
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CodiEmpr = CodiEmp%
                On Error GoTo 0
                For KKI% = 1 To .Fields.Count
                   .Fields(KKI% - 1).Value = VALRS(KKI%)
                Next KKI%
                !Kilo_Asig = SALOT
                !IdenLote = LOTPRO$
                !DOCOMPRA = SOLFA$
                .Update
                SALOT = 0
             End If
          End If
       .MoveNext
       Loop
    End With
    ASIGNO.Close
    Set ASIGNO = Nothing
    Call ACSALOTE(CIXI%, LOTPRO$, "NOMESS")
    '
End Sub

Sub GENERALISTAREQ(NORFA$, Optional MODO$)
    '
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CIXI% = XVALO(VALOBADA("OrdeFabr", "COD_INTE", CONDI$, "NOMESS"))
    CANTID# = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS"))
    DESTIINDI$ = VALOBADA("ORDEFABR", "REFEREN", CONDI$, "NOMESS")
    '
    DOCUI$ = TRIM$(Left$(NORFA$, 2))
    NUCO$ = TRIM$(Mid$(NORFA$, 4, 6))
    NOSU$ = TRIM$(Mid$(NORFA$, 11))
    '
    CONDI$ = "CodiCompro = '" & DOCUI$ & "' "
    CONDI$ = CONDI$ & " AND NumeCompro = '" & NUCO$ & "' "
    CONDI$ = CONDI$ & " AND CodSucu = '" & NOSU$ & "' "
    '
    If MODO$ = "REFRESH" Then
        ' ACTUALIZA EL LOTINGRE EN BASE A LO QUE LUEGO BORRA EN EL BALAREQMO
        CONDI2$ = CONDI$ & " AND KILO_ASIG > 0"
        Call ACTUREGISTRO("BALAREQMO", "KILO_ASIG", CONDI2$, 0, REGAF&, "NOMESS")
        '
        SQLX$ = "SELECT COD_INTE,IDENLOTE FROM BALAREQMO"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        Set rst = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With rst
          Do While Not .EOF
            CI% = XVALO(!Cod_Inte)
            IDLO$ = SAFETEXT$(!IdenLote)
            Call ACSALOTE(CI%, IDLO$, "NOMESS")
          .MoveNext
          Loop
        End With
        ' ACTUALIZA EL LOTINGRE EN BASE A LO QUE LUEGO BORRA EN EL BALAREQMO
        CONDI2$ = CONDI$ & " AND KILO_REQUE > 0"
        Call BORRAREGISTROS("BALAREQMO", CONDI2$, RAFE&)
        Call COMUNI("Borrada la Anterior Lista de Requerimientos\de la Solicitud de Fabricacion '" + NORFA$ + "'.")
    End If
    '
    CONDI2$ = CONDI$ & " AND KILO_REQUE > 0"
    CANRE& = CantRegistros("BALAREQMO", CONDI2$)
    If CANRE& > 0 Then
       Call MENSERR(24, "Ya hay Lista de Requerimientos Definidos\para la Orden de Fabricación '" & NORFA$ & "'.")
       GoTo 99
    End If
    '
    ' cargar la formula patron
    If CIXI% < 1 Then
       Call MENSERR(24, "Codigo de Producto No Válido\para la Orden de Fabricación '" & NORFA$ & "'.")
       GoTo 99
    End If
    '
    CONDI$ = "CODICONJ = " & CIXI%
    CANRE& = CantRegistros("FORMUPAT", CONDI$)
    If CANRE& < 1 Then
       Call MENSERR(24, "No hay Formula Patrón\para el Código '" & CODEXT$(CIXI%) & "'.")
       GoTo 99
    End If
    '
    LOTX# = LOTESTAN(CIXI%)
    If LOTX# < 1 Then
       Call MENSERR(24, "No hay Lote Standard\para el Código '" & CODEXT$(CIXI%) & "'.")
       GoTo 99
    End If
    '
    SQLX$ = "SELECT * FROM FORMUPAT WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim FMLAPTMP As ADODB.Recordset
    Set FMLAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JR& = 0
    With FMLAPTMP
      Do While Not .EOF
        CODILMTO% = XVALO(!CODIELEM)
        UBRUTO# = XVALO(!USOBRUTO)
        Cantidad = CANTID# * (UBRUTO# / LOTX#)
        '
        JR& = JR& + 1
        XX$ = REGBLAN$("LISTAPIK")
        Call REPLA(XX$, MKI(CODILMTO%), 1, 2)
        Call REPLA(XX$, MKS$(Cantidad), 3, 4)
        Call GRAREG("!LISTAPIK", XX$, JR&)
        .MoveNext
      Loop
    End With
    Call SETULTREG("!LISTAPIK", JR&)

    FMLAPTMP.Close
    Set FMLAPTMP = Nothing
    '
20  Call ABREORFAB
    DORFA$ = VALOBADA("ORDEFABR", "Referen", "IdSubOr = '" & NORFA$ & "'")
    VTB% = VENTABU%("LISTAPIK/TITUPAN=Lista de Requerimientos - " & NORFA$ & " " & DORFA$)
    If VTB% < 0 Then GoTo 99    ' CANCELADO
    '
    ' GUARDA EL REQUERIMIENTO EN EL BALAREQMO
    Screen.MousePointer = 11
    Dim BALARE As ADODB.Recordset
    Set BALARE = New ADODB.Recordset
    SQLX$ = "SELECT * FROM BALAREQMO WHERE IDENREQUE = '" & IDREQ$ & "'"
    BALARE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    DOCUI$ = TRIM$(Left$(NORFA$, 2))
    NUCO$ = TRIM$(Mid$(NORFA$, 4, 6))
    NOSU$ = TRIM$(Mid$(NORFA$, 11))
    IDREQ$ = NORFA$
    '
    For KU& = 1 To ULTREG&("!LISTAPIK")
       XX$ = REGLEIDO$("!LISTAPIK", KU&)
       CODILMTO% = CVI(Mid$(XX$, 1, 2))
       Cantidad = CVS(Mid$(XX$, 3, 4))
       '
       BALARE.AddNew   ' YA CONVERTIDO
             On Error Resume Next
       BALARE!CodiEmpr = CodiEmp%
             On Error GoTo 0
       BALARE!CODICOMPRO = DOCUI$
       BALARE!NUMECOMPRO = XVALO(NUCO$)
       BALARE!CODSUCU = XVALO(NOSU$)
         NORI$ = TRIM$(Str$(KU&))
       BALARE!NuOrItem = KU&
         DOCING$ = IDREQ$ + "/" + NORI$
       BALARE!IdenReque = DOCING$
       BALARE!Cod_Inte = CODILMTO%
       BALARE!IdenLote = ""
       BALARE!Cantidad = Cantidad
       BALARE!CodEnvase = ""
       BALARE!KILO_REQUE = Cantidad
       BALARE!Kilo_Asig = 0
       BALARE!Kilo_Consu = 0
       BALARE!FeReMovi = Now
       BALARE!Fechmov = Now
       BALARE!MarqAsi = 0
       BALARE!CREADOPOR = "FLEX"
       BALARE!DESTINO = NORFA$
       BALARE.Update
    Next KU&
    BALARE.Close
    Set BALARE = Nothing
    Screen.MousePointer = 1
    '
    ' GENERACION DE INDIREQU
    NX$ = Space(64)
    Call REPLA(NX$, IDREQ$, 1, 16)
    Call REPLA(NX$, DESTIINDI$, 17, 48)
    '
    Call FILTERPUTRECORD$("INDIREQU", 1, IDREQ$, NX$)
    '
    Call FRESHECHO("INDIREQU")
    Call CIERRARCH("*.*")
    Call COMUNI("Lista de Requerimientos Creada OK\Para " + NORFA$ + " - " + DORFA$)
    '
99 End Sub
'
Sub ASIGNAREQUERI(NORFA$, DEPORI%, NULIPI&)
    '
    Dim REQUERI As ADODB.Recordset
    Set REQUERI = New ADODB.Recordset
    '
    SQLX$ = "SELECT COD_INTE, IdenReque, SUM(KILO_REQUE) AS TOREQUE, SUM(KILO_ASIG) AS TOASIG FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE idenreque LIKE '" & NORFA$ & "%' "
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, IdenReque "
    Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    DESTIRE$ = ""
    If Left$(NORFA$, 2) = "OF" Or Left$(NORFA$, 2) = "SF" Then
       DESTIRE$ = NORFA$
    End If
    '
    If Left$(NORFA$, 2) = Control("", "CODMORFA") Then
       DESTIRE$ = NORFA$
    End If
    '
    '
    HOII% = HOY(HOS$)
    IDASI$ = Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2) + "-" + TRIM$(Str$(NULIPI&)) + "-"
    NORIT% = 0
    '
    With REQUERI
       If Not (.EOF And .BOF) Then
          Do While Not .EOF
             HUBOMOASI% = 0
             If XVALO(!TOASIG) >= XVALO(!TOREQUE) - 0.05 Then   ' CUBIERTA LA ASIGNACION, PONE MARCA
                  IDREQ$ = SAFETEXT$(!IdenReque)                ' ASIGNA CUMPLIDA Y VA AL SIGUIENTE
                  CODIN% = XVALO(!Cod_Inte)
                  SQLY$ = "UPDATE BALAREQMO "
                  SQLY$ = SQLY$ + " SET MARQASI = 1 "
                  SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "'"
                  FLEXCONN.Execute SQLY$
                  GoTo 19                         ' MOVENEXT
                Else
                  CODIMAT% = XVALO(!Cod_Inte)                  ' CODIGO PRODUCTO REQUERIDO
                  KILPENASI = XVALO(!TOREQUE) - XVALO(!TOASIG)   ' KILOS PENDIENTES DE ASIGNACION
                  ' SE APLICA EL REDONDEO POR LOTE INDIVISIBLE
                  CANTIPEN# = KILPENASI
                  KILPENASI = REDON_INDIVIS(CODIMAT%, CANTIPEN#)
                  '
                  IDREQ$ = SAFETEXT$(!IdenReque)              ' ASIGNA CUMPLIDA Y VA AL SIGUIENTE
                  CONDX$ = "Cod_Inte = " & CODIMAT% & " "
                  CONDX$ = CONDX$ + "AND IdenReque = '" & IDREQ$ & "' "
                  CONDX$ = CONDX$ & "AND KILO_REQUE > 0 "
                  LOPEDIDO$ = VALOBADA("BALAREQMO", "IDENLOTE", CONDX$) ' PARTIDA SOLICITADA
                  If TRIM$(LOPEDIDO$) = "" Then GoTo 25 ' PEDIDO SIN LOTE PREFERIDO
                  '
                  ' PROCESO POR LOTE SOLICITADO
                  Call ACSALOTE(CODIMAT%, LOPEDIDO$, "NOMESS")
                  SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
                  SQLY$ = SQLY$ + "WHERE Cod_Inte = " & CODIMAT% & " "
                  SQLY$ = SQLY$ + "AND (Libera = 'APRO' OR Libera = 'CUAR') "
                  SQLY$ = SQLY$ + "AND LoteCol_Prov = '" & LOPEDIDO$ & "' "
                  SQLY$ = SQLY$ + "AND (CanAlta - CanConsu - CanReser) > 0 "
                  SQLY$ = SQLY$ + "AND DepoUbic = " & DEPORI% & " " 'Lote tiene que estar en Deposito Origen
                  SQLY$ = SQLY$ + "ORDER BY CanSaldo ASC "
                  Dim LODISPO As ADODB.Recordset
                  Set LODISPO = FLEXCONN.Execute(FILTSQL$(SQLY$))
                  '
                  With LODISPO
                     If Not (.EOF And .BOF) Then
                       Do While Not .EOF
                         If KILPENASI < 0.05 Then GoTo 29
                         IDLOTX$ = SAFETEXT$(!IdenLote)
                         CIXI% = CODIMAT%
                         CONDI$ = "Cod_Inte = " & CIXI% & " "
                         CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "' "
                         KILDISLO = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
                         KILAASI = KILDISLO: If KILPENASI < KILDISLO Then KILAASI = KILPENASI
                           LINDI = LINDIVIS(CIXI%)
                           If LINDI > 0 Then
                              KILAASI = LINDI * Int(KILAASI / LINDI)
                           End If
                         '
                         POSILOTE$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
                         CODENVAS$ = VALOBADA("LOTINGRE", "CodEnvase", CONDI$)
                         COEFI = 1
                         KILORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
                         CANORI = XVALO(VALOBADA("LOTINGRE", "Cantidad", CONDI$))
                           If KILORI > 0 Then COEFI = CANORI / KILORI
                         '
                         ' GRABAR LA ASIGNACION EN BALAREQMO
                         CCX1$ = TRIM$(VALOBADA("BALAREQMO", "CodiCompro", CONDX$))
                         NCX1 = XVALO(VALOBADA("BALAREQMO", "NumeCompro", CONDX$))
                         CSX1 = XVALO(VALOBADA("BALAREQMO", "CodSucu", CONDX$))
                         NOI1 = XVALO(VALOBADA("BALAREQMO", "NuOrItem", CONDX$))
                         DESTIPEDIDO$ = VALOBADA("BALAREQMO", "Destino", CONDX$)
                         
                         If KILAASI >= 0.05 Then
                            SQLZ$ = "SELECT * FROM BALAREQMO "
                            SQLZ$ = SQLZ$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            Dim ASIGNA As ADODB.Recordset
                            Set ASIGNA = New ADODB.Recordset
                            ASIGNA.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                            With ASIGNA
                               .AddNew   ' YA CONVERTIDO
                               On Error Resume Next
                               !CodiEmpr = CodiEmp%
                               On Error GoTo 0
                               !CODICOMPRO = CCX1$
                               !NUMECOMPRO = NCX1
                               !CODSUCU = CSX1
                               !NuOrItem = NOI1
                               !IdenReque = IDREQ$
                               !Cod_Inte = CIXI%
                               !IdenLote = IDLOTX$
                               !Posicion = POSILOTE$
                               !Cantidad = KILAASI * COEFI         ' falta dederminar
                               !CodEnvase = CODENVAS$
                               !KILO_REQUE = 0
                               !Kilo_Asig = KILAASI
                               !NuLisPick = NULIPI&
                               !DeLisPick = DELIPI$
                               !Kilo_Consu = 0
                               !FeReMovi = Now
                               !Fechmov = Now
                               !MarqAsi = 0
                               !CREADOPOR = "FLEX"
                               !DESTINO = DESTIPEDIDO$
                                   NORIT% = NORIT% + 1
                               !Ident_Asig = IDASI$ + TRIM$(Str$(NORIT%))
                               !DepoMovi = DEPORI%                'Deposito Origen
                               .Update
                               HUBOASIG% = 1
                               HUBOMOASI% = 1
                            End With
                            ASIGNA.Close
                            Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS") ' ACTUALIZAR CON ACSALOTE
                            '
                            SQLY$ = "UPDATE BALAREQMO "
                            SQLY$ = SQLY$ + " SET NuLisPick = " & NULIPI& & " "
                            SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            SQLY$ = SQLY$ + " AND Cod_Inte = " & CIXI% & " "
                            SQLY$ = SQLY$ + " AND Kilo_Reque > 0 "
                            FLEXCONN.Execute SQLY$
                            '
                            KILPENASI = KILPENASI - KILAASI
                            If KILPENASI < 0.05 Then GoTo 29
                         End If
                         '
                       .MoveNext
                       Loop
                     End If
                  End With
                  LODISPO.Close
                  Set LODISPO = Nothing
                  '
                  SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
                  SQLY$ = SQLY$ + "WHERE Cod_Inte = " & CODIMAT% & " "
                  SQLY$ = SQLY$ + "AND (Libera = 'APRO' OR Libera = 'CUAR') "
                  SQLY$ = SQLY$ + "AND DepoUbic = " & DEPORI% & " " 'Lote tiene que estar en Deposito Origen
                  Set LODISPO = New ADODB.Recordset
                  Set LODISPO = FLEXCONN.Execute(FILTSQL$(SQLY$))
                  With LODISPO
                     Do While Not .EOF
                        IDLOTX$ = SAFETEXT$(!IdenLote)
                        Call ACSALOTE(CODIMAT%, IDLOTX$, "NOMESS")
                     .MoveNext
                     Loop
                  End With
                  LODISPO.Close
                  Set LODISPO = Nothing
                  '
25                VUELTA% = 1
                  '
26                SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
                  SQLY$ = SQLY$ + "WHERE Cod_Inte = " & CODIMAT% & " "
                  SQLY$ = SQLY$ + "AND (Libera = 'APRO' OR Libera = 'CUAR') "
                  SQLY$ = SQLY$ + "AND DepoUbic = " & DEPORI% & " " 'Lote tiene que estar en Deposito Origen
                  SQLY$ = SQLY$ + "AND (CanAlta - CanConsu - CanReser) > 0 "
                  If DESTIRE$ <> "" And VUELTA% = 1 Then
                     SQLY$ = SQLY$ + " AND COD_DESTI = '" & DESTIRE$ & "' "
                  End If
                  SQLY$ = SQLY$ + "ORDER BY FechVenc ASC "   ' POR CRITERIO FEFO
                  Set LODISPO = FLEXCONN.Execute(FILTSQL$(SQLY$))
                  '
                  With LODISPO
                     If Not (.EOF And .BOF) Then
                       Do While Not .EOF
                         If KILPENASI < 0.05 Then GoTo 29
                         If SAFETEXT$(!Cod_Desti) <> "" Then
                            If SAFETEXT$(!Cod_Desti) <> DESTIRE$ Then
                               GoTo 27
                            End If
                         End If
                         IDLOTX$ = SAFETEXT$(!IdenLote)
                         CIXI% = CODIMAT%
                         CONDI$ = "Cod_Inte = " & CIXI% & " "
                         CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "' "
                         KILDISLO = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
                         KILAASI = KILDISLO: If KILPENASI < KILDISLO Then KILAASI = KILPENASI
                           LINDI = LINDIVIS(CIXI%)
                           If LINDI > 0 Then
                              KILAASI = LINDI * Int(KILAASI / LINDI)
                           End If
                         POSILOTE$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
                         CODENVAS$ = VALOBADA("LOTINGRE", "CodEnvase", CONDI$)
                         COEFI = 1
                         KILORI = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
                         CANORI = XVALO(VALOBADA("LOTINGRE", "Cantidad", CONDI$))
                           If KILORI > 0 Then COEFI = CANORI / KILORI
                           '
                         ' GRABAR LA ASIGNACION EN BALAREQMO
                         CCX1$ = TRIM$(VALOBADA("BALAREQMO", "CodiCompro", CONDX$))
                         NCX1 = XVALO(VALOBADA("BALAREQMO", "NumeCompro", CONDX$))
                         CSX1 = XVALO(VALOBADA("BALAREQMO", "CodSucu", CONDX$))
                         NOI1 = XVALO(VALOBADA("BALAREQMO", "NuOrItem", CONDX$))
                         DESTIPEDIDO$ = VALOBADA("BALAREQMO", "Destino", CONDX$)
                        
                         If KILAASI >= 0.05 Then
                            SQLZ$ = "SELECT * FROM BALAREQMO "
                            SQLZ$ = SQLZ$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            Set ASIGNA = New ADODB.Recordset
                            ASIGNA.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                            With ASIGNA
                               .AddNew   ' YA CONVERTIDO
                               On Error Resume Next
                               !CodiEmpr = CodiEmp%
                               On Error GoTo 0
                               !CODICOMPRO = CCX1$
                               !NUMECOMPRO = NCX1
                               !CODSUCU = CSX1
                               !NuOrItem = NOI1
                               !IdenReque = IDREQ$
                               !Cod_Inte = CIXI%
                               !IdenLote = IDLOTX$
                               !Posicion = POSILOTE$
                               !Cantidad = KILAASI * COEFI
                               !CodEnvase = CODENVAS$
                               !KILO_REQUE = 0
                               !Kilo_Asig = KILAASI
                               !NuLisPick = NULIPI&
                               !DeLisPick = DELIPI$
                               !Kilo_Consu = 0
                               !FeReMovi = Now
                               !Fechmov = Now
                               !MarqAsi = 0
                               !CREADOPOR = "FLEX"
                               !DESTINO = DESTIPEDIDO$
                                   NORIT% = NORIT% + 1
                               !Ident_Asig = IDASI$ + TRIM$(Str$(NORIT%))
                               .Update
                               HUBOASIG% = 1
                               HUBOMOASI% = 1
                            End With
                            '
                            ASIGNA.Close
                            Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS") ' ACTUALIZAR CON ACSALOTE
                            '
                            SQLY$ = "UPDATE BALAREQMO "
                            SQLY$ = SQLY$ + " SET NuLisPick = " & NULIPI& & " "
                            SQLY$ = SQLY$ + " WHERE IdenReque = '" & IDREQ$ & "' "
                            SQLY$ = SQLY$ + " AND Cod_Inte = " & CIXI% & " "
                            SQLY$ = SQLY$ + " AND Kilo_Reque > 0 "
                            FLEXCONN.Execute SQLY$
                            '
                            KILPENASI = KILPENASI - KILAASI
                            If KILPENASI < 0.05 Then GoTo 29
                         End If
                         '
27                     .MoveNext
                       Loop
                     End If
                  End With
                  LODISPO.Close
                  Set LODISPO = Nothing
                  If VUELTA% = 1 Then
                     VUELTA% = 2: GoTo 26
                  End If
                  '
29           End If
             '
             If HUBOMOASI% = 1 Then
               CONDI$ = "COD_INTE = " & CODIMAT% & " AND IdenReque = '" & IDREQ$ & "'"
               Call ACTUREGISTRO("BALAREQMO", "NulisPick", CONDI$, NULIPI&, REGAFE&, "NOMESS")
               Call ACTUREGISTRO("BALAREQMO", "DelisPick", CONDI$, DELIPI$, REGAFE&, "NOMESS")
             End If
             '
19        .MoveNext
          Loop
       End If
    End With
    '
'    SQLX$ = "SELECT IdenReque, kilo_reque, KILO_SALDO FROM BALAREQMO WITH(NOLOCK) "
'    SQLX$ = SQLX$ + " WHERE idenreque LIKE '" & NORFA$ & "%' "
'    SQLX$ = SQLX$ + " AND KILO_REQUE > 0"
'    Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    '
'    With REQUERI
'      If Not (.BOF And .EOF) Then
'        Do While Not .EOF
'          IDREQ$ = !IdenReque
'          TOASI = VALOBADA("BALAREQMO", "SUM(KILO_ASIG)", "IDENREQUE = '" & IDREQ$ & "'")
'          SALDOX$ = TRIM$(FORMATNUM$(XVALO(!KILO_REQUE) - TOASI, "F15.6"))
'          If XVALO(SALDOX$) <= 0 Then SALDOX$ = "0"
'          If Abs(!KILO_SALDO - XVALO(SALDOX$)) >= 0.0005 Then
'             Call ACTUREGISTRO("BALAREQMO", "KILO_SALDO", "IDENREQUE = '" & IDREQ$ & "' AND KILO_REQUE > 0", SALDOX, RAFE&)
'          End If
'        .MoveNext
'        Loop
'      End If
'    End With
'    REQUERI.Close
'    Set REQUERI = Nothing
    '
    SQLX$ = "SELECT IdenReque, SUM(kilo_reque) AS KIREQ, SUM(KILO_ASIG) AS KIASI,SUM(KILO_SALDO) AS KISAL FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE idenreque LIKE '" & NORFA$ & "%' "
    SQLX$ = SQLX$ + " GROUP BY IDENREQUE"
    Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With REQUERI
      Do While Not .EOF
        SALTEORI = XVALO(!KIREQ) - XVALO(!KIASI)
        If SALTEORI < 0 Then SALTEORI = 0
        If Abs(XVALO(!KISAL) - SALTEORI) >= 0.0005 Then
           SALDOX$ = TRIM$(FORMATNUM$(SALTEORI, "F15.6"))
           Call ACTUREGISTRO("BALAREQMO", "KILO_SALDO", "IDENREQUE = '" & IDREQ$ & "' AND KILO_REQUE > 0", SALDOX, RAFE&)
        End If
      .MoveNext
      Loop
    End With
    REQUERI.Close
    Set REQUERI = Nothing
    '
    'Pone en cero cantidad neta y bruta
    SQLX$ = "SELECT IdenReque FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE idenreque LIKE '" & NORFA$ & "%' "
    'SQLX$ = SQLX$ + " AND (CANTIDAD <> 0 OR CODENVASE <> '') "
    SQLX$ = SQLX$ + " AND CODENVASE <> '' "
    Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REQUERI
      Do While Not .EOF
        IDREQ$ = SAFETEXT$(!IdenReque)
'        Call ACTUREGISTRO("BALAREQMO", "Cantidad", "idenreque = '" & IDREQ$ & "' ", 0, RAFE&)
        Call ACTUREGISTRO("BALAREQMO", "CodEnvase", "idenreque = '" & IDREQ$ & "' ", "", RAFE&)
      .MoveNext
      Loop
    End With
    REQUERI.Close
    Set REQUERI = Nothing
    '
    SQLX$ = "SELECT IdenReque, BALAREQMO.IdenLote AS IDLOT, NuLisPick ,CanAlta , CanConsu, BALAREQMO.CANTIDAD AS CANAA FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " INNER JOIN LOTINGRE "
    SQLX$ = SQLX$ + " ON BALAREQMO.IDENLOTE = LOTINGRE.IDENLOTE "
    SQLX$ = SQLX$ + " WHERE idenreque LIKE '" & NORFA$ & "%' "
    SQLX$ = SQLX$ + " AND Kilo_Asig > 0 "
'    SQLX$ = SQLX$ + " AND (ABS(CanAlta - CanConsu - BALAREQMO.CANTIDAD) > 0.0005 OR NULISPICK < 1)"
    SQLX$ = SQLX$ + " AND ABS(CanAlta - CanConsu - BALAREQMO.CANTIDAD) > 0.0005 "
    Set REQUERI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With REQUERI
       Do While Not .EOF
 AAA1 = XVALO(!CanAlta)
 bbb1 = XVALO(!CanConsu)
 CCC1 = XVALO(!CANAA)
          IDREQ$ = SAFETEXT$(!IdenReque)
          LOTASI$ = SAFETEXT$(!IDLOT)
          SALOT = XVALO(VALOBADA("LOTINGRE", "CanAlta - CanConsu", "IDENlote = '" & LOTASI$ & "'"))
          SALDOX$ = TRIM$(FORMATNUM$(SALOT, "F15.6")): If SALDOX = "" Then SALDOX = 0
          Call ACTUREGISTRO("BALAREQMO", "Cantidad", "IDENREQUE = '" & IDREQ$ & "' AND KILO_ASIG > 0", SALDOX, RAFE&)
'          If SALOT > 0 Then
'            If XVALO(!NuLisPick) < 1 Then
'              CONDIX$ = "IDENREQUE = '" & IDREQ$ & "'"
'              CONDIX$ = CONDIX$ + " AND KILO_ASIG > 0"
'              CONDIX$ = CONDIX$ + " AND IDENLOTE = '" & LOTASI$ & "'"
'              Call ACTUREGISTRO("BALAREQMO", "NuLisPick", CONDIX$, NULIPI&, RAFE&)
'            End If
'          End If
          '
       .MoveNext
       Loop
    End With
    REQUERI.Close
    Set REQUERI = Nothing
    '
    '
    'Registra en NEBRUNE en caso de tener criterio de rep. <> 1
    '
    CONDI$ = "DESTINO='" & NORFA$ & "'"
    Call BORRAREGISTROS("NEBRUNE", CONDI$, REG&)
    '
    Dim RESERX As ADODB.Recordset
    Set RESERX = New ADODB.Recordset
    SQL1$ = "SELECT * FROM NEBRUNE WHERE COD_INTE = 0"
    RESERX.Open FILTSQL$(SQL1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    Dim BALAQ As ADODB.Recordset
    Set BALAQ = New ADODB.Recordset
    SQL$ = "select * from balareqmo where DESTINO ='" & NORFA$ & "' AND KILO_SALDO > 0"
    Set BALAQ = FLEXCONN.Execute(FILTSQL$(SQL$))
    If BALAQ.BOF = True And BALAQ.EOF = True Then GoTo 99
    '
    With BALAQ
      .MoveFirst
      Do While Not (.EOF)
      '
      CODX% = XVALO(!Cod_Inte)
      REF$ = VALOBADA("OrdeFabr", "Referen", "IdSubOr='" & NORFA & "'")
      
      FECHENTRE = VALOBADA("OrdeFabr", "FechEnSol", "IdSubOr='" & NORFA & "'")
      If FECHENTRE = "" Then
         HOII% = HOY(HOS$)
         FECHENTRE = CVDAT(HOII%)
      End If
      
      XX = CRITREP%(CODX%)
      '
      If CRITREP%(CODX%) <> 1 Then
         RESERX.AddNew   ' YA CONVERTIDO
               On Error Resume Next
         RESERX!CodiEmpr = CodiEmp%
               On Error GoTo 0
         RESERX!Cod_Inte = CODX%
         RESERX!CANT_REQUE = XVALO(!KILO_SALDO)
         RESERX!Fech_Nec = FECHENTRE
         RESERX!DESTINO = SAFETEXT(!DESTINO)
         RESERX!REFEREN = REF$
         RESERX.Update
      End If
      .MoveNext
      Loop
    End With
    '
    RESERX.Close
    Set RESERX = Nothing
    BALAQ.Close
    Set BALAQ = Nothing
99
    
    '
End Sub
'
Sub ASIGNAMAT(NORFA$)
   'Asignacion de Materiales
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then GoTo 99
    DepOrig% = XVALO(DEPO$)
    
25  ' CALCULO POR ORDENES DE FABRICACION
    NULIPI& = 1 + XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
    DELIPI$ = SAFETEXT$(Now)
    '
    DELIPI$ = DELIPI$ + " - " + NORFA$
    Call REEMPLAZACOMPRAS(NORFA$, DEPORI%)
    NULIPI& = 0
    Call ASIGNAREQUERI(NORFA$, DepOrig%, NULIPI&)
    '
40  UNULIPI& = XVALO(VALOBADA("BALAREQMO", "MAX(NuLisPick)", "Kilo_Asig > 0"))
    If UNULIPI& = NULIPI& Then
       Call FILTERPUTRECORD("INDIPICK", 1, MKS$(NULIPI&), MKS$(NULIPI&) + DELIPI$)
       Call FILESORT("INDIPICK", "", 1, 1, "ASC")
       Call FRESHECHO("INDIPICK")
    End If
    '
    Call FRESHECHO("INDIPICK")
    '
    Call COMUNI("Proceso de Cálculo Completo.")

99
End Sub

Sub REGENOPER(SOLFA$)
    '
    SQLX$ = "DELETE FROM OPERFAB "
    SQLX$ = SQLX$ + "WHERE IdSubOr = '" & SOLFA$ & "'"
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    CONDI$ = "IdSubOr = '" & SOLFA$ & "'"
    CIJK% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
    CAPROYE# = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$, "NOMESS"))
    SQLX$ = "SELECT * FROM SecOper WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE CodIConj=" & CIJK% & " "
    SQLX$ = SQLX$ + " AND StatOper = 1 "
    SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
    '
    Dim SECOTEMP As ADODB.Recordset
    Set SECOTEMP = New ADODB.Recordset
    Set SECOTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Dim OPERFFFAB As ADODB.Recordset
    Set OPERFFFAB = New ADODB.Recordset
    SQLX$ = "SELECT * FROM OperFab WHERE IdSubOr = '" & SOLFA$ & "'"
    OPERFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With SECOTEMP
       Do While Not .EOF
          NO% = XVALO(!NUMEOPER)
          
          CAREALI# = CANREALI(SOLFA$, NO%)
          CAPENDE# = MAXIVAL(CAPROYE# - CAREALI#, 0)
          '
          ATO% = XVALO(!AltOper)
          If ATO% < 1 Then   ' SOLO OPERACIONES BASICAS
             DOPE$ = SAFETEXT$(!DESCOPER)
             EQ1$ = SAFETEXT$(!CodMaq0)
             EQAUX$ = ""
                If Not IsNull(!CodMaq1) Then EQAUX$ = AJUSTI$(!CodMaq1, 6)
                If Not IsNull(!CodMaq2) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq2, 6)
                If Not IsNull(!CodMaq3) Then EQAUX$ = EQAUX$ + AJUSTI$(!CodMaq3, 6)
             LOTESTA = MAXIVAL(XVALO(!LOTESTAN), 1)
             CALOTES = 1
             CAOPS = XVALO(!CantiOps): If CAOPS < 0.01 Then CAOPS = 1
             CICLOHE = XVALO(!PieCiclo): If CICLOHE < 1 Then CICLOHE = 1
             MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
                 If CAREALI# > 0.05 Then MINPREP = 0
             MINFIJO = 60 * XVALO(!HorsFijo) * CALOTES
             MINVARI = CAPENDE# * XVALO(!MinuStan)
             MINHERR = CAPENDE# * 60 * XVALO(!HorsMHer) / CICLOHE
             MINDEMO = (MINFIJO + MINVARI + MINHERR) * XVALO(!PorDemor) / 100
             MINLIMP = 60 * XVALO(!HorsLimp) * CALOTES
             MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO + MINLIMP
             HORAPROD = MINLOTE / 60
             '
             BVOF% = BINVAL%(AJUSTI$(SOLFA$, 12))
             NUOROP% = 0
             FEAL% = HOY(HOS$)
             HOAL% = Int(HORANUM(Time$))
             USUA% = USNBR% Mod 100
             DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
             INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
             PROGBLA$ = Space$(6) + String$(26, 0)
             '
             With OPERFFFAB
                .AddNew   ' YA CONVERTIDO
                   On Error Resume Next
                !CodiEmpr = CodiEmp%
                   On Error GoTo 0
                !IDSUBOR = SOLFA$
                !Cod_Inte = CIJK%
                !COD_EXTE = CODEXT$(CIJK%)
                !DESCRIP = DESCRIT0$(CIJK%)
                !NUMEOPER = NO%
                !DESCOPER = DOPE$
                !MAQUINA = EQ1$
                !DescMaq = ECOARCH("PADMAQ", EQ1$)
                !EquiAux = EQAUX$
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
                !OBSERVA = " "
                !HoBatchRec = ""
                .Update
             End With
          End If
       .MoveNext
       Loop
    End With
    OPERFFFAB.Close
    Set OPERFFFAB = Nothing
    SECOTEMP.Close
    Set SECOTEMP = Nothing
    '
End Sub
'
Sub GRABAARMADOR(NORFA$)
       '
       Call ABREORFAB
       '
       'VALIDO QUE TENGA CONFIGURADO UN DEPOSITO DE ARMADORES
       DEPARMA% = XVALO(Control("", "DEPOARMA"))
       If DEPARMA% < 1 Then
          Call COMUNI("No se puede Asignar Armador\Si no Tiene Configurado un Depósito de Armadores Externos")
          Exit Sub
       End If
       '
       'VERIFICO QUE NO TEMGA UN ARMADOR ASOCIADO
       CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
       ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
       If ARMAD% > 0 Then
          Call COMUNI("La Orden de Fabricación " & NORFA$ & "\Ya Tiene Asociado al Armador Externo")
          Exit Sub
       End If
       '
       'VALIDO QUE NO ESTE ANULADA
       CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND StatuOrf > 8 "
       OFABRANU& = CantRegistros("ORDEFABR", CONDI$)
       If OFABRANU& > 0 Then
          Call COMUNI("Imposible Asignar Armador a esta Orden de Fabricación " & NORFA$ & "\La misma se Encuentra Anulada")
          Exit Sub
       End If
       '
       'VALIDO QUE NO TENGA CIERRES PARCIALES
       CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND (CANAPRO > 0.05 OR CANRECH > 0.05) "
       OFABRAPRO& = CantRegistros("ORDEFABR", CONDI$)
       If OFABRAPRO& > 0 Then
          Call COMUNI("Imposible Asignar Armador a esta Orden de Fabricación " & NORFA$ & "\Ya Tiene Realizados Cierres Parciales o Totales")
          Exit Sub
       End If
       '
       CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
       CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$))
       '
       'VALIDO QUE SE HAYAN CREADO LOS NUMEROS DE SERIE, solo para trazables nivel 2
       If ESTRAZABLE%(CIXI%) = 2 Then
          CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
          NUSE& = CantRegistros&("TANUMSE", CONDI$)
          If NUSE& < 1 Then
             Call COMUNI("Imposible Asignar Armador\No se Han Generado los Números de Serie")
             Exit Sub
          End If
       End If
       '
       'PRESENTO EL OBJETO
       DepOrig% = DEPARMA%
       Call REPLA(VDEF$, NORFA$, 1, 12)
       Call REPLA(VDEF$, Chr$(DEPARMA%), 13, 1)
       '
30     OBX$ = OBJPLA$("ORFAARMA", VDEF$)
       If OBX$ = "" Then Exit Sub
       '
       DEPOARMA% = Asc(Mid$(OBX$, 13, 1))
       ARMAD% = CVI(Mid$(OBX$, 14, 2))
       
       ' VALIDO EL ARMADOR
       If ECOARCH$("ARMADORS", MKI$(ARMAD%)) = "" Then
          Call COMUNI(" Armador no Válido")
          VDEF$ = OBX$
          GoTo 30
       End If
       '
       TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
       ARMADOR$ = TRIM$(Mid$(TX$, 3, 64))
       proveed% = CVI(Mid$(TX$, 68, 2))
       TIPOARM$ = TRIM$(UCase$(Mid$(TX$, 67, 1)))
       '
       'Válido que el armador sea externo
       If TIPOARM$ <> "E" Then
        Call COMUNI("Tipo de Armador \(" & ARMAD% & " " & ARMADOR$ & ") No Valido")
           GoTo 30
       End If
       '
       If proveed% < 1 Then
          Call COMUNI("Falta Asignar Proveedor al Armador " & ARMAD% & "\" & ARMADOR$)
          GoTo 30
       End If
       '
       'GENERACION DE ORDEN DE COMPRA e IMPRESION DE ORDEN DE SERVICIO
31     NUOCONUE& = PRONUOCO&
       'SE MODIFICA PARA QUE QUEDE RESERVADO EL NUMERO DE ORDEN DE COMPRA QUE SE UTILIZA MAS ADELANTE
       If CantRegistros("OCOMENCA", "NUME_OCOM=" & NUOCONUE&, "NOMESS") > 0 Then GoTo 31
       Call CreaRegistro("OCOMENCA", "NUME_OCOM", NUOCONUE&, "NOMESS")
       NORSE& = PRONUORSE&
       '
       SQLX$ = "SELECT Armador, SecciPro, AnotaOrf, CanProy, Cod_Inte FROM ORDEFABR "
       SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
       Dim RS1 As ADODB.Recordset
       Set RS1 = New ADODB.Recordset
20     On Error Resume Next
       RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       '
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 20
       End If
       On Error GoTo 0
       CUERPO$ = "Asignación de Armador Externo" & vbCrLf
       CUERPO$ = CUERPO$ + " (" & ARMAD% & " " & TRIM$(ECOARCH$("ARMADORS", MKI$(ARMAD%))) & " )" & vbCrLf
       CUERPO$ = CUERPO$ + " O.Compra Nro.: " & NUOCONUE& & vbCrLf
       CUERPO$ = CUERPO$ + " O.Servicio Nro.: " & NORSE&
       HISTO$ = FORMATEA_ANOTADOR$(CUERPO$)
       With RS1
          Do While Not .EOF
            !ARMADOR = ARMAD%
            !SecciPro = 0
            !AnotaOrf = SAFETEXT$(!AnotaOrf) & HISTO$
            CANTIEQUI = XVALO(!CanProy)
            CI% = XVALO(!Cod_Inte)
            .Update
            .MoveNext
          Loop
       End With
       RS1.Close
       Set RS1 = Nothing
       '
       If ESTRAZABLE%(CI%) = 2 Then
           SQLX$ = "SELECT UBICACION, Aext_Ensa, Historial, NumeroSerie FROM TANUMSE "
           SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' AND DEPONUMSE <= 0 AND COD_INTE = " & CI%
           Dim RS2 As ADODB.Recordset
           Set RS2 = New ADODB.Recordset
27         On Error Resume Next
           RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           '
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                 On Error GoTo 0
                 Call CapturaErrorOpen
                 GoTo 27
           End If
           On Error GoTo 0
           RACL$ = RASOCLI$(proveed% * (-1))
           CUERPO$ = "Asignación Amador: " & Left$(ARMADOR$, 30) & vbCrLf
           CUERPO$ = CUERPO$ + "Prov.: " & proveed% & " - " & Left$(RACL$, 40) & vbCrLf
           CUERPO$ = CUERPO$ + "O.Compra Nro.: " & NUOCONUE& & vbCrLf
           CUERPO$ = CUERPO$ + "O.Servicio Nro.: " & NORSE&
           HISTO$ = FORMATEA_ANOTADOR$(CUERPO$)
           With RS2
              Do While Not .EOF
                !UBICACION = RACL$
                !Aext_Ensa = ARMADOR$
                NSERIE$ = SAFETEXT$(!NUMEROSERIE)
                !Historial = SAFETEXT$(!Historial) & HISTO$
                .Update
                .MoveNext
              Loop
           End With
           RS2.Close
           Set RS2 = Nothing
       End If
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     Screen.MousePointer = 11
     '
     'GENERO ORDEN DE COMPRA
     OK% = 1
     '
     OBSERX$ = "Orden de Servicio " + TRIM$(Str$(NORSE&)) + " - Orden de Fabricación " + NORFA$
     '
     'GRABO ORDEN DE COMPRA EN OCOMDETA
     SQLX$ = "SELECT * FROM OCOMDETA WHERE NUME_OCOM < 1 " 'CONDICION INEXISTENTE
     Dim OCOMTMP As ADODB.Recordset
     Set OCOMTMP = New ADODB.Recordset
25   On Error Resume Next
     OCOMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0
     '
     With OCOMTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = NUOCONUE&
        !Femi_Ocom = CVDAT(HO%)
        !Npro_Ocom = proveed%
        !Cod_Inte = CI%
        !Uni_Mas = UNIMED$(CI%)
        !cant_ocom = CANTIEQUI
        !Uni_Com = UNIMED$(CI%)
        !Coef_Unids = 1
            LT% = LEAD_TIME%(CI%)
        !Fentre_Sol = CVDAT(DIASPOST(HO%, LT%))
        !CONSIMAT = 1
        !Cant_Rec = 0
        !Cod_Oper = 0
        !Doc_Orig = NORFA$
        !N_ORSERV = NORSE&
        '
        !NuOrd_Item = 1
        !Raso_Prov = RASOCLI$((-1) * proveed%)
        !COD_EXTE = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = ""
        !Sect_Ocom = ""
        !Medi_Ocom = ""
        !Valo_Unid = 0
        !Unids_Ocom = CANTIEQUI
        !Desc_Item = 0
        !inet_item = 0
        !Tsum_Ocom = 0
        !Stat_Ocom = 0
          MONEX% = MONECOSTO%(CI%): If MONEX% < 1 Then MONEX% = 1
        !Mone_Emis = MONEX%
        !mone_cos = MONEX%
        !IContab = 0
            PreUnid# = COSUNI(CI%) * TICAMONE(MONEX%) 'SE AGREGA COSTO DE A.MAESTRO A PEDIDO DE GUILLERMO STERN
        !Pre_Unit = PreUnid#
        !prun_neto = PreUnid#
        .Update
     End With
     OCOMTMP.Close
     Set OCOMTMP = Nothing
     '
     'GRABO ORDEN DE COMPRA EN OCOMENCA
     FechEmis% = HO%
     FEX = FechEmis%
     REFE$ = FECHATEX$(FEX) + " - " + TRIM$(NUORFAB$) + " - " + RASOCLI$((-1) * proveed%)
     CPAG% = CPAGCLI%((-1) * proveed%)
     '
     'SE MODIFICA PORQUE MAS ARRIBA SE UTILIZA UN CREAREGISTRO PARA RESERVAR EL NUMERO EN OCOMENCA
     'SQLX$ = "SELECT * FROM OCOMENCA WHERE NUME_OCOM < 1" 'CONDICION INEXISTENTE
     SQLX$ = "SELECT * FROM OCOMENCA WHERE NUME_OCOM =" & NUOCONUE& 'CONDICION INEXISTENTE
     Dim OCOMENCTMP As ADODB.Recordset
     Set OCOMENCTMP = New ADODB.Recordset
26   On Error Resume Next
     OCOMENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
     End If
     On Error GoTo 0
     
     With OCOMENCTMP
        'SE COMENTA PORQUE AHORA SOLO ACTUALIZA
        '.AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = NUOCONUE&
        !Refe_Ocom = Left$(REFE$, 64)
        !Femi_Ocom = CVDAT(FechEmis%)
        !Npro_Ocom = proveed%
        !Tsum_Ocom = 2 'FASSON
        !Stat_Ocom = 0 'ABIERTA
        !Cpag_Ocom = CPAG%
        !AtSr_Ocom = Left$(ATEN$, 64)
        !Obse_Ocom = OBSERX$
        !Raso_Prov = Left$(RASOCLI$((-1) * proveed%), 48)
        !Fgen_Ocom = Now
        !Usua_Gen = USNBR%
        !tcam_ocom = 1
        !TOBRU_OCOM = 0
        !IDes_Ocom = 0
        !Iiva_Ocom = 0
        !Total_Ocom = 0
        !Embala_Ocom = 0
        !Obse_Ocom = ""
        !Pdes_Ocom = ""
        !Lentrega = ""
        !mone_cos = 1
        !Mone_Emis = 1
        !Anex_Ocom = ""
        !Iotr1_Ocom = 0
        .Update
     End With
     OCOMENCTMP.Close
     Set OCOMENCTMP = Nothing
     'FIN ORDEN DE COMPRA
     
     'IMPRESION DE ORDEN DE COMPRA
     
     HOII% = HOY(HOS$)
     FORIPRE$ = TRIM$(Control$("", "FORMOCOM"))
     If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
           Exit Sub
         End If
         CANCELPRINT% = 0
         If Control$("OCOMPRA1", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
         End If
         '
         Screen.MousePointer = 11
         FECHDOC$ = FETEXTO$(HOII%)
         NUMEDOC$ = TRIM$(Str$(NUOCONUE&))
         While Len(NUMEDOC$) < 6
           NUMEDOC$ = "0" + NUMEDOC$
         Wend
         TIPODOC$ = "Nota de Pedido"
         '
         CODPARAM$(1) = "COND-PAG"
         CODPARAM$(2) = "OBSERVA"
         CODPARAM$(3) = "IMP-NETO"
         CODPARAM$(4) = "POR-IVA"
         CODPARAM$(5) = "IMP-IVA"
         CODPARAM$(6) = "IMP-TOTA"
         CODPARAM$(7) = "VENDEDOR"
         CODPARAM$(8) = "LIST-PRE"
         CODPARAM$(9) = "ZONA-VEN"
         CODPARAM$(10) = "FECH-LAN"
         CODPARAM$(11) = "IMPNEDES"
         CODPARAM$(12) = "TRANSPOR"
         CODPARAM$(13) = "POR-DESC"
         CODPARAM$(14) = "IMP-DESC"
         CAPARDOC% = 14
         '
         DOCPARAM$(1) = FORPAG$
         DOCPARAM$(2) = OBSE$
         DOCPARAM$(3) = Str$(TOTANE#)
         DOCPARAM$(4) = Str$(ALIVA)
         DOCPARAM$(5) = Str$(TOTIVA#)
         DOCPARAM$(6) = Str$(TOTATO#)
         DOCPARAM$(7) = Mid$(X1$, 147, 20)
         DOCPARAM$(8) = Mid$(X1$, 67, 40)
         DOCPARAM$(9) = Mid$(X1$, 167, 40)
         DOCPARAM$(10) = Mid$(X1$, 107, 40)
         DOCPARAM$(11) = Str$(TOTANEDES#)
         DOCPARAM$(12) = TRIM$(Mid$(ENCAOCOM$, 163, 30)) ' ATENCION SR.
         DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
         DOCPARAM$(14) = Str$(IDESCUEN#)
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "UNIMED"
         CODTABU1$(4) = "CANTIDAD"
         CODTABU1$(5) = "FECH-VEN"
         CODTABU1$(6) = "REF-MAT1"
         CODTABU1$(7) = "PRE-LIST"
         CODTABU1$(8) = "PORDESCU"
         CODTABU1$(9) = "PRE-NETO"
         CODTABU1$(10) = "NET-ITEM"
         CODTABU1$(11) = "CAKILOS"
         CODTABU1$(12) = "F-PLANO"
         CODTABU1$(13) = "ORD-ITEM"
         CACOLTAB1% = 13
         '
         DESTIDOC% = (-1) * proveed%
         CAITAB1% = 0
         CAITCRU% = 0
         '
         JJ& = 0
         
         CI0% = CI%
         CANTU = CANTIEQUI
         PreUnid# = 0 ' PARA QUE NO SALGA CON PRECIO EN LA ORDEN DE COMPRA ,
         PORDE = 0
         FEX = HOII%
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = Str$(CANTU)
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Str$(CONSIMAT%)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
         TETABU1$(11, CAITAB1%) = Str$(CANTU * PESUNI(CI0%))
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(CAITAB%))
         '
         CAITAB1% = CAITAB1% + 1
         For KKII% = 1 To 13: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
         TETABU1$(2, CAITAB1%) = ""
         '
         Call PRINTDOC("FORMOCOM")   ' Orden de Compra
     End If
     '
     'LIMPIO LAS VARIABLES DE IMPRESION
     Call LIMPIATETAB
     '
     'IMPRESION DE ORDEN DE SERVICIO
     '
     FECHA% = HOY(HOS$)
     FORIPRE$ = TRIM$(Control$("", "FORSERVI"))
     If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
          RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
          Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Servicio:\  \'" + RMCC$ + "'.")
          Exit Sub
        End If
        CANCELPRINT% = 0
        If Control$("ORSERVI1", "SUPRINT") = "SI" Then
          CANCELPRINT% = 1
        End If
        '
        FEX = FECHA%
        FECHDOC$ = FECHATEX$(FEX)
        NUMEDOC$ = TRIM$(Str$(NORSE&))
        TIPODOC$ = "Orden de Servicio"
        '
        CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
        CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
        CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
        CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
        CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
        CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
        CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
        CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
        FEX = FENTRESOL%
        CODPARAM$(9) = "FECH-VEN": DOCPARAM$(9) = FECHATEX$(FEX)
        CODPARAM$(10) = "DOCUORIG": DOCPARAM$(10) = NORFA$
        CODPARAM$(11) = "NOP-OPER": DOCPARAM$(11) = Str$(NOPE%)
        CODPARAM$(12) = "DES-OPER": DOCPARAM$(12) = DOPE$
        CODPARAM$(13) = "CAN-COMP": DOCPARAM$(13) = Str$(CAINF#)
        CODPARAM$(14) = "VENDEDOR": DOCPARAM$(14) = TRIM$(Mid$(ENCAOCOM$, 163, 30))          ' ATENCION SR.
        CAPARDOC% = 14
        '
        CODTABU1$(1) = "COD-MAT"
        CODTABU1$(2) = "DES-MAT"
        CODTABU1$(3) = "NRO-LOTE"
        CODTABU1$(4) = "CAKILOS"
        CODTABU1$(5) = "UNIMED"
        CACOLTAB1% = 5
        '
        DESTIDOC% = (-1) * proveed%
        CAITAB1% = 1
        '
        UNID$ = UNIMED$(CI%)
        CANS$ = FORMATNUM$(CAN#, "F12." & CNTDC$(CI%))
        KILO$ = FORMATNUM$(CAN# * PESUNI(CI%), "F12." & CNTDC$(CI%))
        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
        TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
        TETABU1$(4, CAITAB1%) = KILO$
        TETABU1$(5, CAITAB1%) = UNID$


        SQLX$ = "SELECT * FROM TANUMSE WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
        Dim TNTMP As ADODB.Recordset
        Set TNTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With TNTMP
          Do While Not .EOF
             NUMESER$ = SAFETEXT$(!NUMEROSERIE)
             TETABU1$(3, CAITAB1%) = NUMESER$
             CAITAB1% = CAITAB1% + 1
          .MoveNext
          Loop
        End With
        TNTMP.Close
        Set TNTMP = Nothing
        '
        CAITAB2% = 0
        CODTABU2$(1) = "NOP-OPER"
        CODTABU2$(2) = "DES-OPER"
        CODTABU2$(3) = "CAN-COMP"
        CACOLTAB2% = 3
        '
        SQLX$ = " SELECT * FROM OPERFAB WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
        SQLX$ = SQLX$ + " ORDER BY NumeOper ASC "
        Dim OPERFAPTMP As ADODB.Recordset
        Set OPERFAPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With OPERFAPTMP
          Do While Not .EOF
                CAITAB2% = CAITAB2% + 1
                NOPE% = XVALO(!NUMEOPER)
                DOPE$ = SAFETEXT(!DESCOPER)
                CAREA# = XVALO(!CANTPROG)
                TETABU2$(1, CAITAB2%) = Str$(NOPE%)
                TETABU2$(2, CAITAB2%) = DOPE$
                TETABU2$(3, CAITAB2%) = FORMATNUM$(CAREA#, "F20.6")
            .MoveNext
          Loop
        End With
        
     Call PRINTDOC("FORSERVI")   ' Orden de Servicio
        '
     End If

     Call COMUNI("Asignación de Armador Realizada")

End Sub
Sub GRABAARMADOR_GABAL(NORFA$, FENTRE%)
        '
        Call ABREORFAB
        '
        'VERIFICO QUE NO TEMGA UN ARMADOR ASOCIADO
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        ARMAD% = XVALO(VALOBADA("ORDEFABR", "ARMADOR", CONDI$, "NOMESS"))
        If ARMAD% > 0 Then
           Call COMUNI("La Orden de Fabricación " & NORFA$ & "\Ya Tiene Asociado al Armador Externo")
           Exit Sub
        End If
        '
        'VALIDO QUE NO ESTE ANULADA
        CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND StatuOrf > 8 "
        OFABRANU& = CantRegistros("ORDEFABR", CONDI$)
        If OFABRANU& > 0 Then
           Call COMUNI("Imposible Asignar Armador a esta Orden de Fabricación " & NORFA$ & "\La misma se Encuentra Anulada")
           Exit Sub
        End If
        '
        'VALIDO QUE NO TENGA CIERRES PARCIALES
        CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND (CANAPRO > 0.05 OR CANRECH > 0.05) "
        OFABRAPRO& = CantRegistros("ORDEFABR", CONDI$)
        If OFABRAPRO& > 0 Then
           Call COMUNI("Imposible Asignar Armador a esta Orden de Fabricación " & NORFA$ & "\Ya Tiene Realizados Cierres Parciales o Totales")
           Exit Sub
        End If
        '
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$))
        '
        'OBJETO PARA ASIGNAR ARMADOR A LA OF
        DEPARMA% = 0: DepOrig% = DEPARMA%
        DEPOCEN% = XVALO(Control$("", "DEPOCEN")) ' DEPOSITO PAÑOL
        Call REPLA(VDEF$, NORFA$, 1, 12)
        Call REPLA(VDEF$, Chr$(DEPOCEN%), 13, 1)
30      OBX$ = OBJPLA$("ORFAARMG", VDEF$)
        If OBX$ = "" Then Exit Sub
        ARMAD% = CVI(Mid$(OBX$, 14, 2))
        'VERIFICA QUE HAYA ELEGIDO UN ARMADOR VALIDO
        If ECOARCH$("ARMADORS", MKI$(ARMAD%)) = "" Then
           Call COMUNI(" Armador no Válido")
           VDEF$ = OBX$
           GoTo 30
        End If
        '
        TX$ = FILTERGETRECORD$("ARMADORS", 1, MKI$(ARMAD%))
        ARMADOR$ = TRIM$(Mid$(TX$, 3, 64)) ' ARMADOR
        proveed% = CVI(Mid$(TX$, 68, 2)) ' NUMERO DE PROVEEDOR ASIGNADO AL ARMADOR
        TIPOARM$ = TRIM$(UCase$(Mid$(TX$, 67, 1))) ' TIPO DE ARMADOR (EXTERNO O INTERNO)
        Deposito% = Asc(Mid$(TX$, 71, 1)) ' DEPOSITO DEL ARMADOR
        '
        'VALIDA QUE SEA ARMADOR EXTERNO
        If TIPOARM$ <> "E" Then
         Call COMUNI("Tipo de Armador \(" & ARMAD% & " " & ARMADOR$ & ") No Valido")
            GoTo 30
        End If
        '
        'VALIDA QUE EL ARMADOR TENGA PROVEEDOR
        If proveed% < 1 Then
           Call COMUNI("Falta Asignar Proveedor al Armador " & ARMAD% & "\" & ARMADOR$)
           VDEF$ = OBX$
           GoTo 30
        End If
        '
        'VALIDA QUE EL ARMADOR TENGA DEPOSITO ASIGNADO
        If ECOARCH$("TADEPOSI", Chr$(Deposito%)) = "" Then
           Call COMUNI("Falta Asignar Deposito al Armador " & ARMAD% & "\" & ARMADOR$)
           VDEF$ = OBX$
           GoTo 30
        End If
        '
        'GRABO EL ARMADOR EN LA OF Y PONGO EN 0 EL SECTOR PRODUCTIVO
        CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
        Call ACTUREGISTRO("ORDEFABR", "ARMADOR", CONDI$, ARMAD%, REGAF&, "NOMESS")
        'Call ACTUREGISTRO("ORDEFABR", "SecciPro", CONDI$, 0, REGAF&, "NOMESS")
        '
        'SE RESERVA EL NUMERO DE ORDEN DE COMPRA GRABANDO UN REGISTRO EN BLANCO
31      NUOCONUE& = PRONUOCO&
        If CantRegistros("OCOMENCA", "NUME_OCOM=" & NUOCONUE&, "NOMESS") > 0 Then GoTo 31
        Call CreaRegistro("OCOMENCA", "NUME_OCOM", NUOCONUE&, "NOMESS")
        '
        'GRABA NUMERO DE ORDEN DE COMPRA EN OF
        Call ACTUREGISTRO("ORDEFABR", "NUME_OCOM", CONDI$, NUOCONUE&, REGAF&)
        '
        'REGISTRA EN EL ANOTADOR DE LA OF LA ASIGNACION DEL ARMADOR
        TTXYZ$ = Space$(76)
        HOII% = HOY(HOS$)
        Call REPLA(TTXYZ$, HOS$, 1, 8)
        Call REPLA(TTXYZ$, "Asignación de Armador Externo", 11, 56)
        Call REPLA(TTXYZ$, USERTER$, 69, 24)
        Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
        '
        TTXYZ$ = Space$(76)
        Call REPLA(TTXYZ$, "(" & ARMAD% & " " & TRIM$(ECOARCH$("ARMADORS", MKI$(ARMAD%))) & ")", 11, 56)
        Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
        Call REPLA(TTXYZ$, "O.Compra Nro.: " & NUOCONUE&, 11, 56)
        Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
        '
        Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
        '
     'GENERACION DE ORDEN DE COMPRA
     CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
     CANTIAFAB = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS")) 'CANTIDAD A FABRICAR / ARMAR
     NORSE& = XVALO(VALOBADA("ORDEFABR", "NumeOrFa", CONDI$, "NOMESS"))
     CI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     OK% = 1
     OBSERX$ = "Orden de Fabricación " + NORFA$
     FECHENT% = FENTRE%
     
     COS1$ = COSPESOS(CI%)
     'PRECOMPRA# = XVALO(InputBox("Precio de Compra", "Valorización de Compras", COS1$))
     PRECOMPRA# = COS1$
     DESCUX = 0
     '
     'GRABO ORDEN DE COMPRA EN OCOMDETA
     Screen.MousePointer = 11
     SQLX$ = "SELECT * FROM OCOMDETA WHERE NUME_OCOM < 1 " 'CONDICION INEXISTENTE
     Dim OCOMTMP As ADODB.Recordset
     Set OCOMTMP = New ADODB.Recordset
25   On Error Resume Next
     OCOMTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0
     '
     With OCOMTMP
        .AddNew
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = NUOCONUE&
        !Femi_Ocom = CVDAT(HO%)
        !Npro_Ocom = proveed%
        !Cod_Inte = CI%
        !Uni_Mas = UNIMED$(CI%)
        !cant_ocom = CANTIAFAB
        !Uni_Com = UNIMED$(CI%)
        !Coef_Unids = 1
          !Fentre_Sol = CVDAT(FECHENT%)
        !CONSIMAT = 1
        !Cant_Rec = 0
        !Cod_Oper = 0
        !Doc_Orig = NORFA$
        !N_ORSERV = NORSE&
        '
        !NuOrd_Item = 1
        !Raso_Prov = RASOCLI$((-1) * proveed%)
        !COD_EXTE = CODEXT$(CI%)
        !Descritem = DESCRIT0$(CI%)
        !Delibre = ""
        !Sect_Ocom = ""
        !Medi_Ocom = ""
        !Valo_Unid = 0
        !Unids_Ocom = CANTIAFAB
        !Desc_Item = XVALO(DESCUX)
        !Tsum_Ocom = 0
        !Stat_Ocom = 0
        !Mone_Emis = 1
        !mone_cos = 1
        !IContab = 0
        !Pre_Unit = PRECOMPRA#
        !prun_neto = PRECOMPRA#
        !inet_item = CANTIAFAB * PRECOMPRA#
        !PrUn_Neto_Pes = PRECOMPRA#
        Monecos$ = "$ "
        VAUNI$ = Monecos$ + Mid$(FORMATNUM$(PRECOMPRA#, "F12.4"), 3)
        !Valo_Unid = VAUNI$
        '
        .Update
     End With
     OCOMTMP.Close
     Set OCOMTMP = Nothing
     '
     'GRABO ORDEN DE COMPRA EN OCOMENCA
     'CAMBIE VARIABLES FECHA Y NORFA POR QUE NO TENIAN INFORMACION
     REFE$ = FECHATEX$(FENTRE%) + " - " + TRIM$(NORFA$) + " - " + RASOCLI$((-1) * proveed%)
     CPAG% = CPAGCLI%((-1) * proveed%)
     '
     SQLX$ = "SELECT * FROM OCOMENCA WHERE NUME_OCOM =" & NUOCONUE& 'CONDICION INEXISTENTE
     Dim OCOMENCTMP As ADODB.Recordset
     Set OCOMENCTMP = New ADODB.Recordset
26   On Error Resume Next
     OCOMENCTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
     End If
     On Error GoTo 0
     
     With OCOMENCTMP
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !Nume_Ocom = NUOCONUE&
        !Refe_Ocom = Left$(REFE$, 64)
        !Femi_Ocom = CVDAT(HO%)
        !Npro_Ocom = proveed%
        !Tsum_Ocom = 2 'FASSON
        !Stat_Ocom = 0 'ABIERTA
        !Cpag_Ocom = CPAG%
        !AtSr_Ocom = ""
        !Obse_Ocom = OBSERX$
        !Raso_Prov = Left$(RASOCLI$((-1) * proveed%), 48)
        !Fgen_Ocom = Now
        !Usua_Gen = USNBR%
        !tcam_ocom = 1
        !IDes_Ocom = 0
        !Embala_Ocom = 0
        !Pdes_Ocom = ""
        !Lentrega = ""
        !mone_cos = 1
        !Mone_Emis = 1
        !Anex_Ocom = ""
        !Iotr1_Ocom = 0
         PVCL% = PIVACLI%(proveed%)
         If PVCL% <> 2 Then
           If PVCL% <> 4 Then
             If PVCL% <> 5 Then
               APLIVA% = 1
             End If
           End If
         End If
         TOTANEDES# = (CANTIAFAB * PRECOMPRA#) * (1 - (XVALO(DESCUX) / 100))
         TOTIVA# = ((1.21) * TOTANEDES#) - TOTANEDES#
         TOTATO# = TOTANEDES# + TOTIVA#
         !TOBRU_OCOM = TOTATO#
         !Iiva_Ocom = TOTIVA#
         !Total_Ocom = TOTATO#
         !Id_Comprob = "OC." + SAFETEXT(NUOCONUE&)
         STATUOC% = 0
         !Stat_Ocom = STATUOC%
        .Update
     End With
     OCOMENCTMP.Close
     Set OCOMENCTMP = Nothing
     '
     'ACTUALIZA LISTA DE SELECCION DE COMPRAS
     CONDI$ = "Nume_Ocom > 0 ORDER BY Nume_Ocom"
     Call CARGALISEL("INDIOCOM", "OCOMENCA", "Nume_Ocom/F7;Refe_OcoM/A44", CONDI$, "DISTINCT")
     Call FRESHECHO("INDIOCOM")
     '
'     'IMPRESION DE ORDEN DE COMPRA
'     If CONTROL$("OCOMPRA", "MODOAPRO") = "POSTERIOR" Then
'          If CONTROL$("OCOMPRA", "IMPRIME") = "ALAPROBA" Then
'             CANCELPRINT% = 1
'          End If
'     End If
     '
     HOII% = HOY(HOS$)
     FORIPRE$ = TRIM$(Control$("", "FORMOCOM"))
     If TRIM$(Control$("", "ORCOMOF")) <> "" Then FORIPRE$ = TRIM$(Control$("", "ORCOMOF"))
     If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
           Exit Sub
         End If
         CANCELPRINT% = 0
         If Control$("OCOMPRA1", "SUPRINT") = "SI" Then
           CANCELPRINT% = 1
         End If
         '
         Screen.MousePointer = 11
         FECHDOC$ = FETEXTO$(HOII%)
         NUMEDOC$ = TRIM$(Str$(NUOCONUE&))
         While Len(NUMEDOC$) < 6
           NUMEDOC$ = "0" + NUMEDOC$
         Wend
         TIPODOC$ = "Nota de Pedido"
         '
         CODPARAM$(1) = "COND-PAG"
         CODPARAM$(2) = "OBSERVA"
         CODPARAM$(3) = "IMP-NETO"
         CODPARAM$(4) = "POR-IVA"
         CODPARAM$(5) = "IMP-IVA"
         CODPARAM$(6) = "IMP-TOTA"
         CODPARAM$(7) = "VENDEDOR"
         CODPARAM$(8) = "LIST-PRE"
         CODPARAM$(9) = "ZONA-VEN"
         CODPARAM$(10) = "FECH-LAN"
         CODPARAM$(11) = "IMPNEDES"
         CODPARAM$(12) = "TRANSPOR"
         CODPARAM$(13) = "POR-DESC"
         CODPARAM$(14) = "IMP-DESC"
         CAPARDOC% = 14
         '
         DOCPARAM$(1) = FORPAG$
         DOCPARAM$(2) = OBSERX$
         DOCPARAM$(3) = Str$(TOTANEDES#)
         DOCPARAM$(4) = Str$(ALIVA)
         DOCPARAM$(5) = Str$(TOTIVA#)
         DOCPARAM$(6) = Str$(TOTATO#)
         DOCPARAM$(7) = Mid$(X1$, 147, 20)
         DOCPARAM$(8) = Mid$(X1$, 67, 40)
         DOCPARAM$(9) = Mid$(X1$, 167, 40)
         DOCPARAM$(10) = Mid$(X1$, 107, 40)
         DOCPARAM$(11) = Str$(TOTANEDES#)
         DOCPARAM$(12) = TRIM$(Mid$(ENCAOCOM$, 163, 30)) ' ATENCION SR.
         DOCPARAM$(13) = TRIM$(PORDESCUENTO$)
         DOCPARAM$(14) = Str$(IDESCUEN#)
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "UNIMED"
         CODTABU1$(4) = "CANTIDAD"
         CODTABU1$(5) = "FECH-VEN"
         CODTABU1$(6) = "REF-MAT1"
         CODTABU1$(7) = "PRE-LIST"
         CODTABU1$(8) = "PORDESCU"
         CODTABU1$(9) = "PRE-NETO"
         CODTABU1$(10) = "NET-ITEM"
         CODTABU1$(11) = "NUBULTOS"
         CODTABU1$(12) = "REF-MAT4"
         CODTABU1$(13) = "ORD-ITEM"
         CACOLTAB1% = 13
         '
         DESTIDOC% = (-1) * proveed%
         CAITAB1% = 0
         CAITCRU% = 0
         '
         JJ& = 0
         
         CI0% = CI%
         CANTU = CANTIAFAB
         PreUnid# = PRECOMPRA#
         PORDE = XVALO(DESCUX)
         FEX = FECHENT%
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 97, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 59, 4)
         TETABU1$(4, CAITAB1%) = Str$(CANTU)
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Str$(CONSIMAT%)  'consigna m.prima
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
         TETABU1$(11, CAITAB1%) = TRIM$(CANBULTOS(CI0%, CDbl(CANTU), "VENTA")) ' CANTIDAD DE BULTOS
            LINVE = ROUND(LINDIVEN(CI%), 0)
            If LINVE >= 1 Then
              EMBA$ = TRIM$(ATRIB1$(CI%))
              If EMBA$ = "" Then EMBA$ = "Bultos"
              TIPUN$ = UNIMED$(CI%)
              PRVASE$ = EMBA$ & " de " & LINVE & " " & TIPUN$
            End If
         TETABU1$(12, CAITAB1%) = TRIM$(PRVASE$)
         TETABU1$(13, CAITAB1%) = TRIM$(Str$(CAITAB%))
         '
         CAITAB1% = CAITAB1% + 1
         For KKII% = 1 To 13: TETABU1$(KKII%, CAITAB1%) = "": Next KKII%
         TETABU1$(2, CAITAB1%) = ""
         '
         Call PRINTDOC("@" + FORIPRE$)  ' Orden de Compra
     End If
     '
     'LIMPIO LAS VARIABLES DE IMPRESION
     Call LIMPIATETAB
     '
     'IMPRIME LISTA DE MATERIALES
     CAIT% = CantRegistros&("FORMULAS", "CodIConj =" & CI%)
     If CAIT% > 0 Then
        FILTX$ = TRIM$(Str$(CI%)) & ";" & TRIM$(Str$(NUOCONUE&)) & ";" & TRIM$(Str$(CANTIAFAB))
        Call STDFORM("LISTAMAT", FILTX$, "PRINT")
     End If
     '
     Call COMUNI("Asignación de Armador Realizada")

End Sub


Sub CARGALACOMPRA_GABAL(NORFA$, NPROV%, PRECOMPRA#, FECHENT%, DESCU, OK%)
    '
    OK% = 0
    '
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CI0% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$))
    CANTU = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
    '
    Call ABRECOMPRAS
    '
    DECRI$ = DESCRIT0$(CI0%)
    UMECO$ = UNIMED$(CI0%)
    PRUNI = 0
    FENTREIT% = HOY(HO$)
    JJ& = 1
    
    REBLA$ = REGBLAN$("OCOMDETA")
    Call SETULTREG("!OCOMDETA", 0)
    '
    ZZ$ = REBLA$
    FENTREIT% = FECHENT%
    Call REPLA(ZZ$, MKI$(CI0%), 9, 2)
    Call REPLA(ZZ$, UNIMED$(CI0%), 11, 4)
    Call REPLA(ZZ$, MKS$(CANTU), 51, 4)
    Call REPLA(ZZ$, UMECO$, 59, 4)
    Call REPLA(ZZ$, MKS$(1), 63, 4)
    Call REPLA(ZZ$, MKI$(FENTREIT%), 95, 2)
    Call REPLA(ZZ$, MKD$(0), 67, 8)
    Call REPLA(ZZ$, 0, 75, 4)
    Call REPLA(ZZ$, MKD$(0), 79, 8)
    Call REPLA(ZZ$, MKD$(0), 87, 8)
    '
    Call GRAREG("!OCOMDETA", ZZ$, 1)
    Call SETULTREG("!OCOMDETA", 1)
    ATRI$ = "Ingreso de Precio para Orden de Compra"
    ATRI$ = ATRI$ + "/POSEDIT(1)=2"
    ATRI$ = ATRI$ + "/POSEDIT(2)=2"
    ATRI$ = ATRI$ + "/POSEDIT(3)=2"
    ATRI$ = ATRI$ + "/POSEDIT(4)=2"
    ATRI$ = ATRI$ + "/POSEDIT(5)=2"
    ATRI$ = ATRI$ + "/POSEDIT(6)=2"
    ATRI$ = ATRI$ + "/POSEDIT(9)=2"
    ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
   
    VTB% = VENTABU%("CARDETCO" & "/" & ATRI$)
    '
    If VTB% < 0 Then OK% = -1
    
    ZZ$ = REGLEIDO$("!OCOMDETA", 1)
    PRECOMPRA# = CVD(Mid$(ZZ$, 67, 8))
    FECHENT = CVI(Mid$(ZZ$, 95, 2)): If FECHENT < HOY(HO$) Then FECHENT = HOY(HO$)
    DESCU = CVS(Mid$(ZZ$, 75, 4)): If DESCU < 0.005 Then DESCU = 0
    
End Sub


Function PRONUOCO&()
   '
   PRONU& = 1
   Call APERBASDAT("COMPRAS")
   
   SQLX$ = "Select max(Nume_ocom) as maximo from OcomEnca WITH(NOLOCK)"
   Dim OEncaTemp As ADODB.Recordset
   Set OEncaTemp = FLEXCONN.Execute(SQLX$)
   If Not IsNull(OEncaTemp!maximo) Then
      PRONU& = 1 + OEncaTemp!maximo
   End If
   OEncaTemp.Close
   Set OEncaTemp = Nothing
   PRONUOCO& = PRONU&

End Function

Function PRONUORSE&()
   '
   PRONU& = 1
   Call APERBASDAT("COMPRAS")
   
   SQLX$ = "Select max(N_OrServ) as maximo from OCOMDETA WITH(NOLOCK)"
   Dim OCOMTMP As ADODB.Recordset
   Set OCOMTMP = FLEXCONN.Execute(SQLX$)
   If Not IsNull(OCOMTMP!maximo) Then
      PRONU& = 1 + OCOMTMP!maximo
   End If
   OCOMTMP.Close
   Set OCOMTMP = Nothing
   PRONUORSE& = PRONU&

End Function




Function HAYDISTO_MODIF%(CIX%, NORFAI&, CACONJU)
     '
     HDS% = 0
     If CHECOMPO% = 0 Then
        CHECOMPO% = (-1)
        If Control$("", "CHECOMPO") = "SI" Then
          If Control$("", "PRECOLAN") <> "SI" Then
            CHECOMPO% = 1
            ' PRESENTA MENSAJE DE ADVERTENCIA SI
            '   VERIFICA COMPONENTES Y NO LOS PRESENTA AL LANZAMIENTO
          End If
        End If
     End If
     '
     If CHECOMPO% <> 1 Then
        HAYDISTO_MODIF% = 1
        Exit Function
     End If
     '
     If CIX% < 1 Or CIX% > NUMAS% Then
        HAYDISTO_MODIF% = 1
        Exit Function
     End If
     '
     COD$ = CODEXT$(CIX%)
     CODCON$ = COD$
     Call LEEEXPLOMUL(COD$, 1, 1)
     For II% = 1 To CAIT%
        CI2% = CIJ%(II%)
        CAN = CACONJU * USOI(II%)
        STDIS = STOCKDIS(CI2%)
        STROF = STORESOF(CI2%, NORFAI&)
        If CAN > STDIS + STROF Then
           MENSA$ = "El Stock Disponible (" + TRIM$(CODEXT$(CI2%)) + ") es Insuficiente\para Lanzar Esta Orden de Fabricación."
           If COMALTER%(MENSA$ + "\\Lanzar Igual\Cancelar") = 1 Then
              HDS% = 1
           End If
           GoTo 99
        End If
     Next II%
     HDS% = 1
     '
99   HAYDISTO_MODIF% = HDS%
     '
End Function
'
Sub GENRESER_MODIF(NROFLLA$)
     '
10   NROOF$ = NROFLLA$
     HO% = HOY(HOS$)
     '
     If PRECOLAN% = 0 Then
        PRECOLAN% = (-1)
        If Control$("", "PRECOLAN") = "SI" Then
           PRECOLAN% = 1
        End If
     End If
     '
     If TRIM$(NROOF$) <> "" Then
        NORFA$ = AJUSTI$(NROOF$, 12)
        If TRIM$(ECOARCH$("ECORDEP", NORFA$)) <> "" Then
            GoTo 120
        End If
     End If
     NROOF$ = ""
     '
100  Call ACTECOR(1)     ' parametro 1 para seleccionar solo llamzadas y suspendidas
     Screen.MousePointer = 1
     NROOF$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
110  Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
120  Call ABREORFAB
     CONDI$ = "IdSubOr = '" & NROOF$ & "'"
     CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
     CANPROYE# = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$, "NOMESS"))
     CANAPROB# = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$, "NOMESS"))
     CAPEN# = CANPROYE# - CANAPROB#
     If CAPEN# < 0 Then CAPEN# = 0
     '
     ' MODIFICADA LA RE-GENERACION DE RESERVAS POR O/F PARA REALIZAR EL CÁLCULO POR
     ' EL TOTAL PROYECTADO DE LA O/F Y LUEGO DESCONTAR LO CONSUMIDO DE ESA MISMA
     ' ORDEN DE FABRIVCACION - EPB - 16 JUNIO 2009
          '
1410 If CAPEN# < 0.1 Then
       Call MENSERR(24, "No hay Pendiente para la O/F '" + TRIM$(NORFA$) + "'.")
     End If
     '
     CICONJ% = CIXI%
     CODICONJ$ = CODEXT$(CIXI%)
     NORFABRI$ = NORFA$
     CACONJU# = CANPROYE#    ' CAPEN#
     VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
     '
1412 Call BLOQARCH("RESERVOF")
     Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 3)
     If VACONTROL% < 0 Then
        GoTo 1499
     End If
     '
     If RECLEN%("MATEROF") > 32 Then
        Call ASIGNALOTE_MATEROF_MODIF(NORFA$, OK%)
        If OK% < 0 Then GoTo 1412 ' SI NO SE ASIGNARON LOTES A COMPONENTES
     End If
     
     Call ABRESERVA
     Call BORRARESERVA(NORFA$)
     '
     'SE BORRA EL DELETE POR QUE SE REALIZA EN LA RUTINA BORRARESERVA
'     SQLX$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
'     FLEXCONN.Execute SQLX$
     '
1415 Dim RESERVVVA As ADODB.Recordset
     Set RESERVVVA = New ADODB.Recordset
     SQLX$ = " SELECT * FROM Reserva where IdSubOr = '" & NORFA$ & "' "
     RESERVVVA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
     NORIT% = 0
     '
     For U& = 1 To ULTREG&("!MATEROF")
        ZZ$ = REGLEIDO$("!MATEROF", U&)
        CIIX% = CVI(Left$(ZZ$, 2))
        CANTX# = CVS(Mid$(ZZ$, 5, 4))
        USOII = 0: If Abs(CANPROYE#) >= 0.05 Then USOII = CANTX# / CANPROYE#
        LOTE$ = TRIM$(Mid$(ZZ$, 33, 12))
        '
        If Abs(CANTX#) >= 0.05 Then
          With RESERVVVA
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !Cod_Inte = CIIX%
            !COD_EXTE = CODEXT$(CIIX%)
            !DESCRIP = DESCRIT0$(CIIX%)
            !IDSUBOR = NORFA$
            !FechRes = Now
            !CantRes = CANTX#
               CANCONSUM = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", "Cod_Inte = " & CIIX% & " AND CodiMovi = 'CF' AND DoPriLar = '" & NORFA$ & "'", "NOMESS"))
            !CantCons = CANCONSUM
               SALDOEN = CANTX# - CANCONSUM
               If SALDOEN < 0 Then SALDOEN = 0
            !Saldo_Entre = SALDOEN
            !IdenLote = LOTE$
            !Deposito = 0
            !UsoUnit = USOII
               NORIT% = NORIT% + 1
            !NuOrItem = NORIT%
            !NUME_SOLINT = 0
            .Update
          End With
        End If
     Next U&
     RESERVVVA.Close
     Set RESERVVVA = Nothing
     '
1499 Call LIBARCH("RESERVOF")
     If TRIM$(NROFLLA$) = "" Then
        GoTo 100
     End If
     '
End Sub
'
Sub BORRARESERVA(NORFA$)
      'SUB RECORRE EL RESERVA Y LEE Y ACTUALIZA EL LOTE
      Call ABRESERVA
      SQLX$ = " SELECT * FROM Reserva "
      SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
      Dim RESERVMTP As ADODB.Recordset
      Set RESERVMTP = New ADODB.Recordset
      RESERVMTP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      With RESERVMTP
        Do While Not .EOF
            CI% = XVALO(!Cod_Inte)
            If ESTRAZABLE(CI%) < 1 Then GoTo 30
            !CantRes = XVALO(!CantCons)
            LOTEX$ = TRIM$(SAFETEXT$(Left$(!IdenLote, 16)))
            .Update
            If LOTEX$ <> "" Then
               Call ACSALOTE(CI%, LOTEX$, "NOMESS")
            End If
30          .MoveNext
        Loop
      End With
    RESERVMTP.Close
    Set RESERVMTP = Nothing
    '
    'BORRRA LA RESEVA DE ESTA ORDEN DE FABRICACION
    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NORFA$ & "'", REGAF&, "NOMESS")
    
End Sub
Sub ASIGNALOTE_RESERVA(NORFA$, OK%)

      '
'SE CAMBIO LA CNTIDD DE DECIMLES QUE ERA FIJA DE 1  X  & CNTDC$(CI%)
'ESTO EN ADMIORF07 (MODUILO)
    OK% = 0
    TRAZCOMPO.List2.Clear
    For YX& = 1 To ULTREG&("!RESERVA")
       YY$ = REGLEIDO$("!RESERVA", YX&)
       CI% = CVI(Mid$(YY$, 1, 2))
       USOII = CVS(Mid$(YY$, 75, 4))
       If ESTRAZABLE%(CI%) = 1 Then
         CAN# = ROUND(CVD(Mid$(YY$, 19, 8)), 3)
         If CAN# > FRACMIN(CI%) Then
            TX$ = ""
            Call REPLA(TX$, CODEXT$(CI%), 1, 16)
            Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN#, "F10." & CNTDC$(CI%)), 50, 11)
            Call REPLA(TX$, FORMATNUM$(USOII, "F12." & CNTDC$(CI%)), 85, 12)
            TRAZCOMPO.List2.AddItem TX$
         End If
       End If
       '
   Next YX&
   
   'Presentar formulario para asignar lotes
   If TRAZCOMPO.List2.ListCount > 0 Then
     TRAZCOMPO.Show 1
     If TRAZCOMPO.List2.ListCount < 1 Then
          TRAZCOMPO.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         j& = 0
         For YX& = 1 To ULTREG&("!RESERVA")
          YY$ = REGLEIDO$("!RESERVA", YX&)
          CI% = CVI(Mid$(YY$, 1, 2))
          CAN# = CVD(Mid$(YY$, 19, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If ESTRAZABLE%(CI%) < 1 Then
                 j& = j& + 1
                 Call GRAREG("!RESERVA", YY$, j&)
              End If
            End If
          End If
        Next YX&
        FE% = HOY(HO$)
        '
        For I& = 1 To TRAZCOMPO.List2.ListCount
             'CODIGO DE LIST DE TRAZCOMPO
              CI% = CODINT%(TRIM$(Left$(TRAZCOMPO.List2.List(I& - 1), 16)))
              LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(I& - 1), 62, 16))
              CANTI# = XVALO(Mid$(TRAZCOMPO.List2.List(I& - 1), 50, 11))
              USOII = XVALO(Mid$(TRAZCOMPO.List2.List(I& - 1), 85, 12))
              ESTRAZ% = ESTRAZABLE%(CI%)
              '
              YY$ = REGBLAN("!RESERVA")
              Call REPLA(YY$, MKI$(CI%), 1, 2)
              Call REPLA(YY$, NORFA$, 3, 12)
              Call REPLA(YY$, MKI$(CI%), 15, 2)
              Call REPLA(YY$, MKI$(FE%), 17, 2)
              Call REPLA(YY$, MKD$(CANTI#), 19, 8)
              Call REPLA(YY$, MKD$(CANTI#), 35, 8)
              Call REPLA(YY$, LOTE$, 43, 16)
              Call REPLA(YY$, MKS$(ESTRAZ%), 61, 4)
              Call REPLA(YY$, MKS$(USOII), 75, 4)
              '
              j& = j& + 1
              Call GRAREG("!RESERVA", YY$, j&)
        Next I&
        '
        Call SETULTREG("!RESERVA", j&)
     End If
   End If
   '
   TRAZCOMPO.List2.Clear
 
End Sub

Sub ASIGNALOTE_MATEROF_MODIF(NORFA$, OK%)
      '
    OK% = 0
    For YX& = 1 To ULTREG&("!MATEROF")
       YY$ = REGLEIDO$("!MATEROF", YX&)
       CI% = CVI(Left$(YY$, 2))
       If ESTRAZABLE%(CI%) = 1 Then
         CAN# = CVS(Mid$(YY$, 5, 4))
         If CAN# > 0.005 Then
            TX$ = ""
            Call REPLA(TX$, CODEXT$(CI%), 1, 16)
            Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN#, "F10.1"), 50, 11)
            TRAZCOMPO.List2.AddItem TX$
         End If
       End If
       '
   Next YX&
   
   'Presentar formulario para asignar lotes
   If TRAZCOMPO.List2.ListCount > 0 Then
     TRAZCOMPO.Show 1
     If TRAZCOMPO.List2.ListCount < 1 Then
          TRAZCOMPO.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         j& = 0
         For YX& = 1 To ULTREG&("!MATEROF")
          YY$ = REGLEIDO$("!MATEROF", YX&)
          CI% = CVI(Mid$(YY$, 1, 2))
          CAN# = CVD(Mid$(YY$, 19, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If ESTRAZABLE%(CI%) < 1 Then
                 j& = j& + 1
                 Call GRAREG("!MATEROF", YY$, j&)
              End If
            End If
          End If
        Next YX&
        '
        For I& = 1 To TRAZCOMPO.List2.ListCount
              'CODIGO DE LIST DE TRAZCOMPO
              CI% = CODINT%(TRIM$(Left$(TRAZCOMPO.List2.List(I& - 1), 16)))
              'GRABO EN EL MATEROF EL NUMERO DE LOTE
              LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.List(I& - 1), 62, 16))
              CANTI# = XVALO(Mid$(TRAZCOMPO.List2.List(I& - 1), 50, 11))
              '
              YY$ = REGBLAN("!MATEROF")
              Call REPLA(YY$, MKI$(CI%), 1, 2)
              Call REPLA(YY$, MKS$(CANTI#), 5, 4)
              Call REPLA(YY$, LOTE$, 33, 16)
              '
              j& = j& + 1
              Call GRAREG("!MATEROF", YY$, j&)
        Next I&
        '
        Call SETULTREG("!MATEROF", j&)

     End If
   End If
   '
   TRAZCOMPO.List2.Clear

End Sub

Sub CONSUMECOMPO(CIXI%, NORFAX$, CANCONSUM, FECONSUM%, DEPX%)
    '
    SALCONSU = CANCONSUM
    '
    ' BARRIDA SOBRE LOS LOTES RESERVADOS DE ESE COMPONENTE
    CONDI2$ = " IdSubOr = '" & NORFAX$ & "' AND COD_INTE = " & CIXI%
    STRSQLX$ = "Select * from RESERVA WITH(NOLOCK) WHERE " + CONDI2$
'    Dim RRES As ADODB.Recordset
'    Set RRES = New ADODB.Recordset
'25  On Error Resume Next
'    RRES.Open (STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 25
'    End If
'    On Error GoTo 0
    Set RRES = READSET(STRSQLX$)
    '
    U& = 0
    With RRES
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         QORRES = XVALO(!CantRes)
         QYACON = XVALO(!CantCons)
         SALRES = QORRES - QYACON
         QCON = SALRES: If QCON > SALCONSU Then QCON = SALCONSU
         '
         If QCON >= FRACMIN(CIXI%) Then
            IDLO$ = SAFETEXT$(!IdenLote)
            '
            ' REALIZAR EL MOVIMIENTO DE STOCK
            Call MOVICONSUFA(FECONSUM%, CIXI%, IDLO$, NORFAX$, QCON, DEPX%)
            '
            ' ACTUALIZAR CONSUMIDO EN LA RESERVA
            Call ACUCONRES(CIXI%, IDLO$, NORFAX$)
            '
            SALCONSU = SALCONSU - QCON
            If SALCONSU < FRACMIN(CIXI%) Then Exit Sub
         End If
      .MoveNext
      Loop
    End With
    RRES.Close
    Set RRES = Nothing
    '
    If SALCONSU >= FRACMIN(CIXI%) / 2 Then
       ' HACER UN MOVIMIENTO ADICIONAL DE STOCK DE MATERIAL NO RESERVADO
       CONDIX$ = " IdSubOr = '" & NORFAX$ & "' AND Cod_Inte = " & CIXI% & ""
       IDLO$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
' IR A BUSCAR ALGUN LOTE CON SALDO NO RESERVADO
       Call MOVICONSUFA(FECONSUM%, CIXI%, IDLO$, NORFAX$, SALCONSU, DEPX%)
    End If
    '
End Sub

Sub MOVICONSUFA(FECONSUM%, CIXI%, IDLO$, NORFAX$, QCON, DEPX%)
    '
    CMOV$ = "CF"
    DOPRI$ = NORFAX$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
    EXISTE_REGISTRO% = 0
    '
    ' BUSCAR SI SE PUEDE ACUMULAR A UN MOVIMIENTO YA REGISTRADO
    CONDI$ = "CODIMOVI = '" & CMOV$ & "' "
    CONDI$ = CONDI$ & " AND COD_INTE = " & CIXI%
    CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "' "
    CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
    '
    SQLX$ = "SELECT CANTSALID, FECHMOV FROM MOVISTO WHERE " & CONDI$
    Dim MOVICONSU As ADODB.Recordset
    Set MOVICONSU = New ADODB.Recordset
    MOVICONSU.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    On Error Resume Next
    With MOVICONSU
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           !Cantsalid = XVALO(!Cantsalid) + QCON
           !Fechmov = CVDAT(FECONSUM%)
           .Update
           EXISTE_REGISTRO% = 1
        End If
        On Error GoTo 0
    End With
    MOVICONSU.Close
    Set MOVICONSU = Nothing
    
    If EXISTE_REGISTRO% = 0 Then
         CODICONJ$ = Label7
         REFE$ = "CONSUMO " + CODICONJ$
         Call MOVISTOK(FECONSUM%, CIXI%, IDLO$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, (-1) * QCON)
    End If
'    If CantRegistros("MOVISTO", CONDI$) = 1 Then
'         ' ACUMULA EN REGISTRO ANTERIOR
'         TOCONSU = QCON + XVALO(VALOBADA("MOVISTO", "CANTSALID", CONDI$, "NOMESS"))
'         Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI$, TOCONSU, RAFE&, "NOMESS")
'         Call ACTUREGISTRO("MOVISTO", "FECHMOV", CONDI$, CVDAT(FECONSUM%), RAFE&, "NOMESS")
'       Else
'         ' CREA UN NUEVO REGISTRO
'         CODICONJ$ = Label7
'         REFE$ = "CONSUMO " + CODICONJ$
'         Call MOVISTOK(FECONSUM%, CIXI%, IDLO$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, (-1) * QCON)
'    End If
    '
End Sub

Sub ACUCONRES(CIXI%, IDLO$, NORFAX$)
    '
    CMOV$ = "CF"
    DOPRI$ = NORFAX$
    DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
    '
    ' BUSCAR SI SE PUEDE ACUMULAR A UN MOVIMIENTO YA REGISTRADO
    CONDI$ = "CODIMOVI = '" & CMOV$ & "' "
    CONDI$ = CONDI$ & " AND COD_INTE = " & CIXI%
    CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "' "
    CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
    CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
    '
    ACMOV = VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS")
    '
    CONDIX$ = " IdSubOr = '" & NORFAX$ & "' AND Cod_Inte = " & CIXI% & " AND IDENLOTE = '" & IDLO$ & "'"
    SQLX$ = "SELECT CANTCONS, SALDO_ENTRE, CANTRES FROM RESERVA WHERE " & CONDIX$
    Dim ReserTemp As ADODB.Recordset
    Set ReserTemp = New ADODB.Recordset
    ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    On Error Resume Next
    With ReserTemp
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           TORES = XVALO(!CantRes)
           TCONSU = ACMOV: If TCONSU > TORES Then TCONSU = TORES
           SALENT = TORES - TCONSU
           !CantCons = TCONSU
           !Saldo_Entre = SALENT
           .Update
        End If
        On Error GoTo 0
    End With
    ReserTemp.Close
    Set ReserTemp = Nothing
    
    
'    QTR& = CantRegistros("RESERVA", CONDIX$, "NOMESS")
'    If QTR& = 1 Then
'       TORES = XVALO(VALOBADA("RESERVA", "CANTRES", CONDIX$, "NOMESS"))
'       TCONSU = ACMOV: If TCONSU > TORES Then TCONSU = TORES
'       SALENT = TORES - TCONSU
'       'SE COMENTA ESTA LINEA POR QUE ESTABA ACTUALIZANDO LA CANTIDAD RESERVADA EN VEZ DE LA CONSUMIDA
''       Call ACTUREGISTRO("RESERVA", "CANTRES", CONDIX$, TCONSU, RAFE&, "NOMESS")
'       Call ACTUREGISTRO("RESERVA", "CANTCONS", CONDIX$, TCONSU, RAFE&, "NOMESS")
'       Call ACTUREGISTRO("RESERVA", "SALDO_ENTRE", CONDIX$, SALENT, RAFE&, "NOMESS")
'    End If
    '
End Sub

Sub APCIEOF()
   '
   Dim USUASEC(10) As Single
   '
10 FIN& = ULTREG&("ORFAPEN")
   Call SETULTREG("ORFABAP", 0)
   Call TRALOCAL("ORFABAP", "")
   '
   ' TOMA EL NUMERO DE USUARIO ACTIVO Y VERIFICA SI EL MISMO ESTÁ EN ALGUN SECTOR
   USUA% = USNBR% Mod 100
   If USUA% < 1 Then USUA% = 1
   REGI_USUASEC$ = RECFILT$("USUASEC", 1, MKI$(USUA%))
   If REGI_USUASEC$ = "" Then
        Call COMUNI("Imposible Continuar.\El Usuario '" & USERNAME & "' No Posee Sector Asociado")
        Exit Sub
   End If
   JJ% = 0
   ' CARGA LOS SECTORES EN LOS QUE TRABAJA SOBRE EL VECTOR USUASEC
   For IIK% = 1 To Len(REGI_USUASEC$) Step 4
        REGI& = CVS(Mid$(REGI_USUASEC, IIK%, 4))
        XX$ = REGLEIDO$("USUASEC", REGI&)
        JJ% = JJ% + 1
        USUASEC(JJ%) = CVS(Mid$(XX$, 3, 4))
   Next IIK%
   '
   '*** CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO ORFAPEN CON LA OF PENDIENTES DE APROBACION
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("ORFAPEN", U&)
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     CMOV$ = Mid$(TXTX$, 21, 2)
     If CIII% > 0 And CMOV$ = "OF" Then
       DOPRICO$ = ""
       DOPRI$ = TRIM$(Mid$(TXTX$, 23, 10))
       '
       ' VERIFICA SI EL USUARIO ACTUAL PERTENECE AL SECTOR ASIGNADO A LA OF PENDIENTE DE APROBACION
       ENCONTRO% = 0
       'DEPOEN% = XVALO(VALOBADA("MASTER", "DEPOENTRA", "CODINT=" & CIII%))
       DEPOEN% = Asc(Mid$(TXTX$, 83, 1))
       For IIJ% = 1 To JJ%
            REGI_SECTOR = FILTERGETRECORD$("SECPROD", 1, MKS$(USUASEC(IIJ%)))
            DEPOSEC% = Asc(Mid$(REGI_SECTOR, 81, 1))
            If DEPOSEC% = DEPOEN% Then
                ENCONTRO% = 1
                Exit For
            End If
       Next IIJ%
       If ENCONTRO% = 0 And USUA% <> 1 Then GoTo 20
       ' FIN VERIFICACION
       '
       DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
       CantIngre = CVD(Mid$(TXTX$, 88, 8))
       CANTSALI = CVD(Mid$(TXTX$, 96, 8))
       AJUST = CantIngre - CANTSALI
       CLA1$ = TRIM$(Mid$(TXTX$, 112, 10))
       Call REPLA(TTXX$, TRIM$(CLA1$), 1, 16)
       Call REPLA(TTXX$, DOPRI$, 18, 16)
       Call REPLA(TTXX$, CODEXT$(CIII%), 36, 16)
       Call REPLA(TTXX$, AJUSTI$(DESCRIT0$(CIII%), 24), 52, 24)
       Call REPLA(TTXX$, FORMATNUM$(AJUST, "F10.1"), 78, 10)
       Call REGAPP("!ORFABAP", TTXX$)
     End If
20 Next U&
   Call CIERRARCH("!ORFABAP")
   'Call TRACENTRAL("ORFABAP", "")
   If ULTREG&("!ORFABAP") < 1 Then Call COMUNI("No Hay O-Fabricación Para Autorizar"): Exit Sub
   Call FRESHECHO("!ORFABAP")
   '
   '*** PRESENTO LISTA DE SELECCION
50 Call SELECHO("!ORFABAP/Ordenes de Fabricación Pendientes de Aprobación")
   CLA0$ = TRIM$(VALACT1$("!ORFABAP"))
   If CLA0$ = "" Then Exit Sub
   DOPRI$ = TRIM$(Mid$(VALACT2$("!ORFABAP"), 1, 16))
   '
   '*** BUSCA REGISTRO DE LA OF ELEGIDA EN ORFAPEN PARA GRABAR EL MOVIMIENTO DE STOCK DE ENTRADA
   For III& = 1 To ULTREG&("ORFAPEN")
     TXTX$ = REGLEIDO$("ORFAPEN", III&)
     CMOV$ = Mid$(TXTX$, 21, 2)
     CLA1$ = TRIM$(Mid$(TXTX$, 112, 10))
     If CLA1$ = CLA0$ And CMOV$ = "OF" Then
        REG& = III&
        Exit For
     End If
   Next III&
   If REG& > 0 Then
         '*** JUNTA LA INFORMACION PARA REGISTRAR LA ENTRADA DEL PRODUCTO FABRICADO AL STOCK
         TX1$ = REGLEIDO$("ORFAPEN", REG&)
         '
         DP2% = CVI(Mid$(TX1$, 127, 2))
         If ECOARCH$("TADEPOSI", Chr$(DP2%)) = "" Then
              Call COMUNI("Imposible Continuar.\Falta Definir Deposito de Consumo.")
              Call SELECHO("TADEPOSI")
              DP2% = XVALO(VALACT1$("TADEPOSI"))
              If ECOARCH$("TADEPOSI", Chr$(DP2%)) = "" Then GoTo 50
         End If
         REFE$ = TRIM$(Mid$(TX1$, 43, 30))
         FF% = CVI(Mid$(TX1$, 1, 2))  'TOMA LA FECHA EN QUE SE GENERO EL CIERRE, NO LA DE LA APROBACION
         CIII% = CVI(Mid$(TX1$, 3, 2))
         LTX$ = TRIM$(Mid$(TX1$, 5, 16))
         CMOV$ = TRIM$(Mid$(TX1$, 21, 2))
         DOPRICO$ = TRIM$(Mid$(TX1$, 23, 10))
            While Len(DOPRICO$) < 11: DOPRICO$ = Mid$(DOPRICO$, 1, 3) & "0" & Mid$(DOPRICO$, 4): Wend
         DOSECO$ = TRIM$(Mid$(TX1$, 33, 10))
            While Len(DOSECO$) < 11: DOSECO$ = Mid$(DOSECO$, 1, 3) & "0" & Mid$(DOSECO$, 4): Wend
         DEPO% = Asc(Mid$(TX1$, 83, 1))
         CANTIENT1 = CVD(Mid$(TX1$, 88, 8))
         CANTISAL1 = CVD(Mid$(TX1$, 96, 8))
         CANTI = CANTIENT1 - CANTISAL1
         NUORIT% = 1
         MONEII% = 1
         CODICONJ$ = CODEXT$(CIII%)
         '
         '*** CONSULTA AL USUARIO SI ACEPTA LA CANTIDAD INFORMADA QUE SE FABRICO EN EL CIERRE
         CANTIAPRO = CANTI
         CANMODIF = XVALO(InputBox("Aprobación para:" & vbCrLf & CODEXT$(CIII%) & " - " & DESCRIT0$(CIII%), "Aprobación " & DOPRICO$, Str$(CANTIAPRO)))
         '
         If CANMODIF > 0 Then
            OPXX% = COMALTER%("¿Confirma Cantidades Aprobadas Para '" & CODEXT$(CIII%) & "' Por " _
            & TRIM$(FORMATNUM$(CANMODIF, "F10.1")) & "?\\No, Cancelar\Si, Continuar")
            If OPXX% = 2 Then
                 Call BLOQARCH("ORFAPEN")
                 Screen.MousePointer = 11
                 '*** GENERA EL MOVIMIENTO DE STOCK TANTO DE LA OF, COMO EL CONSUMO DE OF O EL SCRAP
                 CANTI = CANMODIF
                 Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
                 ARGBU$ = Mid$(TX1$, 23, 10) ' OF BUSCADA
                 Call REPLA(TX1$, "ANULARORFA", 23, 10)
                 Call GRAREG("ORFAPEN", TX1$, REG&)
                 '
                 '*** JUNTA LA INFORMACION PARA REGISTRAR EL CONSUMO DE LOS COMPONENTES DEL PROD. FABRICADO
                 NORFABRI$ = DOPRICO$
                 CACONJU# = CANTI
                 For II& = 1 To ULTREG&("ORFAPEN")
                    TTX1$ = REGLEIDO$("ORFAPEN", II&)
                    '
                    ORFA$ = TRIM$(Mid$(TTX1$, 23, 10))
                    CMOV$ = TRIM$(Mid$(TTX1$, 21, 2))
                    CLA1$ = TRIM$(Mid$(TTX1$, 112, 10))
                    If ORFA$ <> ARGBU$ Then GoTo 146                    'VALIDA QUE SEA LA OF
                    If CMOV$ = "OF" Then GoTo 146                       'VALIDA QUE EL TIPO DE MOV. SEA OF
                    'If CLA1$ <> CLA0$ And CLA1$ <> "" Then GoTo 146     'VERIFICA QUE SEA LA CLAVE A CERRAR
                    '
                    DP2% = CVI(Mid$(TTX1$, 127, 2))
                    REFE$ = TRIM$(Mid$(TTX1$, 43, 30))
                    FF% = CVI(Mid$(TTX1$, 1, 2))  'TOMA LA FECHA EN QUE SE GENERO EL CIERRE, NO LA DE LA APROBACION
                    CI2% = CVI(Mid$(TTX1$, 3, 2))
                    LTX$ = TRIM$(Mid$(TTX1$, 5, 16))
                    CMOV$ = TRIM$(Mid$(TTX1$, 21, 2))
                    DOPRICO$ = TRIM$(Mid$(TTX1$, 23, 10))
                       While Len(DOPRICO$) < 11: DOPRICO$ = Mid$(DOPRICO$, 1, 3) & "0" & Mid$(DOPRICO$, 4): Wend
                    DOSECO$ = TRIM$(Mid$(TX1$, 33, 10))
                       While Len(DOSECO$) < 11: DOSECO$ = Mid$(DOSECO$, 1, 3) & "0" & Mid$(DOSECO$, 4): Wend
                    DEPO% = Asc(Mid$(TTX1$, 83, 1))
                    CANTIENT1 = CVD(Mid$(TTX1$, 88, 8))
                    CANTISAL1 = CVD(Mid$(TTX1$, 96, 8))
                    CAN = CANTIENT1 - CANTISAL1
                    MONEII% = 1
                    '
                    NORFA$ = NORFABRI$
                    NUORIT% = 0
                    FECHA% = FF%
                    '
                    If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
                      CAN = CAN / 1000
                    End If
                    '
                    USOUNI# = CVD(Mid$(TTX1$, 9, 8))
                    RERES& = CVS(Mid$(TTX1$, 17, 4))
                    If CI2% > 0 Then
                     If CI2% <= NUMAS% Then
                       If Abs(CAN) >= 0.05 Then
                         CONDIX$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2% & ""
                         LOTE$ = VALOBADA("Reserva", "IdenLote", CONDIX$, "NOMESS")
                         CMOV$ = "CF"
                         DOPRI$ = NORFA$
                         DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
                         REFE$ = "CONSUMO " + CODICONJ$
                         If CAN > 0 Then
                            REFE$ = "RECUPERO " + CODICONJ$
                         End If
                         '*** GENERA EL MOVIMIENTO DE STOCK DE LOS COMPONENTES
                         DEPX% = DEPO%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
144                      NUORIT% = NUORIT% + 1
                         Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                         '
                         '*** ACTUALIZA LA RESERVA DE MATERIALES
                         SALCON# = (-1) * CAN
                         SQLX$ = "SELECT * FROM Reserva "
                         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
                         SQLX$ = SQLX$ + " AND Cod_Inte = " & CI2% & " "
                         Dim RESETEMP As ADODB.Recordset
                         Set RESETEMP = New ADODB.Recordset
                         RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                         On Error Resume Next
                         With RESETEMP
                           .MoveFirst
                           If Err < 1 Then
                             On Error GoTo 0
                             Do While Not .EOF
                                CARESE# = XVALO(!CantRes)
                                CACOANT# = XVALO(!CantCons)
                                If CARESE# - CACOANT# >= SALCON# Then
                                     !CantCons = CACOANT# + SALCON#
                                     SALCON# = 0
                                   Else
                                     SALCON# = SALCON# - (CARESE# - CACOANT#)
                                     !CantCons = XVALO(!CantRes)
                                End If
                                .Update
                                .MoveNext
                             Loop
                           End If
                           On Error GoTo 0
                         End With
                         Set RESETEMP = Nothing
                         Call REPLA(TTX1$, "ANULARORFA", 23, 10)
                         Call GRAREG("ORFAPEN", TTX1$, II&)
                       End If
                     End If
                    End If
146              Next II&
                 '
                 'BORRO EL REGISTRO DEL ARCHIVO DE PENDIENTES DE APROBACION
                 REG& = 0
                 For II& = 1 To ULTREG&("ORFAPEN")
                    XX$ = REGLEIDO$("ORFAPEN", II&)
                    MARCA$ = Mid$(XX$, 23, 10)
                    If MARCA$ <> "ANULARORFA" Then
                        REG& = REG& + 1
                        Call GRAREG("ORFAPEN", XX$, REG&)
                    Else
                    A = B
                    End If
                 Next II&
                 Call SETULTREG("ORFAPEN", REG&)
                 Call CIERRARCH("ORFAPEN")
                 '
                 CIERRA% = 0
                 ' **** ACTUALIZA LA TABLA DE ORDENES DE FABRICACION
                 SQLX$ = " SELECT * FROM  ORDEFABR "
                 SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
                 Dim ABREORFFFAB As ADODB.Recordset
                 Set ABREORFFFAB = New ADODB.Recordset
                 ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                 On Error Resume Next
                 With ABREORFFFAB
                   .MoveFirst
                   If Err < 1 Then
                      On Error GoTo 0
                     !CanApro = XVALO(!CanApro) + CANTI
                     !CanRech = XVALO(!CanRech) + SCRAP
                     If (!CanApro - !CanProy) > 0.05 Then
                       !Statuorf = 3         ' MARCA COMO CERRADA
                       CIERRA% = 1
                     End If
                     !FechCierr = CVDAT(FECHACIOF%)
                     .Update
                   End If
                   On Error GoTo 0
                 End With
                 Set ABREORFFFAB = Nothing
                 '
                 If CIERRA% = 1 Then
                    Call ABRESERVA
                    SQLX$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
                    FLEXCONN.Execute SQLX$
                 End If
                 '
                 Call LIBARCH("ORFAPEN")
                 '
                 Call COMUNI("Aprobación Procesada")
            End If
         End If
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub

Sub IMPRE_ETIQUETA(Optional MODO%)

10  Call SELECHO("ECORDEF/Indice de Ordenes de Fabricación (Pendientes Cerradas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then Exit Sub
    
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CANTI = XVALO(VALOBADA("ORDEFABR", "CanApro", CONDI$, "NOMESS"))
    CI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$, "NOMESS"))
    '
    CONDI$ = "IDENLOTE ='" & NORFA$ & "' AND COD_INTE = " & CI%
    '
    ' EtiqxCaja# = XVALO(VALOBADA("LOTINGRE", "EtiqxCaja", CONDI$))
    EtiqxRollo# = XVALO(VALOBADA("LOTINGRE", "EtiqxRollo", CONDI$, "NOMESS"))
    RolloxCaja# = XVALO(VALOBADA("LOTINGRE", "RolxCaja", CONDI$, "NOMESS"))
    '
    If CANTI < 1 Then
       Call COMUNI("Imposible Realizar Reimpresión\No Hay Cantidades Cerradas Para Esta Orden")
       GoTo 10
    End If
    Call REPLA(VDEF$, NORFA$, 1, 12)
    Call REPLA(VDEF$, MKI(CI%), 13, 2)
    Call REPLA(VDEF$, MKS(CANTI), 15, 4)
    '
    ' CARGA DATOS EN EL OBJETO
    Call REPLA(VDEF$, MKS(EtiqxRollo#), 19, 4)
    Call REPLA(VDEF$, MKS(RolloxCaja#), 23, 4)
    Call REPLA(VDEF$, MKS(EtiqxCaja#), 27, 4)
    '
    OBJETIQUETA$ = "ETIQORFA"
    If MODO% = 1 Then OBJETIQUETA$ = "ETIQORF1"
30  OBX$ = OBJPLA$(OBJETIQUETA$, VDEF$) ' ABRE EL OBJETO
    '
    If OBX$ = "" Then GoTo 10
    '
    ' OBTIENE DATOS DEL OBJETO
    EtiqxRollo# = CVS(Mid$(OBX$, 19, 4))
    RolloxCaja# = CVS(Mid$(OBX$, 23, 4))
    CANTI = CVS(Mid$(OBX$, 15, 4))
    '
    If EtiqxRollo# < 1 Then
      Call COMUNI("Cantidad de Etiquetas por Rollo No Válida")
      VDEF$ = OBX$
      GoTo 30
    End If
    '
    If RolloxCaja# < 1 Then
      Call COMUNI("Cantidad de Rollos por Cajas No Válida")
      VDEF$ = OBX$
      GoTo 30
    End If
    '
    ' ACTUALIZAR EN LOTINGRE
    EtiqxCaja# = EtiqxRollo# * RolloxCaja#
    '
    ' NO GRABA LA CANTIDAD DE ETIQUETAS POR CAJA CUANDO SE TRATA DE UN RE-IMPRESION DE ETIQUETA
    If MODO% = 0 Then Call ACTUREGISTRO("LOTINGRE", "EtiqxCaja", CONDI$, EtiqxCaja#, REG&)
    Call ACTUREGISTRO("LOTINGRE", "EtiqxRollo", CONDI$, EtiqxRollo#, REG&)
    Call ACTUREGISTRO("LOTINGRE", "RolxCaja", CONDI$, RolloxCaja#, REG&)
    '
    Call PRINETICIEROF(CI%, NORFA$, CANTI, EtiqxRollo#, RolloxCaja#)
    '

End Sub
Sub PRINETICIEROF(CI%, NUMORF$, CANTI, ETXROLLO, ROLXCAJA)

   If TRIM$(NUMORF$) = "" Then Exit Sub
   '
   FORIPRE$ = TRIM$(Control$("", "ETICIEOF"))

   If FORIPRE$ = "" Then
        Call MENSERR(24, "Falta Definir Formulario Etiqueta de Cierre de O.Fabricaciòn")
        Exit Sub
   End If
   '
   TIPODOC$ = "Etiqueta de Cierre de O.F" ' guardar el documento immpreso
   NUMEDOC$ = NUMORF$
   '
   CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODEXT$(CI%) 'CODIGO
   CODPARAM$(2) = "NUME-COM": DOCPARAM$(2) = NUMORF$ 'o.f
   CODPARAM$(3) = "DES-MAT": DOCPARAM$(3) = DESCRIT$(CI%) 'DESCRIPCION
   CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = ROLXCAJA 'CANTIDAD ROLLOS
   CODPARAM$(5) = "MEDIDA-2":  DOCPARAM$(5) = ETXROLLO 'ETIQUETA POR ROLLO
   CANPORCAJA = ROLXCAJA * ETXROLLO
   CODPARAM$(6) = "CANTIDAD": DOCPARAM$(6) = CANPORCAJA
     HOII% = HOY(HO$)
     FEC = FECHATEX$(HOII%)
   CODPARAM$(7) = "FECH-REC": DOCPARAM$(7) = FEC 'FECHA
     TXXT$ = AJUSTI$(CODEXT$(CI%), 16) & AJUSTI$(NUMORF$, 12) & FORMATNUM$(XVALO(CANPORCAJA), "F6.0")
     If InStr(1, EMPREC$, "ACHERNAR") > 0 Then TXXT$ = CODIBARRA$(TXXT$, 2)
   CODPARAM$(8) = "N-PLANO": DOCPARAM$(8) = TXXT$
   CAPARDOC% = 8
   '
   CANTETIXROL& = Int(CANTI / ETXROLLO) ' CALCULA CANTIDAD DE ETIQUETAS POR ROLLO
   RESTO# = CANTI - (CANTETIXROL& * ETXROLLO) ' SI NO ES EXACTO SUMA 1
   If RESTO# > 0 Then
      CANTETIXROL& = CANTETIXROL& + 1
   End If
   '
   CANTIETIAIMPRI& = Int(CANTETIXROL& / ROLXCAJA) ' CALCULA CANTIDAD DE ETIQUETAS A IMPRIMIR
   RESTO2# = CANTETIXROL& - (CANTIETIAIMPRI& * ROLXCAJA) ' SI NO ES EXACTO SUMA 1
   If RESTO2# > 0 Then
     CANTIETIAIMPRI& = CANTIETIAIMPRI& + 1
   End If
   '
   If CANTIETIAIMPRI& < 1 Then CANTIETIAIMPRI& = 1
   '
   PETI$ = TRIM$(Control$(IDENTER$, "PORTSTIK"))
   Call SETPRINPORT(PETI$)
   
   ACUMULO& = 0
   For I& = 1 To CANTIETIAIMPRI& - 1 ' IMPRIME TODAS MENOS LA ULTIMA
     ACUMULO& = ACUMULO& + CANPORCAJA ' ACUMULA LAS CANTIDADES
     Call PRINTDOC("@" + FORIPRE$)
   Next I&
   '
   If CANTI > ACUMULO& Then 'LA ULTIMA IMPRIME CALCULANDO LOS DATOS
     CANPORCAJA = CANTI - ACUMULO&
     ROLXCAJA = Int(CANPORCAJA / ETXROLLO)
     RESTO3# = CANPORCAJA - (ROLXCAJA * ETXROLLO)
     If RESTO3# > 0 Then
       ROLXCAJA = ROLXCAJA + 1
     End If
     '
     'IMPRESION DE ULTIMA ETIQUETA
     CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = "" 'CANTIDAD ROLLOS
     CODPARAM$(5) = "MEDIDA-2": DOCPARAM$(5) = "" 'CANTIDAD DE ETIQUETAS POR ROLLO
     CODPARAM$(6) = "CANTIDAD": DOCPARAM$(6) = CANPORCAJA
     TXXT$ = AJUSTI$(CODEXT$(CI%), 16) & AJUSTI$(NUMORF$, 12) & FORMATNUM$(XVALO(CANPORCAJA), "F6.0")
     CODPARAM$(8) = "N-PLANO": DOCPARAM$(8) = TXXT$
     Call PRINTDOC("@" + FORIPRE$)
     '
   End If
   '
90 Call SETPRINPORT("RESTORE")
   '
End Sub




