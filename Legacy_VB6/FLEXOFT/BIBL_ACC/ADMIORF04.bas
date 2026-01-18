Attribute VB_Name = "ADMIORF04"
Public CHECOMPO%
Public PRECOLAN%
Public PRECOCIE%
Public FRALOMAX%
Public VALINDOP%
Public INCLUPRI%
Dim APROLAN%
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
5   Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    With RESERTMP
10     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            CARESAN% = CARESAN% + 1
            ReDim Preserve CIXA%(CARESAN%), IDLA$(CARESAN%)
            CIXA%(CARESAN%) = !Cod_Inte
            IDLA$(CARESAN%) = !IdenLote
         .MoveNext
         Loop
       End If
    End With
    RESERTMP.Close
    '
    SQLX$ = "DELETE * FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
    Stocks.Execute SQLX$
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
       Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
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
           If XVALO(Mid$(!IdSubOr, 11)) = 1 Then
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
     '
     Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
     Call REPLA(VDEF$, !Referen, 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, !Unidad, 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, !Destino, 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(!Statuorf), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     For i& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", i&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next i&
     '

1420 YY$ = OBJPLA$("ORFALAN/Lanzamiento de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     '
     For i& = J& + 1 To ULTREG&("!ORFANUL")
        Call GRAREG("!ORFANUL", String$(256, 0), i&)
     Next i&
     '
     FLAN% = CVI(Mid$(YY$, 113, 2))
     If FLAN% < HO% Then FLAN% = HO%
     Call LANORF(NROOFL&, FLAN%)
     '
     If APROLAN% = 1 Then
        J& = 0
        For i& = 1 To ULTREG&("!ORFANUL")
           XX$ = REGLEIDO$("!ORFANUL", i&)
           If CVS(Left$(XX$, 4)) <> NROOFL& Then
              J& = J& + 1
              Call GRAREG("!ORFANUL", XX$, J&)
           End If
        Next i&
        Call SETULTREG("!ORFANUL", J&)
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
     Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
         If XVALO(Mid$(!IdSubOr, 4, 6)) <> ULORF& Then
           ZZ$ = REGBLAN$("!ORFANUL")
           REFXX$ = Left$(TRIM$(Left$(CODEXT$(!Cod_Inte), 15)) + " - " + TRIM$(DESCRIT0$(!Cod_Inte)), 50) + " - " + TRIM$(FORMATNUM$(!CanProy, "F10.1"))
           Call REPLA(ZZ$, MKS$(!NUMEORFA), 1, 4)
           Call REPLA(ZZ$, REFXX$, 5, 64)
           JJ& = JJ& + 1
           Call GRAREG("!ORFANUL", ZZ$, JJ&)
           ULORF& = XVALO(Mid$(!IdSubOr, 4, 6))
         End If
       .MoveNext
       Loop
     End With
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
     '
     Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
     Call REPLA(VDEF$, !Referen, 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, !Unidad, 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, !Destino, 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(1), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
1412 YY$ = OBJPLA$("ORFALAN/Suspensión de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     Screen.MousePointer = 11
     '
     Call ABREORFAB
     With OrdeFabr
       .FindFirst "NumeOrFa = " & NROOFL & " AND StatuOrf < 2"
1415   If .NoMatch = False Then
         .Edit
         !Statuorf = 2
         .Update
         '
         IDSUB$ = !IdSubOr
         With OperFab
           .FindFirst "IdSubOr = '" & IDSUB$ & "'"
1416       If .NoMatch = False Then
             .Edit
             !StatOper = 2
             .Update
             .FindNext "IdSubOr = '" & IDSUB$ & "'"
             GoTo 1416
           End If
         End With
         '
         .FindNext "NumeOrFa = " & NROOFL
         GoTo 1415
       End If
     End With
     OrdeFabr.Close
     OperFab.Close
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     Call COMUNI("La Orden de Fabricación " & NROOFL & "  \ha sido Suspendida")
     GoTo 100
     
'/////****ANTERIOR****/////
'    HO% = HOY(HOS$)
'    LU$ = LUDAT$
'    '
'    TITU$ = "PROCESO DE SUSPENSION DE ORDENES DE FABRICACION"
'
'100  NROOF$ = ""
'
'     '
'     'Call VENTANA(0, 0, 0, 0, "", "CLEAR")
'     'Call ENCAPAN(0, "ORDENES DE FABRICACION")
'     'Call PANTREC(4, 21, TITU$)
'     '
'1400 If INDI% = 0 Then
'       '
'       Screen.MousePointer = 11
'       Call COPYSTRU("ORFANUE", "ORFANUL")
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
'         If ST% = 1 Then      ' solo lanzadas
'            J& = J& + 1
'            Call GRAREG("ORFANUL", XX$, J&)
'         End If
'       Next I&
'       '
'       If J& < 1 Then
'          Call MENSERR(24, "No hay Ordenes de Fabricación\Lanzadas que Puedan Suspenderse.")
'          Call CIERRARCH("ORFANUL")
'          Screen.MousePointer = 1
'          Call CIERRARCH("*.*")
'          Exit Sub
'       End If
'       '
'       For I& = J& + 1 To ULTREG&("ORFANUL")
'         Call GRAREG("ORFANUL", String$(256, 0), I&)
'       Next I&
'       '
'       Call CIERRARCH("ORFANUL")
'       INDI% = 1
'       '
'       Screen.MousePointer = 1
'       '
'     End If
'
'1405 Call FRESHECHO("ORFANUL")
'     If ULTREG&("ORFANUL") < 1 Then INDI% = 0: Exit Sub
'     Call SELECHO("ORFANUL/Ordenes de Trabajo para Suspensión")
'     '
'1410 NROOFL& = XVALO(VALACT1$("ORFANUL"))
'     If NROOFL& < 1 Then
'         INDI% = 0
'         Exit Sub
'     End If
'     '
'     For I& = ULTREG&("ORFANUE") To 1 Step -1
'       XX$ = REGLEIDO$("ORFANUE", I&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
'     Next I&
'     '
'     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
'     INDI% = 0
'     Exit Sub
'     '
'1420 YY$ = OBJPLA$("ORFALAN/Suspensión de Ordenes de Fabricación", XX$)
'     If YY$ = "" Then GoTo 1405
'     '
'     Screen.MousePointer = 11
'     For J& = I& To 1 Step -1
'       XX$ = REGLEIDO$("ORFANUE", J&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then
'          Call REPLA(XX$, Chr$(2), 112, 1)
'          Call GRAREG("ORFANUE", XX$, J&)
'       End If
'     Next J&
'     '
'     J& = 0
'     For I& = 1 To ULTREG&("ORFANUL")
'        XX$ = REGLEIDO$("ORFANUL", I&)
'        If CVS(Left$(XX$, 4)) <> NROOFL& Then
'           J& = J& + 1
'           Call GRAREG("ORFANUL", XX$, J&)
'        End If
'     Next I&
'     '
'     For I& = J& + 1 To ULTREG&("ORFANUL")
'        Call GRAREG("ORFANUL", String$(256, 0), I&)
'     Next I&
'     '
'     For J& = 1 To ULTREG&("ORDEFABR")
'       XX$ = REGLEIDO$("ORDEFABR", J&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then
'          Call REPLA(XX$, Chr$(2), 108, 1)
'          Call GRAREG("ORDEFABR", XX$, J&)
'       End If
'     Next J&
'     '
'     Screen.MousePointer = 1
'     Call CIERRARCH("*.*")
'     GoTo 1405
'     '

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
     Set ORFATEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     On Error Resume Next
     With ORFATEMP
       .MoveFirst
       If Err Then
         Call MENSERR(24, "No hay Ordenes de Fabricación\en Curso para Cambio de Prioridad.")
         Screen.MousePointer = 1
         Exit Sub
       End If
       On Error GoTo 0
       '
       JJ& = 0
       Do While Not .EOF
         If XVALO(Mid$(!IdSubOr, 11)) = 1 Then
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
     '
     Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     VDEF$ = REGBLAN$("ORFANUE")
     With CARGAOBJTMP
       On Error Resume Next
       .MoveFirst
       On Error GoTo 0
       Call REPLA(VDEF$, MKS$(!NUMEORFA), 1, 4)
       Call REPLA(VDEF$, !Referen, 5, 64)
       Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
       Call REPLA(VDEF$, !Unidad, 73, 4)
       Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
       Call REPLA(VDEF$, MKS$(!CanProy), 97, 4)
       Call REPLA(VDEF$, !Destino, 101, 8)
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
     With OrdeFabr
       .FindFirst "NumeOrfa = " & NROOFL&
1425   If .NoMatch = False Then
         .Edit
         !Priorid = PRIO%
         !SecciPro = SEPROD
         !FechEnSol = CVDAT(ENTSOL%)
         !FechLan = CVDAT(LANPRO%)
         .Update
         .FindNext "NumeOrfa = " & NROOFL&
         GoTo 1425
       End If
     End With
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
     HO% = HOY(HOS$)
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
        If CONTROL$("", "PRECOCIE") = "SI" Then
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
104  If EXISTE%("TRAZAB04.EXE") > 0 Then
        Call CONECRUN("TRAZAB04/OFABRIC" + NROOF$, "Registro de Trazabilidad de O/F's")
        Exit Sub
     End If
     '
     Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "'"
     Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With ORDEFABTMP
       On Error Resume Next
       .MoveFirst
       If Err Then GoTo 1410
       On Error GoTo 0
       JJ& = 0
         Y$ = REGBLAN("ORDEFABR")
         '
         Call REPLA(Y$, MKS$(!NUMEORFA), 1, 4)
         Call REPLA(Y$, !Referen, 5, 64)
         Call REPLA(Y$, !IdSubOr, 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechGen)), 81, 2)
         Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
         Call REPLA(Y$, !Unidad, 85, 4)
         Call REPLA(Y$, MKD$(!CanProy), 89, 8)
         '
         On Error Resume Next
         Call REPLA(Y$, !Destino, 97, 8)
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
           CAPENDE# = !CanProy - !CanApro: If CAPENDE# < 0 Then CAPENDE# = 0
           CAPEN1# = CAPENDE#
           If CANTAPROX > 0 Then CAPEN1# = CANTAPROX
         Call REPLA(Y$, MKD$(CAPEN1#), 119, 8)
         Call REPLA(Y$, MKD$(0), 127, 8)
         Call REPLA(Y$, !IdenRetra, 135, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechProyEntre)), 147, 2)
         Call REPLA(Y$, MKI$(CVINT(!FechTempEntre)), 149, 2)
         Call REPLA(Y$, MKI$(CVINT(!FechRecal)), 151, 2)
         Call REPLA(Y$, !Observa, 155, 64)
     End With
     On Error Resume Next
     ORDEFABTMP.Close
     Set ORDEFABTMP = Nothing
     On Error GoTo 0
     '
     CODI1% = CVI(Mid(Y$, 83, 2)) 'PARA MOSTRAR EL DEPOSITO DE CONSUMO Y
     DEPENT1% = DEPENART%(CODI1%) 'DE ENTRADA QUE TIENE POR DEFAULT
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
     If CANTAPROX >= 0.05 Then
'        If Abs(CANTI - CANTAPROX) > 0.05 Or Abs(SCRAP) >= 0.05 Then
         If Abs(CANTI + SCRAP) - CANTAPROX > 0.05 Then
'           Call MENSERR(24, "En Cierre por Parte de Producción\no Válido Cambiar Cantidades.")
            Call MENSERR(24, "En Cierre por Parte de Producción\Cantidades Difieren del Total Informado")
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
     CODICONJ$ = CODEXT$(CI%)
     CODPRO$ = CODICONJ$
     CICONJ% = CI%
     NORFABRI$ = NORFA$
     CACONJU# = CDbl(CANTI + SCRAP)
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
          MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
          If MDXX% = 2 Then CIERRA% = 1
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
     If PUEDEMOVI%(DP2%) < 0 Then GoTo 1420 ' CANCELA SALIDA
     '\\\OT-05-0730-WAR-FIN///
     '
'1.- Actualiza saldo pendiente de la Orden de Fabricacion
     Call ABREORFAB
     With OrdeFabr
       .FindFirst "IdSubOr = '" & NORFA$ & "'"
       If .NoMatch = False Then
         .Edit
         !CanApro = !CanApro + CANTI
         !CanRech = !CanRech + SCRAP
         If CIERRA% = 1 Then
           !Statuorf = 3         ' MARCA COMO CERRADA
         End If
         '\\\OT-06-0044-WAR-07/02/06/// 'TOTAL O PARCIAL
         !FechCierr = CVDAT(FECHANUM(FECHACIOF$))
         '\\\OT-06-0044-WAR-FIN///
         .Update
       End If
     End With
     '
'2.- Genera Movimiento de entrada P.Terminado
     If CANTI >= 0.05 Then
        DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
        NUORIT% = 0
        Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, MONEII%, NC%, DEPX%, CANTI)
        CIMOVI%(CICONJ%) = 1
     End If
     '
     If CICONJ% > 0 Then
        Master.FindFirst ("CODINT = " & CICONJ%)
        If Master.NoMatch = False Then
           Master.Edit
           Master!DepoEntra = DEPX%
           Master!DepoCon = DP2%
           Master.Update
        End If
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
             LOTE$ = ""
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
                Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                CIMOVI%(CI2%) = 1
             End If
             REGMOVI& = ULTREG&("MOVISTO")
             '
             SALCON# = (-1) * CAN
             With Reserva
               .FindFirst "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2%
145            If .NoMatch = False Then
                 .Edit
                 CARESE# = !CantRes
                 CACOANT# = !CantCons
                 If CARESE# - CACOANT# >= SALCON# Then
                      !CantCons = CACOANT# + SALCON#
                      SALCON# = 0
                    Else
                      SALCON# = SALCON# - (CARESE# - CACOANT#)
                      !CantCons = !CantRes
                 End If
                 .Update
                 If SALCON# >= 0.05 Then
                   .FindNext "IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & CI2%
                   GoTo 145
                 End If
               End If
             End With
           End If
         End If
       End If
     Next U&
     '
     Reserva.Close
     Call CIERRARCH("*.*")
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
         If CONTROL$("ORDEFABR", "MODOSCRA") = "PRODUCTO" Then
             CI2% = CODINT%(CODICONJ$)
             LOTE$ = NORFA$
             CMOV$ = "SC"
             DOPRI$ = NORFA$
             DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
             REFE$ = "SCRAP " + CODICONJ$
             DEPX% = DSC%: If DSC% = 99 Then DEPX% = DEPOPRE%(CI2%)
             CAN = SCRAP
             Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
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
                    DEPX% = DSC%: If DSC% = 99 Then DEPX% = DEPOPRE%(CI2%)
                    NUORIT% = NUORIT% + 1
                    Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
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
     'Call VESARRGLO2
     '
     Movisto.Close
     Call CIERRARCH("MOVISTO")
     '
     If CIERRA% = 1 Then             ' cancela operaciones
        '
'5.- Cancela Operaciones Programadas
        Call ABREORFAB
        SQLX$ = "DELETE * FROM OperFab WHERE IdSubOr = '" & NORFA$ & "'"
        Ofabric.Execute SQLX$
    '    With OperFab
146 '      .FindFirst "IdSubOr = '" & NORFA$ & "'"
    '      If .NoMatch = False Then
    '        .Delete
    '        GoTo 146
    '      End If
    '    End With
        '
'6.- Cancela Reservas de Componentes
        Call ABRESERVA
        SQLX$ = "DELETE * FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
        Stocks.Execute SQLX$
    '    With Reserva
147 '      .FindFirst "IdSubOr = '" & NORFA$ & "'"
    '      If .NoMatch = False Then
    '        .Delete
    '        GoTo 147
    '      End If
    '    End With
        '
'7.- Cancela Orden en Indice de Ordenes Pendientes
        Call FILTERDELETE("ECORDEP", 1, NORFA$)
        Call FRESHECHO("ECORDEP")
        '
    End If
    '
149 Screen.MousePointer = 1
    Call LIBARCH("RESERVOF")
    Call CIERRARCH("*.*")
    Call LIBARCH("ECORDEP")
    '
'8.- Imprime Comprobante de cierre de O/F
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
98  If CANTAPROX < 0.05 Then     ' SOLO EN CASO DE CIERRE MANUAL - NO POR PARTE DIARIO
       GoTo 100
    End If
    '
End Sub
'
Sub ORFANU()
'
    HO% = HOY(HOS$)
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
       Set ORFANUTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
             Call REPLA(YY$, !IdSubOr, 1, 12)
             Call REPLA(YY$, !Referen, 13, 52)
             JJ& = JJ& + 1
             Call GRAREG("!ECORANU", YY$, JJ&)
         .MoveNext
         Loop
       End With
       Call SETULTREG("!ECORANU", JJ&)
       Call CIERRARCH("!ECORANU")
       Screen.MousePointer = 1
     End If
     '
     Screen.MousePointer = 1
     '
1405 Call FRESHECHO("!ECORANU")
     If ULTREG&("!ECORANU") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("!ECORANU/Ordenes de Trabajo para Anulación")
     '
1410 NORFA$ = AJUSTI$(VALACT1$("!ECORANU"), 12)
     If TRIM$(NORFA$) = "" Then
         INDI% = 0
         Exit Sub
     End If
     '
     Screen.MousePointer = 11
     '
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
     '
     Set CARGAOBJTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With CARGAOBJTMP
       On Error Resume Next
      .MoveFirst
       On Error GoTo 0
     NROOFL = !NUMEORFA
     Call REPLA(VDEF$, MKS$(NROOFL), 1, 4)
     Call REPLA(VDEF$, !Referen, 5, 64)
     Call REPLA(VDEF$, MKI$(CVINT(!FechGen)), 69, 2)
     Call REPLA(VDEF$, !Unidad, 73, 4)
     Call REPLA(VDEF$, MKD$(!CanProy), 89, 8)
     Call REPLA(VDEF$, !Destino, 101, 8)
     Call REPLA(VDEF$, MKI$(CVINT(!FechEnSol)), 109, 2)
     Call REPLA(VDEF$, Chr$(!Priorid), 111, 1)
     Call REPLA(VDEF$, Chr$(!Statuorf), 112, 1)
     Call REPLA(VDEF$, MKI$(CVINT(!FechLan)), 113, 2)
     Call REPLA(VDEF$, MKI$(CVINT(!FechCierr)), 119, 2)
     End With
     '
1412 YY$ = OBJPLA$("ORFALAN/Anulacion de Ordenes de Fabricación", VDEF$)
     If YY$ = "" Then GoTo 1405
     '
     'ANULACION OF
     '
     Call ABREORFAB
     Call ABRESERVA
     With OrdeFabr
       .FindFirst "NumeOrFa = " & NROOFL & "and IDSUBOR= '" & NORFA$ & "'"
1415   If .NoMatch = False Then
         .Edit
         !Statuorf = 9
         .Update
         '
         IDSUB$ = !IdSubOr
         '
         With Reserva
            .FindFirst "IDSUBOR= '" & IDSUB$ & "'"
1416        If .NoMatch = False Then
              .Edit
              !CantRes = !CantCons
              !Saldo_Entre = !CantRes - !CantCons
              .Update
              With OperFab
                .FindFirst "IdSubOr = '" & IDSUB$ & "'"
1417            If .NoMatch = False Then
                  .Delete
                  .FindNext "IdSubOr = '" & IDSUB$ & "'"
                  GoTo 1417
                End If
              End With
              .FindNext "IdSubOr = '" & IDSUB$ & "'"
              GoTo 1416
            End If
         End With
         '
'         .FindNext "NumeOrFa = " & NROOFL
'         GoTo 1415
       End If
     End With
     OrdeFabr.Close
     OperFab.Close
     '
     If EXISTE%("TRAZAB04.EXE") > 0 Then
        Call ANULARES(NORFA$)
     End If
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
     Screen.MousePointer = 1
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
     With OrdeFabr
       .FindFirst "NumeOrFa = " & NROOFL
1415   If .NoMatch = False Then
         .Edit
         !Statuorf = 1
         !FechLan = CVDAT(FLAN%)
         .Update
         '
         IDSUB$ = !IdSubOr
         With OperFab
           .FindFirst "IdSubOr = '" & IDSUB$ & "'"
1416       If .NoMatch = False Then
             .Edit
             !StatOper = 1
             .Update
             .FindNext "IdSubOr = '" & IDSUB$ & "'"
             GoTo 1416
           End If
         End With
         '
         .FindNext "NumeOrFa = " & NROOFL
         GoTo 1415
       End If
     End With
     OrdeFabr.Close
     OperFab.Close
     '
     Screen.MousePointer = 1
     APROLAN% = 1
     

  
'//////****ANTIGUO****///////
'     '
'     Screen.MousePointer = 11
'     APROLAN% = 0
'     '
'     FIN& = ULTREG&("ORFANUE")
'     For J& = FIN& To 1 Step -1
'       XX$ = REGLEIDO$("ORFANUE", J&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then
'          cix% = CVI(Mid$(XX$, 71, 2))
'          CANORF = CVD(Mid$(XX$, 89, 8))
'          If HAYDISTO%(cix%, NROOFL&, CANORF) <> 1 Then GoTo 1499
'          Call REPLA(XX$, Chr$(1), 112, 1)
'          Call REPLA(XX$, MKI$(FLAN%), 113, 2)
'          Call GRAREG("ORFANUE", XX$, J&)
'          GoTo 1510
'       End If
'     Next J&
'     '
'     Call MENSERR(24, "Lanzamiento Imposible.\Orden no Encontrada.")
'     Exit Sub
'     '
'1510 SECUOR$ = ""
'     For J& = 1 To ULTREG&("ORDEFABR")
'       XX$ = REGLEIDO$("ORDEFABR", J&)
'       If CVS(Left$(XX$, 4)) = NROOFL& Then
'          Call REPLA(XX$, Chr$(1), 108, 1)
'          Call REPLA(XX$, MKI$(FLAN%), 109, 2)
'          Call GRAREG("ORDEFABR", XX$, J&)
'          SECUOR$ = SECUOR$ + MKS$(J&)
'       End If
'     Next J&
'     '
'     Call ACTECOR(0)
'     '
'     For UUU% = 1 To Len(SECUOR$) Step 4
'        REGORDE& = CVS(Mid$(SECUOR$, UUU%, 4))
'        If REGORDE& > 0 Then
'           If REGORDE& <= ULTREG&("ORDEFABR") Then
'              XX$ = REGLEIDO$("ORDEFABR", REGORDE&)
'              IDENLO$ = Mid$(XX$, 69, 12)
'              CI% = CVI(Mid$(XX$, 83, 2))
'              CAPENDE# = CVD(Mid$(XX$, 89, 8))
'              DESTI$ = Mid$(XX$, 97, 8)
'              REFE$ = Mid$(XX$, 5, 64)
'              '
'              Call BLANARCH("!OPERFAB")
'
'              If CI% > 0 Then
'                 If CI% <= NUMAS% Then
'                    '
'GoTo 1480  ' FALTA REVISAR
'                    J1& = 0: YAOPER% = 0
'                    For J& = 1 To ULTREG&("OPERFAB")
'                       YY$ = REGLEIDO$("OPERFAB", J&)
'                       NORFA$ = Left$(YY$, 12)
'                       If NORFA$ = IDENLO$ Then
'                           Call REPLA(YY$, String$(4, 0), 135, 4)
'                           If CVD(Mid$(YY$, 79, 8)) <> CAPENDE# Then
'                              Call COMUNI("Diferencia en las Cantidades Proyectadas\en " + TRIM$(NORFA$) + " Restaurada Automaticamente.")
'                              Call REPLA(YY$, MKD$(CAPENDE#), 79, 8)
'                              CAPEN1# = CAPENDE# - CVD(Mid$(YY$, 87, 8))
'                              If CAPEN1# < 0 Then CAPEN1# = 0
'                              Call REPLA(YY$, MKD$(CAPEN1#), 95, 8)
'                           End If
'                           Call REGAPP("!OPERFAB", YY$)
'                           YAOPER% = 1
'                         Else
'                           J1& = J1& + 1
'                           Call GRAREG("OPERFAB", YY$, J1&)
'                       End If
'                    Next J&
'                    Call SETULTREG("OPERFAB", J1&)
'                    Call CIERRARCH("OPERFAB")
'                    '
''                   programar operaciones standard
'                    '
'                    If YAOPER% = 0 Then
'                      PRISECO& = 0
'                      For TTL& = 1 To ULTREG&("SECOPER")
'                        SCP$ = REGLEIDO$("SECOPER", TTL&)
'                        If CVI(Left$(SCP$, 2)) = CI% Then
'                          PRISECO& = TTL&
'                          GoTo 1450
'                        End If
'                      Next TTL&
'                      '
'                      If PRISECO& > 0 Then
'                        Call BLANARCH("!OPERFAB")
'1450                    For TTL& = PRISECO& To ULTREG&("SECOPER")
'                          SCP$ = REGLEIDO$("SECOPER", TTL&)
'                          If CVI(Left$(SCP$, 2)) = CI% Then
'                            NO% = CVI(Mid$(SCP$, 3, 2))
'                            DOPE$ = Mid$(SCP$, 5, 48)
'                            EQ1$ = Mid$(SCP$, 53, 6)
'                            EQAUX$ = Mid$(SCP$, 103, 12)
'                            LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
'                            CALOTES = 1
'                            MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
'                            MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
'                            MINVARI = CAPENDE# * CVS(Mid$(SCP$, 67, 4))
'                            MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
'                            MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
'                            MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
'                            '
'                            HORAPROD = MINLOTE / 60
'                            CAOPS = CVS(Mid$(SCP$, 83, 4))
'                            '
'' YA ESTA
'                            BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
'                            NUOROP% = 0
'                            FEAL% = HOY(HOS$)
'                            HOAL% = HORANUM(Time$)
'                            USUA% = USNBR% Mod 100
'                            DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
'                            INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
'                            PROGBLA$ = Space$(6) + String$(26, 0)
'                            '
'                            ZZ$ = REGBLAN$("OPERFAB")
'                            Call REPLA(ZZ$, IDENLO$, 1, 12)
'                            Call REPLA(ZZ$, MKI$(CI%), 13, 2)
'                            Call REPLA(ZZ$, MKI$(NO%), 15, 2)
'                            Call REPLA(ZZ$, DOPE$, 17, 48)
'                            Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
'                            Call REPLA(ZZ$, DALTA$, 67, 6)
'                            Call REPLA(ZZ$, EQ1$, 73, 6)
'                            Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
'                            Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
'                            Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
'                            Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
'                            Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
'                            Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
'                            Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
'                            Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
'                            Call REPLA(ZZ$, String$(16, 0), 135, 16)
'                            Call REPLA(ZZ$, INFOBLA$, 151, 48)
'                            Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
'                            Call REPLA(ZZ$, EQAUX$, 203, 18)
'                            Call REPLA(ZZ$, MKS$(1), 221, 4)
'                            Call REPLA(ZZ$, PROGBLA$, 225, 32)
'                            '
'                            Call REGAPP("!OPERFAB", ZZ$)
'                            '
'                          End If
'                        Next TTL&
'                      End If
'                      '
'                      SECUOPSTB$ = RECFILT$("BCOMPLEM", 1, MKI$(CI%))
'                      If Len(SECUOPSTB$) > 0 Then
'                        For TTI% = 1 To Len(SECUOPSTB$) Step 4
'                          TTL& = CVS(Mid$(SECUOPSTB$, TTI%, 4))
'                          SCP$ = REGLEIDO$("BCOMPLEM", TTL&)
'                          If CVI(Left$(SCP$, 2)) = CI% Then
'                            NO% = 999
'                            DOPE$ = Mid$(SCP$, 3, 14)
'                            EQ1$ = Mid$(SCP$, 33, 6)
'                            EQAUX$ = String$(18, 0)
'                            LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
'                            CALOTES = 1
'                            MINPREP = CVS(Mid$(SCP$, 65, 4))
'                            MINFIJO = 0
'                            MINVARI = CVS(Mid$(SCP$, 73, 4))
'                            MINDEMO = 0
'                            MINLIMP = 0
'                            MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
'                            '
'                            If MINLOTE > 0.1 Then
'                              HORAPROD = MINLOTE / 60
'                              CAOPS = 1
'                              '
'' YA ESTA
'                              BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
'                              NUOROP% = 0
'                              FEAL% = HOY(HOS$)
'                              HOAL% = HORANUM(Time$)
'                              USUA% = USNBR% Mod 100
'                              DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
'                              INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
'                              PROGBLA$ = Space$(6) + String$(26, 0)
'                              '
'                              ZZ$ = REGBLAN$("OPERFAB")
'                              Call REPLA(ZZ$, IDENLO$, 1, 12)
'                              Call REPLA(ZZ$, MKI$(CI%), 13, 2)
'                              Call REPLA(ZZ$, MKI$(NO%), 15, 2)
'                              Call REPLA(ZZ$, DOPE$, 17, 48)
'                              Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
'                              Call REPLA(ZZ$, DALTA$, 67, 6)
'                              Call REPLA(ZZ$, EQ1$, 73, 6)
'                              Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
'                              Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
'                              Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
'                              Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
'                              Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
'                              Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
'                              Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
'                              Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
'                              Call REPLA(ZZ$, String$(16, 0), 135, 16)
'                              Call REPLA(ZZ$, INFOBLA$, 151, 48)
'                              Call REPLA(ZZ$, MKS$(0), 199, 4)
'                              Call REPLA(ZZ$, EQAUX$, 203, 18)
'                              Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
'                              Call REPLA(ZZ$, PROGBLA$, 225, 32)
'                              '
'                              Call REGAPP("!OPERFAB", ZZ$)
'                            End If
'                          End If
'                        Next TTI%
'                      End If
'                      '
'                    End If
'                    Call CIERRARCH("!OPERFAB")
'                    '
'                    For TTL& = 1 To ULTREG&("!OPERFAB")
'                      ZZ$ = REGLEIDO$("!OPERFAB", TTL&)
'                      Call REGAPP("OPERFAB", ZZ$)
'                    Next TTL&
'                    Call CIERRARCH("OPERFAB")
'                    Call CIERRARCH("!OPERFAB")
'                    '
'1480                NROFLLA$ = IDENLO$
'                    Call GENRESER(NROFLLA$)
'                    '
'                 End If
'              End If
'           End If
'        End If
'1490 Next UUU%
'     APROLAN% = 1
'     '
'1499 Screen.MousePointer = 1
'     Call CIERRARCH("*.*")

End Sub
'
Sub ORFTINT()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     Call ACTECOR(0)
     '
     TITU$ = "ORDEN DE DISPOSICION DE TINTORERIA"

100  Screen.MousePointer = 1
     OBX1$ = OBJPLA$("SELORTINTO", "")
     If TRIM$(OBX1$) = "" Then
         Exit Sub
     End If
     '
     NROOF$ = Left$(OBX1$, 12)
     PROVEE% = CVI(Mid$(OBX1$, 13, 2))
     '
     Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada como Pendiente")
         GoTo 100
     End If
     '
     If PROVEE% < 1 Or ECOARCH$("PROVED1", MKI$(PROVEE%)) = "" Then
         Call MENSERR(24, "Proveedor no Figura en Padrón")
         GoTo 100
     End If
     '
     For i& = ULTREG&("ORDEFABR") To 1 Step -1
       XX$ = REGLEIDO$("ORDEFABR", i&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
          CAPREN# = CVD(Mid$(XX$, 89, 8))
          If CVD(Mid$(XX$, 119, 8)) > CAPREN# Then
              CAPREN# = CVD(Mid$(XX$, 119, 8))
          End If
          CIBAS% = CVI(Mid$(XX$, 83, 2))
          CODCON$ = CODEXT$(CIBAS%)
          DECON$ = DESCRIT0$(CIBAS%)
          REGOR& = i&
          GoTo 1420
       End If
     Next i&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     GoTo 100
     '
1420 Screen.MousePointer = 11
     NUOF& = XVALO(Mid$(NORFA$, 4))
     Call BLANARCH("!DISPOTIN")
     For U& = 1 To ULTREG&("ESTRUNUE")
        XX$ = REGLEIDO$("ESTRUNUE", U&)
        CIPRE% = CVI(Left$(XX$, 2))
        R2$ = Mid$(XX$, 3, 120) + REGLEIDO$("ESTRUDOS", U&)
        For K% = 1 To Len(R2$) Step 8
           DPSX$ = Mid$(R2$, K%, 2)
           If CVI(DPSX$) = CIBAS% Then
              YY$ = REGBLAN$("DISPOTIN")
              Call REPLA(YY$, MKS$(NUOF&), 1, 4)
              Call REPLA(YY$, MKI$(CIBAS%), 5, 2)
              Call REPLA(YY$, MKI$(CIPRE%), 7, 2)
              Call REPLA(YY$, COLOR$(CIPRE%), 17, 32)
              Call REPLA(YY$, TALLE$(CIPRE%), 49, 16)
              Call REPLA(YY$, MKI$(HO%), 129, 2)
              Call REPLA(YY$, MKI$(PROVEE%), 149, 2)
              Call REGAPP("!DISPOTIN", YY$)
              GoTo 1424
           End If
        Next K%
1424 Next U&
     '
     CADISOTPRO = 0
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVS(Left$(XX$, 4)) = NUOF& Then
           CAPAR = CVS(Mid$(XX$, 9, 4))
           If CVI(Mid$(XX$, 149, 2)) <> PROVEE% Or CVI(Mid$(XX$, 131, 2)) > 0 Then
              CADISOTPRO = CADISOTPRO + CAPAR
              GoTo 1427
           End If
           CIPRE% = CVI(Mid$(XX$, 7, 2))
           For V& = 1 To ULTREG&("!DISPOTIN")
              YY$ = REGLEIDO$("!DISPOTIN", V&)
              If Left$(YY$, 8) = Left$(XX$, 8) Then
                 Call GRAREG("!DISPOTIN", XX$, V&)
              End If
           Next V&
        End If
1427 Next U&
     '
     Call CIERRARCH("!DISPOTIN")
     Screen.MousePointer = 1
     '
1430 TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAPREN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODCON$) + " - " + DECON$
     VTB% = VENTABU%("DISPOTIN/TITUPAN=DISPOSICION PARA TINTORERIA" + TITU$)
     If VTB% < 0 Then GoTo 100
     '
     CADISPO = 0
     For U& = 1 To ULTREG&("!DISPOTIN")
        XX$ = REGLEIDO$("!DISPOTIN", U&)
        CAPAR = CVS(Mid$(XX$, 9, 4))
        CIPRE% = CVI(Mid$(XX$, 7, 2))
        KILOS = CAPAR * PESUNI(CIPRE%)
        Call REPLA(XX$, MKS$(KILOS), 13, 4)
        Call GRAREG("!DISPOTIN", XX$, U&)
        CADISPO = CADISPO + CAPAR
     Next U&
     '
     If CADISPO + CADISOTPRO > CAPREN# + 0.5 Then
        OPXX% = COMALTER%("ATENCION: El Total Dispuesto (" + TRIM$(Str$(Int(CADISPO + CADISOTPRO + 0.5))) + ")\Excede la Cantidad Según O/F (" + TRIM$(Str$(Int(CAPREN# + 0.5))) + ")\\Corregir\Continuar")
        If OPXX% <> 2 Then GoTo 1430
     End If
     '
     J& = 0
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVS(Left$(XX$, 4)) <> NUOF& Or CVI(Mid$(XX$, 149, 2)) <> PROVEE% Or CVI(Mid$(XX$, 131, 2)) > 0 Then
           J& = J& + 1
           Call GRAREG("DISPOTIN", XX$, J&)
        End If
     Next U&
     '
     For U& = 1 To ULTREG&("!DISPOTIN")
        XX$ = REGLEIDO$("!DISPOTIN", U&)
        If CVS(Mid$(XX$, 9, 4)) > 0 Then
           J& = J& + 1
           Call GRAREG("DISPOTIN", XX$, J&)
        End If
     Next U&
     '
     Call CIERRARCH("*.*")
     Call COMUNI("Disposición Completa")
     GoTo 100
     '
End Sub
'


Sub REMITINTO()
     '
     Static DATRANSPO$
     Dim CPREN%(1024), CADIS(1024)
     Dim CCRU%(512), CACRU(512), REGMAC&(512)
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
50   REGXX& = ULTREG&("OCOMPRA")
     NRO& = 1
     If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
     End If
     If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
           X0$ = REGLEIDO$("OCOMPRA", U&)
           If CVS(Left$(X0$, 4)) >= NRO& Then
              NRO& = 1 + CVS(Left$(X0$, 4))
           End If
        Next U&
        Screen.MousePointer = 1
     End If
     '
100  Screen.MousePointer = 1
     EBOL$ = OBJPLA$("SELPROVE", "")
     If EBOL$ = "" Then
         Exit Sub
     End If
     '
     PROVEE% = CVI(EBOL$)
     If PROVEE% < 1 Or ECOARCH$("PROVED1", MKI$(PROVEE%)) = "" Then
        Call MENSERR(24, "Proveedor no Figura en Padrón")
        GoTo 100
     End If
     '
     Screen.MousePointer = 11
     '
     FECHDOC$ = HOS$
     NUMEDOC$ = TRIM$(Str$(NRO&))
     While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
     Wend
     TIPODOC$ = "Orden de Disposicion"
     '
     CODTABU1$(1) = "COD-MAT"
     CODTABU1$(2) = "TALLE"
     CODTABU1$(3) = "COLOR"
     CODTABU1$(4) = "CANTIDAD"
     CODTABU1$(5) = "CAKILOS"
     CACOLTAB1% = 5
     '
     DESTIDOC% = (-1) * PROVEE%
     CAITAB1% = 0
     CAITCRU% = 0
     '
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVI(Mid$(XX$, 149, 2)) = PROVEE% Then
          If CVI(Mid$(XX$, 131, 2)) <= 0 Then
            CAPORCOL = CVS(Mid$(XX$, 9, 4))
            If CAPORCOL > 0 Then
              CICRU% = CVI(Mid$(XX$, 5, 2))
              CIPRE% = CVI(Mid$(XX$, 7, 2))
              TALLE1$ = Mid$(XX$, 49, 16)
              COCOL1$ = Mid$(XX$, 17, 32)
              CACOL1$ = Str$(CAPORCOL)
              KICOL1$ = Str$(CVS(Mid$(XX$, 13, 4)))
              '
              Call REPLA(XX$, MKI$(HO%), 131, 2)
              Call REPLA(XX$, MKS$(NRO&), 133, 4)
              Call GRAREG("DISPOTIN", XX$, U&)
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
              TETABU1$(2, CAITAB1%) = TALLE1$
              TETABU1$(3, CAITAB1%) = COCOL1$
              TETABU1$(4, CAITAB1%) = CACOL1$
              TETABU1$(5, CAITAB1%) = KICOL1$
              '
              CPREN%(CAITAB1%) = CIPRE%
              CADIS(CAITAB1%) = CAPORCOL
              '
              For J% = 1 To CAITCRU%
                 If CCRU%(J%) = CICRU% Then
                    CACRU(J%) = CACRU(J%) + CAPORCOL
                    GoTo 109
                 End If
              Next J%
              CAITCRU% = CAITCRU% + 1
              CCRU%(CAITCRU%) = CICRU%
              CACRU(CAITCRU%) = CAPORCOL
              '
            End If
          End If
        End If
109  Next U&
     '
     If CAITAB1% < 1 Then
        Call MENSERR(24, "No hay Disposiciones Pendientes\para " + TRIM$(RASOCLI$((-1) * PROVEE%)) + ".")
        GoTo 100
     End If
     '
     Call PRINTDOC("FORDITIN")   ' Orden de Disposicion
     '
    'FENTRE% = CVI(Mid$(X1$, 9, 2))
    'If FENTRE% < FEMI% Then
    '    Call MENSERR(24, "Fecha de Entrega Incoherente")
    '    Call REPLA(X0$, MKI$(FEMI%), 9, 2)
    '    GoTo 101
    'End If
    '
     X1$ = REGBLAN$("OCOMPRA")
     Call REPLA(X1$, MKS$(NRO&), 1, 4)
     Call REPLA(X1$, MKI$(HO%), 5, 2)
     Call REPLA(X1$, MKI$(PROVEE%), 7, 2)
     Call REPLA(X1$, MKI$(HO%), 9, 2) ' fecha de entrega)
     Call REPLA(X1$, Chr$(2), 57, 1)
     Call REPLA(X1$, Chr$(0), 62, 1)  ' abierta
     '
     For KKI% = 1 To CAITAB1%
        Call REPLA(X1$, MKI$(CPREN%(KKI%)), 11, 2)
        Call REPLA(X1$, MKS$(CADIS(KKI%)), 13, 4)
        Call REGAPP("OCOMPRA", X1$)
     Next KKI%
     '
     NOCOM$ = TRIM$(Str$(NRO&))
     While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
     NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
     '
     For i% = 1 To CAITCRU%
        YY$ = REGBLAN$("MACONSIG")
        Call REPLA(YY$, MKI$(HO%), 1, 2)
        Call REPLA(YY$, MKI$(PROVEE%), 3, 2)
        Call REPLA(YY$, NOCOM$, 21, 12)
        '
        CII% = CCRU%(i%)
        CANTI# = CDbl(CACRU(i%))
        REGMAC&(i%) = 0
        '
        If TRIM$(CODEXT$(CII%)) <> "" Then
           If CANTI# > 0.05 Then
              Call REPLA(YY$, MKI$(CII%), 33, 2)
              Call REPLA(YY$, UNIMED$(CII%), 51, 4)
              Call REPLA(YY$, MKD$(CANTI#), 55, 8)
              Call REGAPP("MACONSIG", YY$)
              REGMAC&(i%) = ULTREG&("MACONSIG")
           End If
        End If
     Next i%
     '
' REMITO EN CONSIGNACION ENTRECONSI
'///////****CODIGO ANTIGUO*****///////
     CMO$ = TRIM$(CONTROL$("", "CODRECON"))
     If CMO$ = "" Then CMO$ = "RC"
     FIN& = ULTREG&("MOVISTO")
     For U& = FIN& To 1 Step -1
        Call TRACE(20, FIN& - U& + 1, FIN&)
        XX$ = REGLEIDO$("MOVISTO", U&)
        If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
           UNRO& = XVALO(Mid$(XX$, 26, 7))
              GoTo 210
        End If
     Next U&
     UNRO& = 0
     '
210  FECHDOC$ = HOS$
     NUMEDOC$ = TRIM$(Str$(1 + UNRO&))
     While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
     Wend
     NUMEDOC$ = CMO$ + "-" + NUMEDOC$
     REMI$ = NUMEDOC$
     TIPODOC$ = "Remito de Traslado"
     '
     DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
     DATRANSPO$ = DTP$
     '
     CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
     CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
     CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
     CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
     CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
     CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
     CAPARDOC% = 6
     '
     CODTABU1$(1) = "COD-MAT"
     CODTABU1$(2) = "DES-MAT"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "CAKILOS"
     CACOLTAB1% = 4
     '
     DESTIDOC% = (-1) * PROVEE%
     CAITAB1% = 0
     '
     For U& = 1 To CAITCRU%
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = CODEXT$(CCRU%(U&))
        TETABU1$(2, CAITAB1%) = DESCRIT$(CCRU%(U&))
        TETABU1$(3, CAITAB1%) = Str$(CACRU(U&))
        TETABU1$(4, CAITAB1%) = Str$(CACRU(U&) * PESUNI(CCRU%(U&)))
        '
        If REGMAC&(U&) > 0 Then
           ZZ$ = REGLEIDO$("MACONSIG", REGMAC&(U&))
           DOSEC$ = Mid$(ZZ$, 21, 12)
           TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CACRU(U&)
           Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)
           If HO% > CVI(Mid$(ZZ$, 71, 2)) Then
              Call REPLA(ZZ$, MKI$(HO%), 71, 2)
           End If
           Call GRAREG("MACONSIG", ZZ$, REGMAC&(U&))
           '
           LOTE$ = NOCOM$
           Call MOVISTOK(HO%, CCRU%(U&), LOTE$, CMO$, REMILA$, REMI$, DOSELA$, DOSEC$, RASOCLI$((-1) * PROVEE%), NUORIT%, 0, 0, (-1) * PROVEE%, 0, (-1) * CACRU(U&))
           '
        End If
219  Next U&
     '
     Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
     Call CIERRARCH("*.*")
     GoTo 50
     
End Sub

Function HAYDISTO%(CIX%, NORFAI&, CACONJU)
     '
     HDS% = 0
     If CHECOMPO% = 0 Then
        CHECOMPO% = (-1)
        If CONTROL$("", "CHECOMPO") = "SI" Then
          If CONTROL$("", "PRECOLAN") <> "SI" Then
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
        If CONTROL$("", "PRECOLAN") = "SI" Then
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
     SQLX$ = " SELECT * FROM ORDEFABR "
     SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "'"
     Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With ORDEFABTMP
       On Error Resume Next
       .MoveFirst
       If Err Then GoTo 1410
       On Error GoTo 0
       '
       CAPEN# = !CanProy - !CanApro
       If CAPEN# < 0 Then CAPEN# = 0
        CIXI% = !Cod_Inte
     End With
     '
1410 If CAPEN# < 0.1 Then
       Call MENSERR(24, "No hay Pendiente para la O/F '" + TRIM$(NORFA$) + "'.")
     End If
     '
     CICONJ% = CI%
     CODPRO$ = CODEXT$(CIXI%)
     CODICONJ$ = CODPRO$
     NORFABRI$ = NORFA$
     CACONJU# = CAPEN#
     VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
     '
1412 Call BLOQARCH("RESERVOF")
     Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 3)
     If VACONTROL% < 0 Then
        GoTo 1499
     End If
     '
     Call ABRESERVA
     SQLX$ = "DELETE * FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
     Stocks.Execute SQLX$
     '
1415 NORIT% = 0
     For U& = 1 To ULTREG&("!MATEROF")
        ZZ$ = REGLEIDO$("!MATEROF", U&)
        CIIX% = CVI(Left$(ZZ$, 2))
        CANTX# = CVS(Mid$(ZZ$, 5, 4))
        USOII = 0: If Abs(CAPEN#) >= 0.05 Then USOII = CANTX# / CAPEN#
        '
        If Abs(CANTX#) >= 0.05 Then
          With Reserva
            .AddNew
            !Cod_Inte = CIIX%
            !Cod_Exte = CODEXT$(CIIX%)
            !Descrip = DESCRIT0$(CIIX%)
            !IdSubOr = NORFA$
            !FechRes = Now
            !CantRes = CANTX#
            !CantCons = 0
            !Saldo_Entre = CANTX#
            !IdenLote = " "
            !Deposito = 0
            !UsoUnit = USOII
               NORIT% = NORIT% + 1
            !NuOrItem = NORIT%
            .Update
          End With
        End If
     Next U&
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
   CANOF% = 0: J& = 0
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("ORDEFABR")
   For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", i&)
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
               X$ = REGLEIDO$("ECORDEF", i&)
               J& = J& + 1
               Call GRAREG("ECORDEP", X$, J&)
            End If
         End If
      End If
   Next i&
   Call SETULTREG("RESERVA", 0)
   Call CIERRARCH("RESERVA")
   Call SETULTREG("ECORDEP", J&)
   Call CIERRARCH("ECORDEP")
   '
   FIN& = CANOF%
   J& = 0
   '
   For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      '
      For I1& = i& + 1 To FIN&                   ' ORDENAR
         If NOFACTI$(i&) > NOFACTI$(I1&) Then
            NOFAC$ = NOFACTI$(i&)
            NOFACTI$(i&) = NOFACTI$(I1&)
            NOFACTI$(I1&) = NOFAC$
         End If
      Next I1&
      '
      Call GENRESER(NOFACTI$(i&))
   Next i&
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
   NOFAN$ = ""
   With Reserva
     On Error Resume Next
     .MoveFirst
25   If Err < 1 Then
       NOFARES$ = !IdSubOr
       CARES# = !CantRes
       CACON# = !CantCons
       CASAL# = CARES# - CACON#
       .Edit
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
   Reserva.Close
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
     If EXISTE%("TRAZAB04.EXE") > 0 Then
        Call ANULARES(NORFA$)
        GoTo 130
     End If
     '
120  Call ABRESERVA
     With Reserva
125    .FindFirst "IdSubOr = '" & NORFA$ & "'"
       If .NoMatch = False Then
         .Delete
         GoTo 125
       End If
     End With
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
    J& = 0
    '
    Call ABREORFAB
    '
    SQLX$ = " SELECT DISTINCT IdSubOr, StatuOrf, CanApro, CanProy FROM ORDEFABR "
    If MODOI% = 2 Then
        ' SQLX$ = SQLX$ + " WHERE StatuOrf =1 "
        SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "   ' para que no este anulada ni cumplida
        SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
      Else
        SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
    End If
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
        Call REPLA(Y$, !IdSubOr, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IdSubOr), 13, 52)
        '
        JJ& = JJ& + 1
        Call GRAREG("ECORDEP", Y$, JJ&)
        '
      .MoveNext
      Loop
    End With
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
    For i& = 1 To CAEQUI&
      XX$ = REGLEIDO$("PADMAQ", i&)
      EQQ$(i&) = Left$(XX$, 6)
    Next i&
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
          For J% = 1 To CAEQUI&
            If TRIM$(EQQ$(J%)) <> "" Then
              If EQQ$(J%) = EQ1$ Then
                 TPREP(J%) = TPREP(J%) + MINUPREP
                 TOPER(J%) = TOPER(J%) + MINLOTE
                ElseIf EQQ$(J%) = EQ2$ Then
                 TPREP(J%) = TPREP(J%) + MINUPREP
                 TOPER(J%) = TOPER(J%) + MINLOTE
                ElseIf EQQ$(J%) = EQ3$ Then
                 TPREP(J%) = TPREP(J%) + MINUPREP
                 TOPER(J%) = TOPER(J%) + MINLOTE
                ElseIf EQQ$(J%) = EQ4$ Then
                 TPREP(J%) = TPREP(J%) + MINUPREP
                 TOPER(J%) = TOPER(J%) + MINLOTE
              End If
            End If
          Next J%
        End If
      End If
    Next TTL&
    '
    For i& = 1 To CAEQUI&
      XX$ = REGLEIDO$("PADMAQ", i&)
      Call REPLA(XX$, MKS$(60 * TPREP(i&)), 81, 4)
      Call REPLA(XX$, MKS$(60 * TOPER(i&)), 85, 4)
      Call REPLA(XX$, MKS$(60 * (TPREP(i&) + TOPER(i&))), 89, 4)
      Call GRAREG("PADMAQ", XX$, i&)
    Next i&
    '
    Call CIERRARCH("PADMAQ")
    Screen.MousePointer = 1
    '
    Call FRESHALL
    Call FINAL("*TOTIEPRO")
    '
End Sub

Sub VERIDEST()
    '
    STO1% = 0: STO2% = 0
    For i& = 1 To ULTREG&("DESTFAB")
       X$ = REGLEIDO$("DESTFAB", i&)
       CODEST$ = TRIM$(UCase$(Left$(X$, 8)))
       If CODEST$ = "STOCK" Then STO1% = 1
       If CODEST$ = "SEMIELA" Then STO2% = 1
    Next i&
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
   J& = 0
   FIN& = ULTREG&("ORFANUE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORFANUE", U&)
      If Asc(Mid$(XX$, 112, 1)) = 1 Then
         J& = J& + 1
         NUORF(J&) = CVS(Left$(XX$, 4))
         REORF&(J&) = U&
      End If
   Next U&
   '
   FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         NOFA = CVS(Left$(XX$, 4))
         For K% = 1 To J&
            If NUORF(K%) = NOFA Then
               STORF%(K%) = 1
            End If
         Next K%
      End If
   Next U&
   '
   For U& = 1 To J&
      Call TRACE(20, U&, J&)
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
    Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
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
        CAPROG# = !CantProg
        CAREA# = !CantReal
        CAPEN# = CAPROG# - CAREA#
        CIXII% = !Cod_Inte
        NOPEII% = !NumeOper
        '
        If CAPEN# >= 0.05 Then
          If CAPROG# >= 0.05 Then
            COEF = CAPEN# / CAPROG#
            If COEF > 1 Then COEF = 1
            TPREP = !TiemPrep: If CAREA# >= 0.05 Then TPREP = 0
            TFIJO = !TiemFijo
            'TVARI = TIESTAN(CIXII%, NOPEII%, CAPEN#, 1) 'COEF * CVS(Mid$(O$, 111, 4))
            TPROD = TIEMAQUI(CIXII%, NOPEII%, CAPEN#, 1) ' MODO 1 PARA NO INCLUIR PREPARACION
            TVARI = TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
            TDEMO = 0: TMANT = 0
              If (!TiemFijo + !TiemVari) > 0.1 Then
                TDEMO = TVARI * !TiemDemo / (!TiemFijo + !TiemVari)
              End If
            TMANT = COEF * CVS(Mid$(O$, 119, 4))
            TLIMP = CVS(Mid$(O$, 123, 4))
            TOTRO = COEF * CVS(Mid$(O$, 127, 4))
            TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
            '
            YY$ = REGBLAN("OPERFAP")
            Call REPLA(YY$, !IdSubOr, 1, 12)
            Call REPLA(YY$, MKI$(!Cod_Inte), 13, 2)
            Call REPLA(YY$, MKI$(!NumeOper), 15, 2)
            Call REPLA(YY$, !DescOper, 17, 48)
            Call REPLA(YY$, !Destino, 65, 8)
            Call REPLA(YY$, !Maquina, 73, 6)
            Call REPLA(YY$, MKD$(!CantProg), 79, 8)
            Call REPLA(YY$, MKD$(!CantReal), 87, 8)
            Call REPLA(YY$, MKD$(!CantPend), 95, 8)
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
            Call REPLA(YY$, !Observa, 151, 48)
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
    '
    Set OPERFAPTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
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
        CAPROG# = !CantProg
        CAREA# = !CantReal
        CAPEN# = CAPROG# - CAREA#
        CIXII% = !Cod_Inte
        NOPEII% = !NumeOper
        '
        If CAPEN# >= 0.05 Then
          If CAPROG# >= 0.05 Then
            COEF = CAPEN# / CAPROG#
            If COEF > 1 Then COEF = 1
            TPREP = !TiemPrep: If CAREA# >= 0.05 Then TPREP = 0
            TFIJO = !TiemFijo
            'TVARI = TIESTAN(CIXII%, NOPEII%, CAPEN#, 1) 'COEF * CVS(Mid$(O$, 111, 4))
            TPROD = TIEMAQUI(CIXII%, NOPEII%, CAPEN#, 1) ' MODO 1 PARA NO INCLUIR PREPARACION
            TVARI = TPROD - TFIJO: If TVARI < 0 Then TVARI = 0
            TDEMO = 0: TMANT = 0
              If (!TiemFijo + !TiemVari) > 0.1 Then
                TDEMO = TVARI * !TiemDemo / (!TiemFijo + !TiemVari)
              End If
            TMANT = COEF * CVS(Mid$(O$, 119, 4))
            TLIMP = CVS(Mid$(O$, 123, 4))
            TOTRO = COEF * CVS(Mid$(O$, 127, 4))
            TTOTA = TPREP + TFIJO + TVARI + TDEMO + TMANT + TLIMP + TOTRO
            '
            YY$ = REGBLAN("OPERFAB")
            Call REPLA(YY$, !IdSubOr, 1, 12)
            Call REPLA(YY$, MKI$(!Cod_Inte), 13, 2)
            Call REPLA(YY$, MKI$(!NumeOper), 15, 2)
            Call REPLA(YY$, !DescOper, 17, 48)
            'Call REPLA(YY$, !Destino, 65, 8)
            Call REPLA(YY$, !Maquina, 73, 6)
            Call REPLA(YY$, MKD$(!CantProg), 79, 8)
            Call REPLA(YY$, MKD$(!CantReal), 87, 8)
            Call REPLA(YY$, MKD$(!CantPend), 95, 8)
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
            Call REPLA(YY$, !Observa, 151, 48)
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
Sub GENORDEFDAT(MODO%)
  '
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM ORDEFABR "
  If MODO% = 1 Then SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
  SQLX$ = SQLX$ + " ORDER BY IdSubOr "
  '
  Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With ORDEFABTMP
    On Error Resume Next
    .MoveFirst
    If Err Then Call MENSERR(24, "No Existen Ordenes de Fabricación en Curso.\Imposible continuar la carga")
    On Error GoTo 0
    JJ& = 0: J2& = 0
    Do While (.EOF = False)
      Y$ = REGBLAN("ORDEFABR")
      '
      Call REPLA(Y$, MKS$(!NUMEORFA), 1, 4)
      Call REPLA(Y$, !Referen, 5, 64)
      Call REPLA(Y$, !IdSubOr, 69, 12)
      Call REPLA(Y$, MKI$(CVINT(!FechGen)), 81, 2)
      Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
      Call REPLA(Y$, !Unidad, 85, 4)
      Call REPLA(Y$, MKD$(!CanProy), 89, 8)
      '
      On Error Resume Next
      Call REPLA(Y$, !Destino, 97, 8)
      On Error GoTo 0
      '
      Call REPLA(Y$, MKI$(CVINT(!FechEnSol)), 105, 2)
      Call REPLA(Y$, Chr$(!Priorid), 107, 1)
      Call REPLA(Y$, Chr$(!Statuorf), 108, 1)
      Call REPLA(Y$, MKI$(CVINT(!FechLan)), 109, 2)
      Call REPLA(Y$, MKS$(!SecciPro), 113, 4)
      Call REPLA(Y$, MKI$(CVINT(!FechCierr)), 117, 2)
      Call REPLA(Y$, MKD$(!CanApro), 119, 8)
      Call REPLA(Y$, MKD$(!CanRech), 127, 8)
      Call REPLA(Y$, !IdenRetra, 135, 8)
      Call REPLA(Y$, MKI$(CVINT(!FechProyEntre)), 147, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechTempEntre)), 149, 2)
      Call REPLA(Y$, MKI$(CVINT(!FechRecal)), 151, 2)
      Call REPLA(Y$, !Observa, 155, 64)
      '
      JJ& = JJ& + 1
      Call GRAREG("ORDEFABR", Y$, JJ&)
      '
      If !Statuorf < 2 Then
        J2& = J2& + 1
        Call GRAREG("ORDEFABP", Y$, J2&)
      End If
    .MoveNext
    Loop
  End With
  Call SETULTREG("ORDEFABR", JJ&)
  Call CIERRARCH("ORDEFABR")
  Call SETULTREG("ORDEFABP", J2&)
  Call CIERRARCH("ORDEFABP")
End Sub

Sub CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, MODOXX%)
    '
    ' VACONTROL% = 1  presenta componentes para edicion
    '              0  no presenta componentes
    '             -1  devuelto como señal de cancelacion
    '
    Screen.MousePointer = 11
    Call BLANARCH("!MATEROF")
    CAIT% = 0: COEF = 1
    Dim RERES&(1024)
    '
    VERISARRAS% = 0
    If CONTROL$("", "CHECOMPO") = "SI" Then
      VERISARRAS% = 1
    End If
    '
    CIXI% = CODINT%(CODICONJ$)
    If CIXI% < 1 Then
       Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Código de Artículo '" + TRIM$(CODICONJ$) + "' Inexistente\Consulte a su Soporte de Sistemas.")
       Call FINAL("")
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
    SQLX$ = " SELECT Cod_Inte, UsoUnit FROM RESERVA "
    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + "ORDER BY NuOrItem ASC"
    '
    Set RESERTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then GoTo 200   ' No Existen Reservas efectuadas.
      On Error GoTo 0
      CAIT% = 0: JJ& = 0
      Do While (.EOF = False)
        If !Cod_Inte <> CIJ%(CAIT%) Or !UsoUnit <> USOI(CAIT%) Then
          CAIT% = CAIT% + 1
          CIJ%(CAIT%) = !Cod_Inte
          USOI(CAIT%) = !UsoUnit
        End If
      .MoveNext
      Loop
    End With
    On Error Resume Next
    RESERTMP.Close
    Set RESERTMP = Nothing
    On Error GoTo 0
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
        Call VESARRGLO2
      End If
    End If
    '
    FIN& = CAIT%
    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
         If Abs(USOI(I1%)) > 0 Then
           '
           ZZ$ = REGBLAN$("MATEROF")
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
                Else                       ' CASO DE GENERACION DE UNA O/F
                  STDIS = STOCKDIS(CIJ%(I1%))
             End If
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
             FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
    Next I1%
    Call CIERRARCH("!MATEROF")
    Screen.MousePointer = 1
    '
    If VACONTROL% = 1 Then
       TITUPA$ = TITUPA$ + ": "
       If TRIM$(NORFABRI$) <> "" Then TITUPA$ = TITUPA$ + TRIM$(NORFABRI$)
       If TRIM$(CODPRO$) <> "" Then TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
       '
       VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
       '
       If VTB% < 0 Then
            VACONTROL% = (-1)
          Else
            If Abs(CACONJU#) >= 0.05 Then
               For i& = 1 To ULTREG&("!MATEROF")
                  ZZ$ = REGLEIDO$("!MATEROF", i&)
                  CANUE# = CDbl(CVS(Mid$(ZZ$, 5, 4)))
                  USNUE# = CANUE# / CACONJU#
                  Call REPLA(ZZ$, MKD$(USNUE#), 9, 8)
                  Call GRAREG("!MATEROF", ZZ$, i&)
               Next i&
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
   Set INFOFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   CANREALI = 0
   With INFOFABTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         If IsNull(!ReTraba) Or !ReTraba < 1 Then
           CANREALI = CANREALI + !CantReal
         End If
       .MoveNext
       Loop
     End If
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
   Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
   
   STATORFA% = 99 ' no existe
   On Error Resume Next
   ORDEFABTMP.MoveFirst
   If Err < 1 Then
      STATORFA% = ORDEFABTMP!Statuorf
   End If
   '
   NORFAN$ = NUMORFA$
   STATAN% = STATORFA%
   '
End Function
