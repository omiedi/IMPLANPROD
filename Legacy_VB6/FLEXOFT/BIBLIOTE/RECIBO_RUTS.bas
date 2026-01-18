Attribute VB_Name = "RECIBO_RUTS"
Sub CARTAMECO()     ' carga tabla medios de cobranza
    '
    LU$ = LUDAT$
    If EXISTE%(LU$ + "MECOBRA.DAT") = 1 Then
       If COPYFILE%(LU$ + "MECOBRA.DAT", "C:\FLEXOFT\MECOBRA.WKF") <> 1 Then
          command1.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COMECOB/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!MECOBRA")
       X$ = REGLEIDO$("!MECOBRA", U&)
       CDX% = CVI(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 999 Then
             J& = J& + 1
363          Call GRAREG("!MECOBRA", X$, J&)
             If TRIM$(Mid$(X$, 3, 12)) = "" Then ATX% = 1
             If TRIM$(Mid$(X$, 45, 12)) = "" Then ATX% = 1
364          If TRIM$(Mid$(X$, 57, 4)) = "" Then ATX% = 1
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!MECOBRA", J&)
    '
366 Screen.MousePointer = 1
    If ATX% > 0 Then
        Call COMUNI("Datos Incompletos !!!")
        GoTo 360
    End If
    '
    AA% = COPYFILE%("C:\FLEXOFT\MECOBRA.WKF", LU$ + "MECOBRA.DAT")
    Call FINAL("*SOMECO1")
    '
390 End Sub

Sub CARTACOBRA()    ' carga tabla cobradores / talonarios
    '
    Dim MCOBRA%(1000)
    '
    Call CIERRARCH("COBRADOR")
    Call CIERRARCH("!COBRADOR")
    LU$ = LUDAT$
    If EXISTE%(LU$ + "COBRADOR.DAT") = 1 Then
       If COPYFILE%(LU$ + "COBRADOR.DAT", "C:\FLEXOFT\COBRADOR.WKF") <> 1 Then
          Command6.Enabled = True
          Exit Sub
       End If
    End If
    '
360 ATY% = VENTABU%("COBRADOR/NC")
    If ATY% < 0 Then GoTo 390
    '
361 ATX% = 0: J& = 0
    Screen.MousePointer = 11
    For U& = 1 To ULTREG&("!COBRADOR")
       X$ = REGLEIDO$("!COBRADOR", U&)
       CDX% = Asc(Left$(X$, 2))
362    If CDX% > 0 Then
          If CDX% <= 255 Then
             If MCOBRA%(CDX%) = 0 Then
                J& = J& + 1
363             Call GRAREG("!COBRADOR", X$, J&)
                MCOBRA%(CDX%) = 1
             End If
          End If
       End If
365 Next U&
    '
    Call SETULTREG("!COBRADOR", J&)
    Call CIERRARCH("!COBRADOR")
    Call CIERRARCH("COBRADOR")
    '
366 Screen.MousePointer = 1
    '
    AA% = COPYFILE%("C:\FLEXOFT\COBRADOR.WKF", LU$ + "COBRADOR.DAT")
    '
390 End Sub
