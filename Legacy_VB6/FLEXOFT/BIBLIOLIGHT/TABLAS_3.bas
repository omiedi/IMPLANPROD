Attribute VB_Name = "TABLAS_3"
Sub CARGTAREPCOM()

    OPX% = COMALTER%("Elija Opción \\Modificar\Nuevo")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    If OPX% = 1 Then
     'MODIFICACION
     Call SELECHO("REPRCOME")
     NUM% = XVALO(VALACT1$("REPRCOME"))
     If NUM% > 0 Then
       If NUM% < 32767 Then
         If ECOARCH$("REPRCOME", MKI$(NUM%)) <> "" Then
           VDEF$ = FILTERGETRECORD$("REPRCOME", 1, MKI$(NUM%))
           
10         OBX$ = OBJPLA$("REPRESCO", VDEF$)
           If OBX$ <> "" Then
             If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
               Call MENSERR(24, "Falta Nombre Representante")
               VDEF$ = OBX$
               GoTo 10
             End If
             '
             Call BLOQARCH("REPRCOME")
             Call FILTERPUTRECORD("REPRCOME", 1, MKI$(NUM%), OBX$)
             Call CIERRARCH("REPRCOME")
           End If
           Exit Sub
         End If
       End If
     End If
    End If
 
    If OPX% = 2 Then
       'NUEVO
       VDEF$ = REGBLAN$("REPRCOME")
101    PROXI% = 1 + ULTREG&("REPRCOME")
11     RFF$ = RECFILT$("REPRCOME", 1, MKI$(PROXI%))
       If RFF$ <> "" Then
         PROXI% = PROXI% + 1
         GoTo 11
       End If
       '
       Call REPLA(VDEF$, MKI$(PROXI%), 1, 2)
       Call REPLA(VDEF$, Chr$(0), 194, 1) ' IVA EN 0 POR DEFECTO
       '
       OBX$ = OBJPLA$("REPRESCO", VDEF$)
       If OBX$ <> "" Then
         If TRIM$(Mid$(OBX$, 3, 32)) = "" Then
           Call MENSERR(24, "Falta Nombre Representante")
           VDEF$ = OBX$
           GoTo 101
         End If
         '
         Call BLOQARCH("REPRCOME")
21       RFF$ = RECFILT$("REPRCOME", 1, MKI$(PROXI%))
         If RFF$ <> "" Then
           PROXI% = PROXI% + 1
           GoTo 21
         End If
         '
         Call REPLA(OBX$, MKI$(PROXI%), 1, 2)
         Call REGAPP("REPRCOME", OBX$)
         Call CIERRARCH("REPRCOME")
       End If
       '
     End If

End Sub

Sub CARGDIRECOM()
    J& = 0
    For U& = 1 To ULTREG&("AUTORIZA")
        X$ = REGLEIDO$("AUTORIZA", U&)
        CODAUTO% = CVI(Left$(X$, 2))
        If CODAUTO% > 0 Then
            J& = J& + 1
            Call GRAREG("!AUTORIZA", X$, J&)
        End If
    Next U&
    Call SETULTREG("!AUTORIZA", J&)
    Call CIERRARCH("AUTORIZA")
    Call CIERRARCH("!AUTORIZA")
    '
    VTB% = VENTABU%("AUTORIZA")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!AUTORIZA")
        X$ = REGLEIDO$("!AUTORIZA", U&)
        CODAUTO% = CVI(Left$(X$, 2))
        If CODAUTO% > 0 Then
            J& = J& + 1
            Call GRAREG("AUTORIZA", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("AUTORIZA", J&)
    Call CIERRARCH("AUTORIZA")
    Call CIERRARCH("!AUTORIZA")
End Sub

Sub CARGFUNILVE()
    J& = 0
    For U& = 1 To ULTREG&("FUNIVTAS")
        X$ = REGLEIDO$("FUNIVTAS", U&)
        CODFUN% = CVI(Left$(X$, 2))
        If CODFUN% > 0 Then
            J& = J& + 1
            Call GRAREG("!FUNIVTAS", X$, J&)
        End If
    Next U&
    Call SETULTREG("!FUNIVTAS", J&)
    Call CIERRARCH("FUNIVTAS")
    Call CIERRARCH("!FUNIVTAS")
    '
    VTB% = VENTABU%("FUNIVTAS")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!FUNIVTAS")
        X$ = REGLEIDO$("!FUNIVTAS", U&)
        CODFUN% = CVI(Left$(X$, 2))
        If CODFUN% > 0 Then
            J& = J& + 1
            Call GRAREG("FUNIVTAS", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("FUNIVTAS", J&)
    Call CIERRARCH("FUNIVTAS")
    Call CIERRARCH("!FUNIVTAS")


End Sub

Sub CARGTATIPOP()

    J& = 0
    For U& = 1 To ULTREG&("TIPOOPOR")
        X$ = REGLEIDO$("TIPOOPOR", U&)
        CODTIP% = CVI(Left$(X$, 2))
        If CODTIP% > 0 Then
            J& = J& + 1
            Call GRAREG("!TIPOOPOR", X$, J&)
        End If
    Next U&
    Call SETULTREG("!TIPOOPOR", J&)
    Call CIERRARCH("TIPOOPOR")
    Call CIERRARCH("!TIPOOPOR")
    '
    VTB% = VENTABU%("TIPOOPOR")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TIPOOPOR")
        X$ = REGLEIDO$("!TIPOOPOR", U&)
        CODTIP% = CVI(Left$(X$, 2))
        If CODTIP% > 0 Then
            J& = J& + 1
            Call GRAREG("TIPOOPOR", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("TIPOOPOR", J&)
    Call CIERRARCH("TIPOOPOR")
    Call CIERRARCH("!TIPOOPOR")

End Sub
