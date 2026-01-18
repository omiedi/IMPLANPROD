Attribute VB_Name = "TABLAS_1"
Sub CARTADEPOSI()
    '
    
    If ULTREG&("TADEPOSI") < 1 Then
      Call BLANARCH("!TADEPOSI")
      J& = 0
      For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
          If CODEPO% <> 99 Then
            If J& < 62 Then
              J& = J& + 1
              Y$ = REGBLAN$("TADEPOSI")
              Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
              Call REPLA(Y$, Left$(X$, 63), 2, 63)
              Call REGAPP("TADEPOSI", Y$)
            End If
          End If
        End If
      Next U&
      Call CIERRARCH("SUCURSAL")
    End If
    '
    Call TRALOCAL("TADEPOSI", "")
    For U& = 1 To ULTREG&("TADEPOSI")
      XX$ = REGLEIDO$("!TADEPOSI", U&)
      If ECOARCH$("PROVED1", Mid$(XX$, 61, 2)) = "" Then
         Call REPLA(XX$, MKI$(0), 61, 2)
         Call GRAREG("!TADEPOSI", XX$, U&)
      End If
    Next U&
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, J&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, J&)
               Call GRAREG("TADEPOSI", X$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", Y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    Call FRESHECHO("SECPROD")
    '
End Sub

Sub CARTATIMA()
    '
    Call BLANARCH("!TABTIMA")
    '
    For i& = 1 To ULTREG&("TABTIMA")
       '
       X$ = REGLEIDO$("TABTIMA", i&)
       '
       If Asc(Left$(X$, 1)) > 0 Then
           Call REGAPP("!TABTIMA", X$)
       End If
    Next i&
    Call CIERRARCH("!TABTIMA")
    '
10  VTB% = VENTABU%("TABTIMA")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!TABTIMA")
       X$ = REGLEIDO$("!TABTIMA", i&)
       If Asc(Left$(X$, 1)) > 0 Then
          DPP$ = Mid$(X$, 17, 1)
          If ECOARCH$("ECODEP", DPP$) = "" Then
             Call MENSERR(24, "Depósito no Definido")
             GoTo 10
          End If
          If Asc(DPP$) = 99 Then
             Call MENSERR(24, "Depósito 99 no Válido")
             GoTo 10
          End If
          J& = J& + 1
          Call GRAREG("TABTIMA", X$, J&)
          Call GRAREG("TATIMAT", X$, J&)
       End If
    Next i&
    Call SETULTREG("TABTIMA", J&)
    Call CIERRARCH("TABTIMA")
    Call SETULTREG("TATIMAT", J&)
    Call CIERRARCH("TATIMAT")
    Call CIERRARCH("!TABTIMA")
    '
    ' SI EXISTE SE GRABA EL TABTIMB
    If EXISTE(LUDAT$ + "\TABTIMB.RFS") > 0 Then
       Call BLANARCH("TABTIMB")
       'COPIA DATOS
       For i& = 1 To ULTREG&("TABTIMA")
         X$ = REGLEIDO$("TABTIMA", i&)
         A$ = ""
         Call REPLA(A$, Mid(X$, 1, 1), 1, 1)
         Des$ = Mid(X$, 18, 64)
         If TRIM$(Des$) = "" Then
            Des$ = Mid(X$, 2, 15)
         End If
         Call REPLA(A$, Des$, 2, 64)
         Call REGAPP("TABTIMB", A$)
       Next i&
       '
       Call SETULTREG("TABTIMB", i&)
       Call CIERRARCH("TABTIMB")
       '
    End If
    '
    'DESDE AQUI AGREGA A TABLA DATASQL LOS DATOS DE TIPO DE MATERIAL
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    
    IDLISTAX& = XVALO(VALOBADA("DATASQL", "ID_LISTA", "LISTA = 'TABTIMA'", "NOMESS"))
    YAEXISTE% = 0: If IDLISTAX& > 0 Then YAEXISTE% = 1
    If IDLISTAX& < 1 Then IDLISTAX& = 1 + XVALO(VALOBADA("DATASQL", "MAX(ID_LISTA)", "ID_LISTA > 0", "NOMESS"))
    
    IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    
    Call BORRAREGISTROS("DATASQL", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    For i& = 1 To ULTREG&("TABTIMA")
       '
       X$ = REGLEIDO$("TABTIMA", i&)
       NUMX = Asc(Left(X$, 1))
       REF$ = TRIM$(Mid$(X$, 2))
       If Asc(Left$(X$, 1)) > 0 Then
            SQLX$ = "INSERT INTO DATASQL ("
            SQLX$ = SQLX$ + "  CodiEmpr,ID_LISTA,LISTA,CAMPO1,CAMPO2)"
            SQLX$ = SQLX$ + " VALUES(" & CodiEmp%
            SQLX$ = SQLX$ + " ," & IDLISTAX&
            SQLX$ = SQLX$ + " ,'TABTIMA'"
            SQLX$ = SQLX$ + " ," & NUMX
            SQLX$ = SQLX$ + " ,'" & REF$ & "'"
            SQLX$ = SQLX$ + ")"
            FLEXCONN.Execute (SQLX$)
       End If
    Next i&
    
    If YAEXISTE% < 1 Then
        'SI NO EXISTIA EN LA BASE AGREGA TAMBIEN EL ID EN ENCAPRFUN PARA QUE NO SE PISE CON LA CREACION DE OTRA LISTA
        SQLX$ = "INSERT INTO ENCAPRFUN ("
        SQLX$ = SQLX$ + "  CodiEmpr,ID_LISTA,LISTA)"
        SQLX$ = SQLX$ + " VALUES(" & CodiEmp%
        SQLX$ = SQLX$ + " ," & IDLISTAX&
        SQLX$ = SQLX$ + " ,'TABTIMA'"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
    End If

ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
    End If
    
    '
End Sub
'
Sub CARUBIC(CIII%)
    '
    Dim REGUBI&(512)
    If CIII% <= 0 Then
        Call MENSERR(24, "Código de Artículo no Válido")
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("!CARUBIC")
    '
    
    K1& = 0
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CIII%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) = CIII% Then
            If TRIM$(Mid$(X$, 3, 24)) <> "" Then
                Y$ = Left$(X$, 60) + MKS$(U&)
                Call REGAPP("!CARUBIC", Y$)
                K1& = K1& + 1: REGUBI&(K1&) = REG&
            End If
        End If
    Next U&
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CARUBIC/BORDEIZQ=20/BORDESUP=12/NC/TITUPAN=" + TRIM$(CODEXT$(CIII%)) + " - " + DESCRIT$(CIII%))
    '
    If VTB% >= 0 Then
        Screen.MousePointer = 11
        k& = ULTREG&("!CARUBIC")
        K2& = 0
        For i& = 1 To k&
            X$ = REGLEIDO$("!CARUBIC", i&)
            If TRIM$(Mid$(X$, 3, 24)) <> "" Then
                Y$ = MKI$(CIII%) + Mid$(X$, 3, 58) + MKS$(0)
                K2& = K2& + 1
                REG& = REGUBI&(K2&)
                If REG& > 0 Then
                      Call GRAREG("UBISTOCK", Y$, REG&)
                    Else
                      Call REGAPP("UBISTOCK", Y$)
                End If
            End If
        Next i&
        '
        For k& = K2& + 1 To K1&
            REG& = REGUBI&(k&)
            Y$ = MKI$(CIII%) + Space$(58) + MKS$(0)
            Call GRAREG("UBISTOCK", Y$, REG&)
        Next k&
        '
    End If
    '
    Call CIERRARCH("UBISTOCK")
    Screen.MousePointer = 1
    '
End Sub
'



