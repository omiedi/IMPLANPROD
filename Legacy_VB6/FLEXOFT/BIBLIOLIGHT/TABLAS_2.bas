Attribute VB_Name = "TABLAS_2"

Sub CARCONPAGOS()
    '
    Screen.MousePointer = 11
    '\\\OT-06-0221-WAR-02/06/06///
    OPT1% = COMALTER%("Seleccione una Opción:\\Actualizar Tabla\Listar")
    If OPT1% = 1 Then
      '
      Dim CPAX%(512)
      Call GRAREG("TAFEREF", Chr$(1) + "F.Factura", 1)
      Call GRAREG("TAFEREF", Chr$(2) + "Sem.Fact.", 2)
      Call GRAREG("TAFEREF", Chr$(3) + "Quin.Fact", 3)
      Call GRAREG("TAFEREF", Chr$(4) + "Mes Fact.", 4)
      Call SETULTREG("TAFEREF", 4)
      Call CIERRARCH("TAFEREF")
      Call TRALOCAL("CONPAG", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("CONPAG")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!CONPAG")
         X$ = REGLEIDO$("!CONPAG", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Condición Número '0' no Válida.")
            GoTo 10
         End If
         For v& = 1 To U& - 1
            If COCO% = CPAX%(v&) Then
               Call MENSERR(24, "Condición Número '" + TRIM$(Str$(COCO%)) + "' Repetida.")
               GoTo 10
            End If
         Next v&
         CPAX%(U&) = COCO%
         '
         DECO$ = TRIM$(Mid$(X$, 2, 35))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción de la\Condición Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
         FREF% = Asc(Mid$(X$, 37, 1))
         If FREF% < 1 Or FREF% > 4 Then
            Call MENSERR(24, "Código de Fecha de Referencia no Válido\en Condición Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("CONPAG", "")
      Call FRESHECHO("CONPAG")
    ElseIf OPT1% = 2 Then
       ZZ$ = REGBLAN$("CONPAG")
       JJ& = 0
       For U& = 1 To ULTREG&("CONPAG")
         X$ = REGLEIDO$("CONPAG", U&)
         '
         COCO% = Asc(Left$(X$, 1))
         DECO$ = TRIM$(Mid$(X$, 2, 35))
         FREF% = Asc(Mid$(X$, 37, 1))
         P1% = Asc(Mid$(X$, 38, 1))
         P2% = Asc(Mid$(X$, 39, 1))
         P3% = Asc(Mid$(X$, 40, 1))
         P4% = Asc(Mid$(X$, 41, 1))
         P5% = Asc(Mid$(X$, 42, 1))
         P6% = Asc(Mid$(X$, 43, 1))
         '
         Y$ = ZZ$
         Call REPLA(Y$, Chr$(COCO%), 1, 1)
         Call REPLA(Y$, DECO$, 2, 35)
         Call REPLA(Y$, Chr$(FREF%), 37, 1)
         Call REPLA(Y$, Chr$(P1%), 38, 1)
         Call REPLA(Y$, Chr$(P2%), 39, 1)
         Call REPLA(Y$, Chr$(P3%), 40, 1)
         Call REPLA(Y$, Chr$(P4%), 41, 1)
         Call REPLA(Y$, Chr$(P5%), 42, 1)
         Call REPLA(Y$, Chr$(P6%), 43, 1)
         JJ& = JJ& + 1
         Call GRAREG("LICONPAG", Y$, JJ&)
         '
      Next U&
      Call SETULTREG("LICONPAG", JJ&)
      Call FINAL("*LICONPAG")
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Sub CARGRUCOVE()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    Call BLANARCH("!GRUCOVEN")
    For i& = 1 To ULTREG&("GRUCOVEN")
        X$ = REGLEIDO$("GRUCOVEN", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!GRUCOVEN", X$)
        End If
    Next i&
    Call CIERRARCH("GRUCOVEN")
    '
    VTB% = VENTABU%("GRUCOVEN/NC/TITUPAN=Grupos Contables")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!GRUCOVEN")
        X$ = REGLEIDO$("!GRUCOVEN", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            J& = J& + 1
            Call GRAREG("GRUCOVEN", X$, J&)
        End If
    Next i&
    Call SETULTREG("GRUCOVEN", J&)
    Call CIERRARCH("GRUCOVEN")
    
    If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
       Call CREATABLA_GRUCOVENSQL
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       SQLX$ = "DELETE FROM GRUCOVENSQL"
       FLEXCONN.Execute (SQLX$)
       '--------------------

        For i& = 1 To ULTREG&("!GRUCOVEN")
            X$ = REGLEIDO$("!GRUCOVEN", i&)
            GRCONTA% = CVI(Left$(X$, 2))
            If GRCONTA% > 0 Then
               REFERENCIA$ = ECOARCH$("GRUCOVEN", MKI(GRCONTA%))
               sqlx2$ = "INSERT INTO GRUCOVENSQL"
               sqlx2$ = sqlx2$ + " (CODIGO,DESCRIPCION)"
               sqlx2$ = sqlx2$ + " VALUES(" & GRCONTA%
               sqlx2$ = sqlx2$ + " ,'" & REFERENCIA$ & "'"
               sqlx2$ = sqlx2$ + ")"
               FLEXCONN.Execute (sqlx2$)
            End If
        Next i&
        Call CIERRARCH("GRUCOVEN")

       
ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans
           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           Exit Sub
        Else
          FLEXCONN.CommitTrans
          Call COMUNI("Grabación Realizado Exitosamente")
        End If
    End If
    '
End Sub

Sub CARTAPROVIN()
    '
    If ULTREG&("PROVINB") < 1 Then
      For i& = 1 To ULTREG&("PROVINC")
        X$ = REGLEIDO$("PROVINC", i&) + String$(32, 0)
        Call GRAREG("PROVINB", X$, i&)
      Next i&
      Call CIERRARCH("PROVINC")
      Call CIERRARCH("PROVINB")
    End If
    '

    Call BLANARCH("!PROVINB")
    For i& = 1 To ULTREG&("PROVINB")
        X$ = REGLEIDO$("PROVINB", i&)
        If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
            Call REPLA(X$, MKI$(0), 27, 2)
        End If
        If TRIM$(Left$(X$, 2)) <> "" Then
            Call REGAPP("!PROVINB", X$)
        End If
    Next i&
    Call CIERRARCH("PROVINB")
    '
10  VTB% = VENTABU%("PROVINC/NC/TITUPAN=Tabla de Provincias - Zonas de Venta")
    If VTB% < 0 Then Exit Sub
    
    For i& = 1 To ULTREG&("!PROVINB")
        X$ = REGLEIDO$("!PROVINB", i&)
        RETIB% = CVI(Mid$(X$, 29, 2))
        ConvMultLat = CVI(Mid$(X$, 37, 2))
        Provincia$ = TRIM$(Mid$(X$, 2, 24))
        If RETIB% > 0 Then
          CUES$ = CODCUE$(RETIB%)
          If CUES$ = "" Then
             Call COMUNI("Cuenta Contable de Retención No Válida\Para la Provincia de '" & Provincia$ & "'")
             GoTo 10
          End If
          If ConvMultLat < 1 Then
             Call COMUNI("Falta Código de Jurisdicción de Convenio Multilateral\Para la Provincia de '" & Provincia$ & "'")
             GoTo 10
          End If
        End If
    Next i&
    
    J& = 0
    For i& = 1 To ULTREG&("!PROVINB")
        X$ = REGLEIDO$("!PROVINB", i&)
        If TRIM$(Left$(X$, 1)) <> "" Then
            J& = J& + 1
            Call GRAREG("PROVINB", X$, J&)
            Call GRAREG("PROVINC", Left$(X$, 39) + Space$(6), J&)
        End If
    Next i&
    Call SETULTREG("PROVINB", J&)
    Call CIERRARCH("PROVINB")
    Call SETULTREG("PROVINC", J&)
    Call CIERRARCH("PROVINC")
    '
End Sub

Sub CARTASIVA()
   Call TRALOCAL("GRUPIVA", "")
   VTB% = VENTABU%("GRUPIVA/NC/TITUPAN=Tasa IVA por Categoria")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!GRUPIVA")
         X$ = REGLEIDO$("!GRUPIVA", U&)
         Call REPLA(X$, Chr$(U& - 1), 1, 1)
         Call GRAREG("GRUPIVA", X$, U&)
      Next U&
      FIN& = ULTREG&("!GRUPIVA")
      Call SETULTREG("GRUPIVA", FIN&)
   End If
   If ULTREG&("GRUPIVA") > 3 Then
      Call SETULTREG("GRUPIVA", 3)
   End If
   Call CIERRARCH("GRUPIVA")
End Sub
'
Sub CARTACONPRO()
    '
    Call BLANARCH("!TACONPRO")
    For i& = 1 To ULTREG&("TACONPRO")
        X$ = REGLEIDO$("TACONPRO", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!TACONPRO", X$)
        End If
    Next i&
    Call CIERRARCH("TACONPRO")
    '
    VTB% = VENTABU%("TACONPRO/NC/TITUPAN=Rubros De Proceso")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For i& = 1 To ULTREG&("!TACONPRO")
        X$ = REGLEIDO$("!TACONPRO", i&)
        If CVI(Left$(X$, 2)) > 0 Then
            J& = J& + 1
            Call GRAREG("TACONPRO", X$, J&)
        End If
    Next i&
    Call SETULTREG("TACONPRO", J&)
    Call CIERRARCH("TACONPRO")
    '
End Sub

'
Function EsCuentaRetIIBBProvincias%(CUECO%)
    EsCuentaRetIIBBProvincias% = 0
    If CUECO% < 1 Then Exit Function
    
    For i& = 1 To ULTREG&("PROVINB")
        X$ = REGLEIDO$("PROVINB", i&)
        RETIB% = CVI(Mid$(X$, 29, 2))
        If RETIB% > 0 Then
          If RETIB% = CUECO% Then
            'ACUMULO PARA PASAR LA CANTIDAD DE VECES QUE EXISTE.
            EsCuentaRetIIBBProvincias% = EsCuentaRetIIBBProvincias% + 1
          End If
        End If
    Next i&

End Function
