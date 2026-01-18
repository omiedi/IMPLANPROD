Attribute VB_Name = "ADMIPRE1"
'       ADMIPRE.QBS - administracion general de listas de precios
'
' ****************************************************************************
'
' faltan las siguientes opciones para completar este m¢dulo:
'
'       - Leyendas de pie de pagina
'       - Subtitulos y subrayados
'       - Ordenamiento por c¢digo
'       - Formato alternativo descuento por cantidad
'       - Listador con membrete, iva, salida a lotus, salto de hoja
'
' ****************************************************************************
Dim LISNO%
Dim TITU$
Dim MONE$
Dim FL$
Dim ALISTA$
'

Sub APRECIO1()
    '
    If ULTREG&("VALIFORM") < 2 Then
        Call BLANARCH("VALIFORM")
        Call REGAPP("VALIFORM", Chr$(1) + "Nro. Original")
        Call REGAPP("VALIFORM", Chr$(2) + "Descuento por Cantidad")
        Call CIERRARCH("VALIFORM")
    End If
    '
    If ULTREG&("VALIMONE") < 2 Then
        Call BLANARCH("VALIMONE")
        Call REGAPP("VALIMONE", Chr$(1) + "Pesos")
        Call REGAPP("VALIMONE", Chr$(2) + "Dolares U$S")
        Call CIERRARCH("VALIMONE")
    End If
    '
End Sub
'
        'MODO% = ALTERNA%("&11.20//COLOR=C-N/NC&Consulta y Actualizaci¢n\Listados por Impresora\-\Nueva Lista\Re-Configuraci¢n\-\Ordenamiento\Control de C¢digos\-\Ordenes de Compra Terminales\Datos Adicionales de Facturacion\-\Terminar")
        'On MODO% GoTo 1000, 2000, 100, 3000, 3100, 100, 3200, 3300, 100, 3400, 3500
        'GoTo 3900
        '
Sub APRECIO2()

1000    ' ******************************** consulta y actualizacion de precios
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) <= 4 Then
            Exit Sub
        End If
        '
1010    X$ = OBJPLA$("LISPRE-MUESTRA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        CADECI% = Asc(Mid$(X$, 49, 1))
        If CADECI% < 2 Or CADECI% > 4 Then CADECI% = 2
        '
        Call DATLIPRE(X$)        ' datos de lista elegida
        Call CIERRARCH("*.*")
        '
1100    LDT$ = LUDAT$
        Call BLANARCH("!LIPREVEN")
        Call COPYSTRU("LIPREVEN", "LIPRE" + ALISTA$)
        Call REGAPP("LIPRE" + ALISTA$, String$(128, 0))
        Call CIERRARCH("*.*")
        'If COPYFILE%(LDT$ + "LIPRE" + ALISTA$ + ".DAT", "C:\FLEXOFT\LIPREVEN.WKF") = 1 Then
        Call TRALOCAL("LIPRE" + ALISTA$, "LIPREVEN")
            XII% = VENTABU%("LIPREVEN/ARCHIVO=LIPRE" + ALISTA$ + "/NC/TITUPAN=" + TITU$)
            If XII% < 0 Then
                Exit Sub
            End If
            Call CIERRARCH("*.*")
            '
            EPAC$ = TRIM$(UCase$(EMPREAC$))
            If InStr(EPAC$, "DOSIVA") > 0 Then
              If VERILLAVE% <> 1 Then Exit Sub
            End If
            '
            'If COPYFILE%("C:\FLEXOFT\LIPREVEN.WKF", LDT$ + "LIPRE" + ALISTA$ + ".DAT") < 1 Then
            '    Exit Sub
            'End If
            Call TRACENTRAL("LIPRE" + ALISTA$, "LIPREVEN")
            '
1110        XII% = COMALTER%("Fin de Actualizacion\y Consulta de Precios\\Terminar\Reanudar\Descripciones")
            On XII% GoTo 1120, 1100, 1120
            GoTo 1110
            '
1120        LIPREAC$ = "LIPRE" + ALISTA$
            FIN& = ULTREG&(LIPREAC$)
            '
            Screen.MousePointer = 11
            '
            For I& = 1 To FIN&
                Call TRACE(20, I&, FIN&)
                NXX$ = REGLEIDO$(LIPREAC$, I&)
                CI% = CVI(Left$(NXX$, 2))
                MRX$ = Mid$(NXX$, 45, 1)
                If CI% > 0 Then
                    If CI% <= NUMAS% Then
                        PRELI# = CVD(Mid$(NXX$, 62, 8))
                        PREX$ = "  " + TRIM$(CSTRING$(MKD$(PRELI#), 3, 12, CADECI%, 2))
                        DE$ = DESCRIT0$(CI%)
                        Call REPLA$(NXX$, DE$, 3, 42)
                        Call REPLA(NXX$, PREX$ + "  ", 43 - Len(PREX$), 2 + Len(PREX$))
                        Call REPLA(NXX$, MRX$, 44, 1)
                        Call GRAREG(LIPREAC$, NXX$, I&)
                    End If
                End If
            Next I&
            '
            Call CIERRARCH(LIPREAC$)
            Screen.MousePointer = 1
        'End If
        Call CIERRARCH("*.*")
        '
1190 End Sub
'

Sub APRECIO3()

2000    ' ************** impresion de listas de precios
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) <= 4 Then
            Exit Sub
        End If
        '
2010    X$ = OBJPLA$("LISPRE-MUESTRA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        '
        Call DATLIPRE(X$)                                    ' datos de lista elegida
        '
2100    Call BLANARCH("LIPREVEN")
        Screen.MousePointer = 11
        '
        Call HEADERS("LIPREVEN", "")
        Call HEADERS("LIPREVEN", "Lista: " + TITU$)
        '
        LIPREAC$ = "LIPRE" + ALISTA$
        FIN& = ULTREG&(LIPREAC$)
        '
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           NXX$ = REGLEIDO$(LIPREAC$, I&)
           CI% = CVI(Left$(NXX$, 2))
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 RELIB% = Val(Mid$(NXX$, 45, 1))
                 For K% = 1 To RELIB%
                    Call REGAPP("LIPREVEN", MKI$(32767) + Mid$(REGBLAN$("LIPREVEN"), 3))
                 Next K%
                 '
                 Call REGAPP("LIPREVEN", NXX$)
              End If
           End If
        Next I&
        '
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*LILIPRE") '
        '
End Sub
'

Sub APRECIO4()

3000    ' ************************* nueva lista de precios
        '
        If ULTREG&("LISTAS") >= 99 Then
            Call COMUNI("Hay 99 Listas de Precios\ya Definidas\-\Capacidad del Sistema Saturada.")
            Exit Sub
        End If
        '
        NULIS% = 1 + ULTREG&("LISTAS")
        X$ = Chr$(NULIS%) + Mid$(REGBLAN$("LISTAS"), 2)
3010    X$ = OBJPLA$("LISPRE-NUEVA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        '
        If TRIM$(Mid$(X$, 2, 46)) <> "" Then
          If Asc(Mid$(X$, 48, 1)) > 0 Then
            If Asc(Mid$(X$, 48, 1)) <= 2 Then
              If Asc(Mid$(X$, 49, 1)) <= 4 Then
                If Asc(Mid$(X$, 50, 1)) > 0 Then
                  If Asc(Mid$(X$, 50, 1)) <= 2 Then
                    If TRIM$(Mid$(X$, 51, 12)) <> "" Then
                      If CVI(Mid$(X$, 63, 2)) > 0 Then
                        Call REGAPP("LISTAS", X$)
                        Call CIERRARCH("LISTAS")
                        Exit Sub
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
        '
        Call COMUNI("Datos Incompletos\o No Válidos")
        GoTo 3010
        '
End Sub
'
Sub APRECIO5()

3100    ' *************************************************** re-configuracion
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) < 4 + RECLEN%("LISTAS") Then
            Exit Sub
        End If
        '
        REG& = CVS(Mid$(X$, 1 + RECLEN%("LISTAS"), 4))
        '
3110    X$ = OBJPLA$("LISPRE-CAMBIA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        '
        If TRIM$(Mid$(X$, 2, 46)) <> "" Then
          If Asc(Mid$(X$, 48, 1)) > 0 Then
            If Asc(Mid$(X$, 48, 1)) <= 2 Then
              If Asc(Mid$(X$, 49, 1)) <= 4 Then
                If Asc(Mid$(X$, 50, 1)) > 0 Then
                  If Asc(Mid$(X$, 50, 1)) <= 2 Then
                    If TRIM$(Mid$(X$, 51, 12)) <> "" Then
                      If CVI(Mid$(X$, 63, 2)) > 0 Then
                        Call GRAREG("LISTAS", X$, REG&)
                        Call CIERRARCH("LISTAS")
                        Exit Sub
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
        '
        Call COMUNI("Datos Incompletos\o No Válidos")
        GoTo 3110
        '
End Sub

Sub APRECIO6()

3200    ' ***************** ordenamiento listas de precios
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) <= 4 Then
            Exit Sub
        End If
        '
3210    X$ = OBJPLA$("LISPRE-MUESTRA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        '
        Call DATLIPRE(X$)                                    ' datos de lista elegida
        '
3220    LIPREAC$ = "LIPRE" + ALISTA$
        FIN& = ULTREG&(LIPREAC$)
        If FIN& < 1 Then
            Exit Sub
        End If
        Call CIERRARCH(LIPREAC$)
        LDT$ = LUDAT$
        '
        Call REGAPP(LIPREAC$, String$(128, 0))
        Call CIERRARCH("*.*")
        'If COPYFILE%(LDT$ + LIPREAC$ + ".DAT", "C:\FLEXOFT\LIPREVEN.WKF") < 1 Then
        '    Exit Sub
        'End If
        Call TRALOCAL(LIPREAC$, "LIPREVEN")
        '
        Screen.MousePointer = 11
        FORLISPRE.SORTLIST.Clear
        '
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           NXX$ = REGLEIDO$(LIPREAC$, I&)
           CI% = CVI(Left$(NXX$, 2))
           CODE$ = CODEXT$(CI%)
           If CODE$ <> "" Then
              Z$ = Space$(24)
              Call REPLA(Z$, "A" + CODE$, 1, 18)
              Call REPLA(Z$, TRIM$(Str$(I&)), 19, 6)
              FORLISPRE.SORTLIST.AddItem Z$
           End If
        Next I&
        '
        FIN& = FORLISPRE.SORTLIST.ListCount
        '
        J& = 0
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           FORLISPRE.SORTLIST.ListIndex = I& - 1
           TESOR$ = FORLISPRE.SORTLIST.Text
           REG& = Val(Mid$(TESOR$, 19, 6))
           If REG& > 0 Then
               NXX$ = REGLEIDO$("!LIPREVEN", REG&)
               J& = J& + 1
               Call GRAREG(LIPREAC$, NXX$, J&)
           End If
        Next I&
        '
        Call SETULTREG&(LIPREAC$, J&)
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        '
End Sub


Sub APRECIO7()
        '
3300    ' ****** revision de codigos en listas de precios
        '
        X$ = REGSEL$("LISTAS")
        If Len(X$) <= 4 Then
            Exit Sub
        End If
        '
3310    X$ = OBJPLA$("LISPRE-MUESTRA", X$)
        If X$ = "" Then
            Exit Sub
        End If
        '
        Call DATLIPRE(X$)                                    ' datos de lista elegida
        '
        MOCON% = ALTERNA%("Revisión Individual\Eliminación Directa Códigos de Baja")
        If MOCON% < 1 Or MOCON% > 2 Then
            Exit Sub
        End If
        '
        If MOCON% = 2 Then
            If USNBR% < 100 Then
                Call COMUNI("La Opción Elegida Requiere\Privilegios de Supervisor.")
                Exit Sub
            End If
        End If
        '
3320    Call BLOQARCH(LIPREAC$)
        Screen.MousePointer = 11
        '
        LIPREAC$ = "LIPRE" + ALISTA$
        FIN& = ULTREG&(LIPREAC$)
        '
        '
        J& = 0
        For I& = 1 To FIN&
           Call TRACE(20, I&, FIN&)
           NXX$ = REGLEIDO$(LIPREAC$, I&)
           CI% = CVI(Left$(NXX$, 2))
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 COD$ = CODEXT$(CI%)
                 CI1% = CODINT%(COD$)
                 If CI1% < 1 Then
                     '
                     If MOCON% = 2 Then GoTo 3385
                     '
3335                 TXT$ = "Codigo: " + TRIM$(COD$) + "\" + TRIM$(DESCRIT$(CI%))
                     TXT$ = TXT$ + "\Figura de Baja en Maestro de Art¡culos."
                     TXT$ = TXT$ + "\\Ignorar\Reemplazar\Eliminar\Terminar Control"
                     OPX% = COMALTER%(TXT$)
                     On OPX% GoTo 3380, 3350, 3385
                     '
3340                 For K& = I& To FIN&
                        NXX$ = REGLEIDO$(LIPREAC$, K&)
                        J& = J& + 1
                        Call GRAREG(LIPREAC$, NXX$, J&)
                     Next K&
                     GoTo 3390
                     '
3350                 VDEF$ = Left$(NXX$, 2) + MKI$(0)
                     XXX$ = OBJPLA$("CAMBICOD", VDEF$)
                     If Len(XXX$) < 4 Then
                         Beep
                         GoTo 3335
                     End If
                     '
                     CI2% = CVI(Mid$(XXX$, 3, 2))
                     If CI2% < 1 Or CI2% > NUMAS% Then
                         Beep
                         GoTo 3335
                     End If
                     '
                     Call REPLA(NXX$, MKI$(CI2%), 1, 2)
                     Call REPLA(NXX$, DESCRIT0$(CI2%), 3, 42)
                     GoTo 3380
                     '
                 End If
                 '
3380             J& = J& + 1
                 Call GRAREG(LIPREAC$, NXX$, J&)
              End If
           End If
3385    Next I&
        '
        Call SETULTREG(LIPREAC$, J&)
        '
3390    Call CIERRARCH(LIPREAC$)
        Call LIBARCH(LIPREAC$)
        Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 1
        '
End Sub


Sub DATLIPRE(X$)
        '
        ' ******** presentacion de datos de lista elegida
        '
        LISNO% = Asc(Left$(X$, 1))
        TITU$ = TRIM$(Mid$(X$, 2, 46))
        MONE$ = TRIM$(ECOARCH$("VALIMONE", Mid$(X$, 50, 1)))
        TITU$ = TITU$ + " - " + MONE$
        FEX = CVI(Mid$(X$, 63, 2)): FL$ = FECHATEX$(FEX)
                               FL$ = REPLACAR$(FL$, "/", "-")
        TITU$ = TITU$ + " - Vig: " + FL$
        '
        ALISTA$ = TRIM$(Str$(LISNO%)): While Len(ALISTA$) < 3
                                    ALISTA$ = "0" + ALISTA$
                                  Wend
                                  '

        ' *** conversion de formato 'LPRnnn.DAT' a 'LIPREnnn.DAT'
        '
        LU$ = LUDAT$
        ANUE$ = "LIPRE" + ALISTA$ + ".DAT"
        N2% = FreeFile
        N2% = FILEOPEN%(LU$ + ANUE$, 0, 72)
        Dim NUE As String * 72
        '
        If LOF(N2%) < 72 Then
            AVIE$ = "LPR" + ALISTA$ + ".DAT"
            N1% = FreeFile
            N1% = FILEOPEN%(LU$ + AVIE$, 0, 32)
            Dim VIE As String * 32
            '
            Screen.MousePointer = 11
            FIN& = LOF(N1%) / 32
            J& = 0
            For I& = 1 To FIN&
               Get #N1%, I&, VIE$
               CI% = CVI(Mid$(VIE$, 3, 2))
               If CI% > 0 Then
                   If CI% <= NUMAS% Then
                        CII% = CI%: DE$ = DESCRIT$(CII%)
                        NXX$ = String$(72, 0)
                        NXX$ = REPLACE$(NXX$, MKI$(CI%), 1, 2)
                        NXX$ = REPLACE$(NXX$, DE$, 3, 42)
                        NXX$ = REPLACE$(NXX$, Mid$(VIE$, 5, 28), 45, 28)
                        J& = J& + 1
                        LSet NUE$ = NXX$
                        Put #N2%, J&, NUE$
                   End If
               End If
            Next I&
            '
            ANUE$ = "LIPRE" + ALISTA$
            Call COPYSTRU("LIPREVEN", ANUE$)
            '
        End If
        '
        If N2% > 0 Then Close #N2%: N2% = 0
        If N1% > 0 Then Close #N1%: N1% = 0
        '
        Screen.MousePointer = 1
        '
End Sub

Function VERILLAVE%()
   '
   VERILLAVE% = 0
10 OPX% = COMALTER%("Coloque Diskette Llave\\Continuar\Cancelar")
   If OPX% = 1 Then
      On Error GoTo 50
      SVL$ = SERIVOL$("A:")
      TECONTROL$ = ENCRIPTA$(TRIM$(UCase$(App.EXEName)) + SVL$)
      NX1% = FreeFile
      Dim TTX As String * 128
      Open "A:\" + TRIM$(UCase$(App.EXEName)) + ".KEY" For Random Access Read Write As #NX1% Len = 128
      Get #NX1%, 2, TTX$
      LTC% = Len(TECONTROL$)
VERILLAVE% = 1: GoTo 99
      If CVI(Mid$(TTX$, 18, 2)) = LTC% Then
        If Mid$(TTX$, 30, LTC%) = TECONTROL$ Then
          VERILLAVE% = 1
          GoTo 99
        End If
      End If
      Call COMUNI("Llave no Válida")
   End If
   GoTo 99
   '
50 Resume 51
51 GoTo 10
   '
99 End Function
