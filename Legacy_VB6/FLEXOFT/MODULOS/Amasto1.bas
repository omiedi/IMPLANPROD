Attribute VB_Name = "AMASTO1"
Public MODIFI%
Public VALICOD%
Public SUCCESS%
Public TIMA%
Public CRIT%
Public CCOMPO%
Public MONEI%
Public STMI&, STMII%
Public LORE&, LOREI%, LOTEMAI%
Public UNID$
Public COU
Public FECO%
Public PPP
Public FULT%
Public NPRO%
Public GRUCO%
Public FCP%
Public TCOLOR$
Public TTALLE$
Public TTITPE$

Public VACONTRO%
Public AGREPRECIO%

'

Sub Main()
    '
    Call VERJOBID(OKEY%)
    '
    COMMA$ = Command$
    PPX% = InStr(COMMA$, "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    COMMA$ = TRIM$(UCase$(COMMA$))
    If COMMA$ = "NEWITEM" Then
        ALTAITEM.Show
        ALTAITEM.Refresh
      Else
        AMASTOK.Show
        AMASTOK.Refresh
    End If
    
End Sub
'

Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Call MENSERR(24, "Falta Definir Depósitos")
        Call CARTADEPOSI
        GoTo 10
    End If
    '
    If ULTREG&("UMEDIDA") < 10 Then
        Call BLANARCH("UMEDIDA")
        Call REGAPP("UMEDIDA", "U.Unidades")
        Call REGAPP("UMEDIDA", "KgKilogramos")
        Call REGAPP("UMEDIDA", "grGramos")
        Call REGAPP("UMEDIDA", "gKGramos - Kilos")
        Call REGAPP("UMEDIDA", "M.Metros")
        Call REGAPP("UMEDIDA", "M2M.Cuadrados")
        Call REGAPP("UMEDIDA", "M3M.Cúbicos")
        Call REGAPP("UMEDIDA", "mmMilímetros")
        Call REGAPP("UMEDIDA", "cmCentímetros")
        Call REGAPP("UMEDIDA", "LtLitros")
        Call CIERRARCH("UMEDIDA")
    End If
    XX$ = REGLEIDO$("UMEDIDA", 1)
    '
    If ULTREG&("UNIMED") < 10 Then
        Call BLANARCH("UNIMED")
        Call REGAPP("UNIMED", "U.  Unidades" + String$(256, 0))
        Call REGAPP("UNIMED", "Kg  Kilogramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gr  Gramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gK  Gramos - Kilos" + String$(256, 0))
        Call REGAPP("UNIMED", "M.  Metros" + String$(256, 0))
        Call REGAPP("UNIMED", "M2  M.Cuadrados" + String$(256, 0))
        Call REGAPP("UNIMED", "M3  M.Cúbicos" + String$(256, 0))
        Call REGAPP("UNIMED", "mm  Milímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "cm  Centímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "Lt  Litros" + String$(256, 0))
        Call CIERRARCH("UNIMED")
    End If
    XX$ = REGLEIDO$("UNIMED", 1)
    '
    If ULTREG&("TATIMAT") < 1 Then
        Call BLANARCH("TATIMAT")
        Call REGAPP("TATIMAT", Chr$(1) + "Materia Prima")
        Call REGAPP("TATIMAT", Chr$(2) + "Pieza Comercial")
        Call REGAPP("TATIMAT", Chr$(3) + "Pieza Elaborada")
        Call REGAPP("TATIMAT", Chr$(4) + "Semielaborado")
        Call REGAPP("TATIMAT", Chr$(5) + "Prod.Terminado")
        Call REGAPP("TATIMAT", Chr$(8) + "Servicios")
        Call REGAPP("TATIMAT", Chr$(9) + "Improductivo")
        Call CIERRARCH("TATIMAT")
    End If
    XX$ = REGLEIDO$("TATIMAT", 1)
    '
    If ULTREG&("TABTIMA") < 1 Then
       DPPRE% = Asc(Left$(REGLEIDO$("ECODEP", 1), 1))
       For U& = 1 To ULTREG&("TATIMAT")
          x$ = REGLEIDO$("TATIMAT", U&)
          Y$ = REGBLAN$("TABTIMA")
          Call REPLA(Y$, x$, 1, 16)
          Call REPLA(Y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", Y$)
       Next U&
       Call CIERRARCH("TABTIMA")
    End If
    XX$ = REGLEIDO$("TABTIMA", 1)
    '
    If ULTREG&("TACRIRE") < 4 Then
        Call BLANARCH("TACRIRE")
        Call REGAPP("TACRIRE", Chr$(1) + "Por Pedido")
        Call REGAPP("TACRIRE", Chr$(2) + "Por Mínimos")
        Call REGAPP("TACRIRE", Chr$(3) + "Cons.Estadístico")
        Call REGAPP("TACRIRE", Chr$(4) + "Programación")
        Call CIERRARCH("TACRIRE")
    End If
    XX$ = REGLEIDO$("TACRIRE", 1)
    '
    If ULTREG&("TAMONED") < 2 Then
        Call BLANARCH("TAMONED")
        Call REGAPP("TAMONED", Chr$(1) + AJUSTI$("Pesos", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(2) + AJUSTI$("Dolares U$S", 27) + MKS$(1))
        Call CIERRARCH("TAMONED")
    End If
    XX$ = REGLEIDO$("TAMONED", 1)
    '
    If ULTREG&("TACOMPO") < 2 Then
        Call BLANARCH("TACOMPO")
        Call REGAPP("TACOMPO", Chr$(0) + "Manual")
        Call REGAPP("TACOMPO", Chr$(1) + "Automática")
        Call CIERRARCH("TACOMPO")
    End If
    XX$ = REGLEIDO$("TACOMPO", 1)
    '
    If ULTREG&("GRUPIVA") < 1 Then
        Call SETULTREG("GRUPIVA", 0)
        Call REGAPP("GRUPIVA", Chr$(0) + "IVA 21 %    " + MKS$(21))
        Call REGAPP("GRUPIVA", Chr$(1) + "IVA 10,5 %    " + MKS$(10.5))
        Call REGAPP("GRUPIVA", Chr$(2) + "IVA 27 %    " + MKS$(27))
        Call CIERRARCH("GRUPIVA")
    End If
    XX$ = REGLEIDO$("GRUPIVA", 1)
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub NEWITEM()
    '
    HOO% = HOY(HO$)
    VDEF$ = Space$(80) + String$(48, 0)
    Call REPLA(VDEF$, "U.", 81, 2)
    Call REPLA(VDEF$, Chr$(1), 101, 1)
    Call REPLA(VDEF$, MKI$(HOO%), 102, 2)
    '
10  REMANU$ = OBJPLA$("NEWITEM", VDEF$)
    If REMANU$ = "" Then
        Exit Sub
    End If
    '
    VDEF$ = REMANU$
    COD$ = Left$(REMANU$, 16)
    If TRIM$(COD$) = "" Then
        Call MENSERR(24, "Falta Código de Artículo")
        GoTo 10
    End If
    '
    DE$ = Mid$(REMANU$, 17, 64)
    If TRIM$(DE$) = "" Then
        Call MENSERR(24, "Falta Descripción de Artículo")
        GoTo 10
    End If
    '
    UNID$ = Mid$(REMANU$, 81, 2)
    TIMA% = Asc(Mid$(REMANU$, 83, 1))
    CRIT% = Asc(Mid$(REMANU$, 84, 1))
    '
    STMIN = CVS(Mid$(REMANU$, 85, 4))
    STMI& = 0
    If STMIN <= 30000 Then
          STMI& = STMIN
        ElseIf STMIN <= 3000000 Then
          STMI& = (-1) * STMIN / 100
    End If
    STMII% = STMI&
    '
    LOREP = CVS(Mid$(REMANU$, 89, 4))
    LORE& = 0
    If LOREP <= 30000 Then
          LORE& = LOREP
        ElseIf LOREP <= 3000000 Then
          LORE& = (-1) * LOREP / 100
    End If
    LOREI% = LORE&
    '
    LOTEMA = CVS(Mid$(REMANU$, 93, 4))
    If LOTEMA < LOREP Then LOTEMA = LOREP
    LOMA& = LORE&
    If LOTEMA <= 30000 Then
          LOMA& = LOTEMA
        ElseIf LOTEMA <= 3000000 Then
          LOMA& = (-1) * LOTEMA / 100
    End If
    LOTEMAI% = LOMA&
    '
    COU = CVS(Mid$(REMANU$, 97, 4))
    MONEI% = Asc(Mid$(REMANU$, 101, 1))
    FECO% = CVI(Mid$(REMANU$, 102, 2))
    FCP% = CVI(Mid$(REMANU$, 127, 2))
    '
    AMASTOK.Show
    Call FRESHALL
    Call ALTAMAS
    Call FRESHECHO("MASTER")
    AMASTOK.Hide
    Call FRESHALL
    '
End Sub
'

Sub ALTAMAS()
    
    Screen.MousePointer = 11
    SUCCESS% = 0
    '
    NM& = NUMAS%
'    Call BLOQARCH("INVERT")
'    Call BLOQARCH("MASTER")
    '
    If Len(TRIM$(COD$)) < 3 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Alta Imposible\Código no Válido.")
        Exit Sub
    End If
    '
    If CODINT%(COD$) <> 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Alta Imposible\Código Existente.")
        Exit Sub
    End If
    '
    CODD$ = AJUSTI$(COD$, 15)
    For U& = 1 To AMASTOK.List1.ListCount
       If Mid$(AMASTOK.List1.LIST(U& - 1), 2, 15) = CODD$ Then
          Call MENSERR(24, "Alta Imposible\Código Existente.")
          Exit Sub
       End If
    Next U&
    '
    SUCCESS% = (-1)
    '
    CI% = ULTREG&("MASTER")
    '
    If CI% < 1 Then
         Call BLANARCH("SALSTOCK")
         Call REGAPP("MASTER", MKI$(0) + String$(125, 0) + Chr$(1))
         Call REGAPP("SALSTOCK", String$(39, 0) + Chr$(1))
         Call CIERRARCH("MASTER")
         CI% = 1
     End If
     '
     PPP = 0
     FULT% = 0
     FCP% = 0
     '
     REMAS$ = REGBLAN$("MASTER")
     Call REPLA(REMAS$, MKI$(CI%), 1, 2)
     Call REPLA(REMAS$, DE$, 5, 64)
     Call REPLA(REMAS$, String$(8, 0), 69, 8)
     Call REPLA(REMAS$, COD$, 77, 16)
     Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
     Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
     Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
     Call REPLA(REMAS$, String$(7, 0), 96, 7)
     Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
     Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
     Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
     Call REPLA(REMAS$, UNID$, 109, 2)
     Call REPLA(REMAS$, MKS$(COU), 111, 4)
     Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
     Call REPLA(REMAS$, MKS$(PPP), 117, 4)
     Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
     Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
     Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
     Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
     '
     Call REGAPP("MASTER", REMAS$)
     x$ = REGLEIDO$("MASTER", 1&)
     x$ = REPLACE$(x$, MKI$(CI%), 1, 2)
     Call GRAREG("MASTER", x$, 1&)
     '
     For U& = ULTREG&("SALSTOCK") + 1 To ULTREG&("MASTER")
         Call REGAPP("SALSTOCK", String$(39, 0) + Chr$(1))
     Next U&
     Call GRAREG("SALSTOCK", String$(39, 0) + Chr$(1), NM& + 1)
     '
     UX& = ULTREG&("INVERT")
     If UCase$(App.EXEName) = "FORMULAS" Then
        Call REGAPP("INVERT", String$(18, 255))
        IX$ = Space$(16) + MKI$(CI%)
        IX$ = REPLACE$(IX$, COD$, 1, 16)
        Call REPLA(IX$, COD$, 1, 16)
     End If
     '
     ZZ$ = "A" + AJUSTI$(COD$, 16) + CSTRING$(MKI$((-1) * CI%), 1, 7, 0, 2)
     'AMASTOK.List1.AddItem ZZ$
     Call AGRELISTA(ZZ$)
     If UCase$(App.EXEName) = "FORMULAS" Then
        For i& = UX& To 2 Step -1
           II$ = REGLEIDO$("INVERT", i&)
           If Left$(II$, 16) > COD$ Then
                 Call GRAREG("INVERT", II$, i& + 1)
               Else
                 Call GRAREG("INVERT", IX$, i& + 1)
                 GoTo 15800
           End If
        Next i&
        Call GRAREG("INVERT", IX$, 2&)
        '
15800   NVI% = ULTREG&("INVERT") - 1
        II$ = REGLEIDO$("INVERT", 1&)
        II$ = REPLACE$(II$, MKI$(NVI%), 17, 2)
        Call GRAREG("INVERT", II$, 1&)
        '
15850   EX$ = AJUSTI$(COD$, 16) + Mid$(REMAS$, 5, 46) + MKI$(CI%)
        cx$ = MKI$(CI%)
        '
        UX& = ULTREG&("ECOMAST")
        FIECO% = FILENBR%("ECOMAST")
            Dim EL As String * 64
        FICIA% = FILENBR%("CIACTI")
            Dim CL As String * 2
            '
        For i& = UX& To 1 Step -1
           Get #FIECO%, i&, EL$
           Get #FICIA%, i& + 1, CL$
           If Left$(EL$, 16) > COD$ Or EL$ = String$(64, 0) Then
                 Put #FIECO%, i& + 1, EL$
                 Put #FICIA%, i& + 2, CL$
               Else
                 EL$ = EX$: Put #FIECO%, i& + 1, EL$
                 CL$ = cx$: Put #FICIA%, i& + 2, CL$
                 GoTo 15890
           End If
        Next i&
        EL$ = EX$: Put #FIECO%, 1, EL$
        CL$ = cx$: Put #FICIA%, 2, CL$
        '
15890   Get #FICIA%, 1, CL$
        CL$ = MKI$(UX& + 1)
        Put #FICIA%, 1, CL$
        '
     End If
        '
15900   Call CIERRARCH("MASTER")
        Call CIERRARCH("INVERT")
        Call CIERRARCH("SALSTOCK")
        Call CIERRARCH("ECOMAST")
        Call CIERRARCH("CIACTI")
        '
        SUCCESS% = 1
        Screen.MousePointer = 1
        Call COMUNI("Alta de Código " + TRIM$(COD$) + " Completa y Verificada.")
        '
    PORCALPRE = Val(CONTROL$("", "PORCALPRE"))
    If PORCALPRE > 0 Then
      If COU >= 0.005 Then
        PREVENCA = (Int(COU * (100 + PORCALPRE) + 0.5)) / 100
        If COMALTER%("Puede Ahora Agregar el Articulo\a la Primera Lista de Precios\al Valor de " + TRIM$(CSTRING$(MKS$(PREVENCA), 3, 12, 2, 2)) + "\\Agregar\Cancelar") = 1 Then
          XXP$ = REGBLAN$("LIPRE001")
          Call REPLA(XXP$, MKI$(CI%), 1, 2)
          Call REPLA(XXP$, DESCRIT0$(CI%), 3, 42)
          Call REPLA(XXP$, MKD$(CDbl(PREVENCA)), 62, 8)
          Call REGAPP("LIPRE001", XXP$)
          AGREPRECIO% = 1
          Call CIERRARCH("LIPRE001")
        End If
      End If
    End If
    '
End Sub
'

Sub ALTAPORCOPIA()
    '
    Screen.MousePointer = 11
    AMASTOK.List1.Clear
    Dim II As String * 18
    N1% = FILENBR%("INVERT")
    FIN& = NUINV%
    For U& = 1 To FIN&
      Call TRACE(1, U&, FIN&)
      Get #N1%, U& + 1, II$
      ZZ$ = "A" + Left$(II$, 16) + CSTRING$(Mid$(II$, 17, 2), 1, 7, 0, 2)
      'AMASTOK.List1.AddItem ZZ$
      Call AGRELISTA(ZZ$)
    Next U&
    Screen.MousePointer = 1
    '
10  VDEF$ = ""
    OBX$ = OBJPLA$("ALTAPORCOPIA", MKI$(CIORI%))
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    CIORI% = CVI(Left$(OBX$, 2))
    If CIORI% < 1 Or CIORI% > NUMAS% Then
        Call MENSERR(24, "Codigo origen no Válido")
        GoTo 10
    End If
    '
    CODORI% = CIORI%
    VDEF$ = MKI$(CIORI%) + Space$(80)
    Call REPLA(VDEF$, DESCRIT0$(CIORI%), 19, 64)
    OBX$ = OBJPLA$("ALTAPORCOPIB", VDEF$)
    If OBX$ = "" Then
        Exit Sub
    End If
    '
    Call GRACONTROL("", "AMASTO", "ABORT " + HOS$ + " - " + Time$ + " hs.")
    '
    COD$ = TRIM$(Mid$(OBX$, 3, 16))
    CODEST% = ULTREG&("MASTER")
    DE$ = Mid$(OBX$, 19, 64)
    REGMAI& = CIORI% + 1
    REMAS$ = REGLEIDO$("MASTER", REGMAI&)
    '
    TIMA% = Asc(Mid$(REMAS$, 93, 1))
    CRIT% = Asc(Mid$(REMAS$, 94, 1))
    MONEI% = Asc(Mid$(REMAS$, 95, 1))
    LOTEMAI% = CVI(Mid$(REMAS$, 103, 2))
    STMII% = CVI(Mid$(REMAS$, 105, 2))
    LOREI% = CVI(Mid$(REMAS$, 107, 2))
    UNID$ = Mid$(REMAS$, 109, 2)
    COU = CVS(Mid$(REMAS$, 111, 4))
    FECO% = CVI(Mid$(REMAS$, 115, 2))
    PPP = CVS(Mid$(REMAS$, 117, 4))
    FULT% = CVI(Mid$(REMAS$, 121, 2))
    NPRO% = CVI(Mid$(REMAS$, 123, 2))
    GRUCO% = CVI(Mid$(REMAS$, 125, 2))
    FCP% = CVI(Mid$(REMAS$, 127, 2))
    '
    Call ALTAMAS
    If SUCCESS% <> 1 Then
        GoTo 99
    End If
    '
    Screen.MousePointer = 11
    Call FRESHECHO("MASTER")
    j& = 0
    PUNTCI$ = ""
    ESP1$ = "": ESP2$ = ""
    For i& = 1 To ULTREG&("ESTRUNUE")
        EST1$ = REGLEIDO$("ESTRUNUE", i&)
        CIV% = CVI(Left$(EST1$, 2))
        EST2$ = REGLEIDO$("ESTRUDOS", i&)
        If CIV% > 0 Then
           If CIV% <= NUMAS% Then
              If CIV% <> CODEST% Then
                 j& = j& + 1
                 Call GRAREG("ESTRUNUE", EST1$, j&)
                 Call GRAREG("ESTRUDOS", EST2$, j&)
                 PUNTCI$ = PUNTCI$ + MKI$(CIV%)
              End If
           End If
        End If
        '
        If CIV% = CODORI% Then
           If ESP1$ = "" Then
              ESP1$ = EST1$
              ESP2$ = EST2$
           End If
        End If
    Next i&
    '
    Call REPLA(ESP1$, MKI$(CODEST%), 1, 2)
    EST1$ = ESP1$
    EST2$ = ESP2$
    j& = j& + 1
    Call GRAREG("ESTRUNUE", EST1$, j&)
    Call GRAREG("ESTRUDOS", EST2$, j&)
    PUNTCI$ = PUNTCI$ + MKI$(CODEST%)
    '
    For i& = j& + 1 To ULTREG&("ESTRUNUE")
        Call GRAREG("ESTRUNUE", String$(128, 0), i&)
        Call GRAREG("ESTRUDOS", String$(256, 0), i&)
    Next i&
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    Screen.MousePointer = 11
    PRIREORI& = 0
    PRIREDEST& = 0
    FIN& = ULTREG&("SECOPER")
    For i& = 1 To FIN&
       x$ = REGLEIDO$("SECOPER", i&)
       CI% = CVI(Left$(x$, 2))
       If CI% = CODORI% Then
           If PRIREORI& = 0 Then
               PRIREORI& = i&
           End If
       End If
       '
       If CI% = CODEST% Then
           If PRIREDEST& = 0 Then
               PRIREDEST& = i&
           End If
       End If
       '
    Next i&
    '
    If PRIREORI& < 1 Then
        Screen.MousePointer = 1
        GoTo 99
    End If
    '
    If PRIREDEST& > 0 Then
        j& = PRIREDEST& - 1
        For i& = PRIREDEST& To FIN&
           x$ = REGLEIDO$("SECOPER", i&)
           CI% = CVI(Left$(x$, 2))
           If CI% <> CODEST% Then
               j& = j& + 1
               Call GRAREG("SECOPER", x$, j&)
           End If
        Next i&
        '
        For i& = j& + 1 To FIN&
           x$ = String$(128, 0)
           Call GRAREG("SECOPER", x$, i&)
        Next i&
        Call CIERRARCH("SECOPER")
        PRIREORI& = 1
    End If
    '
    FIN& = ULTREG&("SECOPER")
    For i& = PRIREORI& To FIN&
        x$ = REGLEIDO$("SECOPER", i&)
        CI% = CVI(Left$(x$, 2))
        If CI% = CODORI% Then
            Call REPLA(x$, MKI$(CODEST%), 1, 2)
            Call REGAPP("SECOPER", x$)
        End If
    Next i&
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    GoTo 10
    '
End Sub
'

Sub CLEARFIELDS()
    '
    ALTAITEM.Text1.TEXT = ""
    ALTAITEM.Label22.Caption = ""
    ALTAITEM.Text3.TEXT = ""
    ALTAITEM.Text4.TEXT = ""
    ALTAITEM.Text5.TEXT = ""
    ALTAITEM.Text6.TEXT = ""
    ALTAITEM.Text7.TEXT = ""
    ALTAITEM.Text8.TEXT = ""
    ALTAITEM.text9.TEXT = ""
    ALTAITEM.Text10.TEXT = ""
    ALTAITEM.Text11.TEXT = ""
    ALTAITEM.Text12.TEXT = ""
    ALTAITEM.Text13.TEXT = ""
    ALTAITEM.Text14.TEXT = ""
    ALTAITEM.Text15.TEXT = ""
    ALTAITEM.Text16.TEXT = ""
    ALTAITEM.Text17.TEXT = ""
    ALTAITEM.Text18.TEXT = ""
    ALTAITEM.Text19.TEXT = ""
    ALTAITEM.Text20.TEXT = ""
    '
    COD$ = ""
    DE$ = ""
    TIMA% = 0
    CRIT% = 0
    MONEI% = 0
    STMI& = 0: STMII% = 0
    LORE& = 0: LOREI% = 0
    LOMA& = 0: LOTEMAI% = 0
    UNID$ = ""
    COU = 0
    FECO% = 0
    PPP = 0
    FULT% = 0
    NPRO% = 0
    GRUCO% = 0
    FCP% = 0
    '
    ALTAITEM.Refresh
    '
End Sub

Sub CLEARMODI()
    '
    MODIITEM.Text1.TEXT = ""
    MODIITEM.Text2.TEXT = ""
    MODIITEM.Text3.TEXT = ""
    MODIITEM.Text4.TEXT = ""
    MODIITEM.Text5.TEXT = ""
    MODIITEM.Text6.TEXT = ""
    MODIITEM.Text7.TEXT = ""
    MODIITEM.Text8.TEXT = ""
    MODIITEM.text9.TEXT = ""
    MODIITEM.Text10.TEXT = ""
    MODIITEM.Text11.TEXT = ""
    MODIITEM.Text12.TEXT = ""
    MODIITEM.Text13.TEXT = ""
    MODIITEM.Text14.TEXT = ""
    MODIITEM.Text15.TEXT = ""
    MODIITEM.Text16.TEXT = ""
    MODIITEM.Text17.TEXT = ""
    MODIITEM.Text18.TEXT = ""
    MODIITEM.Text19.TEXT = ""
    MODIITEM.Text20.TEXT = ""
    MODIITEM.Text21.TEXT = ""
    MODIITEM.Text22.TEXT = ""
    MODIITEM.Text23.TEXT = ""
    MODIITEM.Text24.TEXT = ""
    MODIITEM.Text25.TEXT = ""
    MODIITEM.Text26.TEXT = ""
    '
    COD$ = ""
    DE$ = ""
    TIMA% = 0
    CRIT% = 0
    MONEI% = 0
    STMI& = 0: STMII% = 0
    LORE& = 0: LOREI% = 0
    LOMA& = 0: LOTEMAI% = 0
    UNID$ = ""
    COU = 0
    FECO% = 0
    PPP = 0
    FULT% = 0
    NPRO% = 0
    GRUCO% = 0
    FCP% = 0
    VALICOD% = 0
    '
    MODIITEM.Refresh
    '
End Sub '

Sub CARDATIT(COD$)
    '
    VALICOD% = 0
    CI% = CODINT%(COD$)
    If CI% < 1 Then
        Exit Sub
    End If
    '
    REG& = CI% + 1
    REMAS$ = REGLEIDO$("MASTER", REG&)
    '
    MODIITEM.Text2.TEXT = Str$(CI%)
    MODIITEM.Text3.TEXT = DESCRIT0$(CI%)
    UNID$ = UNIMED$(CI%)
    MODIITEM.Text4.TEXT = UNID$
    MODIITEM.Text5.TEXT = ECOARCH$("UMEDIDA", UNID$)
    TIMA% = Asc(Mid$(REMAS$, 93, 1))
    MODIITEM.Text6.TEXT = TRIM$(Str$(TIMA%))
    MODIITEM.Text7.TEXT = ECOARCH$("TATIMAT", Chr$(TIMA%))
    CRIT% = Asc(Mid$(REMAS$, 94, 1))
    MODIITEM.Text8.TEXT = TRIM$(Str$(CRIT%))
    MODIITEM.text9.TEXT = ECOARCH$("TACRIRE", Chr$(CRIT%))
    MONEI% = Asc(Mid$(REMAS$, 95, 1))
    MODIITEM.Text13.TEXT = TRIM$(Str$(MONEI%))
    MODIITEM.Text15.TEXT = ECOARCH$("TAMONED", Chr$(MONEI%))
    GRUCO% = CVI(Mid$(REMAS$, 125, 2))
    MODIITEM.Text16.TEXT = TRIM$(Str$(GRUCO%))
    MODIITEM.Text17.TEXT = ECOARCH$("GRUCOVEN", MKI$(GRUCO%))
    NPRO% = CVI(Mid$(REMAS$, 123, 2))
    MODIITEM.Text10.TEXT = TRIM$(Str$(NPRO%))
    MODIITEM.Text11.TEXT = ECOARCH$("PROVED1", MKI$(NPRO%))
    MODIITEM.Text18.TEXT = TRIM$(Str$(STOMIN(CI%)))
    MODIITEM.Text19.TEXT = TRIM$(Str$(LOTESTAN(CI%)))
    MODIITEM.Text20.TEXT = TRIM$(Str$(LOREPOMA(CI%)))
    MODIITEM.Text12.TEXT = TRIM$(CSTRING$(MKS$(COSUNI(CI%)), 3, 12, 4, 2))
    FEX = CVI(Mid$(REMAS$, 115, 2))
    MODIITEM.Text14.TEXT = FECHATEX$(FEX)
    FCP% = CVI(Mid$(REMAS$, 127, 2))
    MODIITEM.Text21.TEXT = TRIM$(Str$(FCP%))
    CC% = Asc(Mid$(REMAS$, 96, 1))
    MODIITEM.Text22.TEXT = LCOLOR$(CC%)
    CT% = Asc(Mid$(REMAS$, 97, 1))
    MODIITEM.Text23.TEXT = LTALLE$(CT%)
    CP = CVS(Mid$(REMAS$, 99, 4))
    MODIITEM.Text24.TEXT = LTITPE$(CP)
    CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
    MODIITEM.Text25.TEXT = TRIM$(Str$(CCOMPO%))
    MODIITEM.Text26.TEXT = ECOARCH$("TACOMPO", Chr$(CCOMPO%))
    '
    VALICOD% = 1
    '
End Sub

Sub MODIMAS()
    '
    Screen.MousePointer = 11
    '
10  CI% = CODINT%(COD$)
    If CI% > 0 Then
        PESU = Val(TTITPE$)
        If InStr(TTITPE$, "/") > 0 Then
           TTQP$ = TTITPE$
           TTQP$ = REPLACAR$(TTQP$, "/", ".")
           PESU = (-1) * Abs(Val(TTQP$))
        End If
        '
        REG& = CI% + 1
        REMAS$ = REGLEIDO$("MASTER", REG&)
        If AJUSTI$(DE$, 64) <> Mid$(REMAS$, 5, 64) Then
            Call FRESHECHO("MASTER")
        End If
        '
        Call GRACONTROL("", "AMASTO", "ABORT " + HOS$ + " - " + Time$ + " hs.")
        '
        Call REPLA(REMAS$, DE$, 5, 64)
        Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
        Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
        Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
        Call REPLA(REMAS$, Chr$(CCOLOR%(TCOLOR$)), 96, 1)
        Call REPLA(REMAS$, Chr$(CTALLE%(TTALLE$)), 97, 1)
        Call REPLA(REMAS$, Chr$(CCOMPO%), 98, 1)
        Call REPLA(REMAS$, MKS$(PESU), 99, 4)
        Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
        Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
        Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
        Call REPLA(REMAS$, UNID$, 109, 2)
        Call REPLA(REMAS$, MKS$(COU), 111, 4)
        Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
        Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
        Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
        Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
        '
        Call GRAREG("MASTER", REMAS$, REG&)
    End If
    '
    CODX$ = AJUSTI$(COD$, 16)
    UX& = ULTREG&("ECOMAST")
    FIECO% = FILENBR%("ECOMAST")
    Dim EL As String * 64
    '
    For U& = 1 To UX&
       Get #FIECO%, U&, EL$
       If Left$(EL$, 16) = CODX$ Then
           EX$ = EL$
           Call REPLA(EX$, DE$, 17, 46)
           EL$ = EX$
           Put #FIECO%, U&, EL$
           GoTo 989
        End If
    Next U&
    '
989 PORCALPRE = Val(CONTROL$("", "PORCALPRE"))
    If PORCALPRE > 0 Then
      If COU >= 0.005 Then
        PREVENCA = (Int(COU * (100 + PORCALPRE) + 0.5)) / 100
        RFF$ = RECFILT$("LIPRE001", 1, MKI$(CI%))
        If Len(RFF$) >= 4 Then
          RELIS& = CVS(Left$(RFF$, 4))
          If RELIS& > 0 Then
            If RELIS& <= ULTREG&("LIPRE001") Then
              If COMALTER%("Puede Ahora Actualizar el Precio de Venta\de la Primera Lista de Precios a " + TRIM$(CSTRING$(MKS$(PREVENCA), 3, 12, 2, 2)) + "\\Actualizar\Cancelar") = 1 Then
                XXP$ = REGLEIDO$("LIPRE001", RELIS&)
                If CVI(Left$(XXP$, 2)) = CI% Then
                  Call REPLA(XXP$, MKD$(CDbl(PREVENCA)), 62, 8)
                  Call GRAREG("LIPRE001", XXP$, RELIS&)
                End If
              End If
              GoTo 999
            End If
          End If
        End If
        '
        If COMALTER%("Puede Ahora Agregar el Articulo\a la Primera Lista de Precios\al Valor de " + TRIM$(CSTRING$(MKS$(PREVENCA), 3, 12, 2, 2)) + "\\Agregar\Cancelar") = 1 Then
          XXP$ = REGBLAN$("LIPRE001")
          Call REPLA(XXP$, MKI$(CI%), 1, 2)
          Call REPLA(XXP$, DESCRIT0$(CI%), 3, 42)
          Call REPLA(XXP$, MKD$(CDbl(PREVENCA)), 62, 8)
          Call REGAPP("LIPRE001", XXP$)
          AGREPRECIO% = 1
        End If
      End If
    End If
    '
999 Call CIERRARCH("*.*")
    Call GRACONTROL("", "AMASTO", "NORMAL")
    Call CLEARMODI
    VALICOD% = (-1)
    '
    Screen.MousePointer = 1
End Sub
'

Sub BAJAMAS()
    '
5   VDEF$ = MKI$(0)
10  XX$ = OBJPLA$("ENTRACODIGO/Baja Lógica de Item de Stock", VDEF$)
    If XX$ = "" Then
       Exit Sub
    End If
    '
    CI% = CVI(XX$)
    If CI% < 1 Or CI% > NUMAS% Then GoTo 10
    Call GRACONTROL("", "AMASTO", "ABORT " + HOS$ + " - " + Time$ + " hs.")
    COD$ = CODEXT$(CI%)
    'Call BLOQARCH("INVERT")
    GoSub 125
    If ENCU% = 1 Then
        If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") = 2 Then
            'Call BLOQARCH("ECOMAST")
            'Call BLOQARCH("CIACTI")
            '
            Screen.MousePointer = 11
            i$ = REGLEIDO$("INVERT", RGI&)
            Call REPLA(i$, Chr$(96), 16, 1)
            Call GRAREG("INVERT", i$, RGI&)
            '
            CII% = CVI(Mid$(i$, 17, 2))
            If CII% > 0 Then
                If CII% <= NUMAS% Then
                    REGMA& = CII% + 1
                    x$ = REGLEIDO$("MASTER", REGMA&)
                    Call REPLA(x$, Chr$(96), 92, 1)
                    Call GRAREG("MASTER", x$, REGMA&)
                End If
            End If
            '
            GoSub 181                    ' eliminar de ECOMAST y CIACTI
            '
        End If
    End If
    Call CIERRARCH("INVERT")
    Call CIERRARCH("MASTER")
    Call GRACONTROL("", "AMASTO", "NORMAL")
    Screen.MousePointer = 1
    '
    GoTo 5
    '

125 If ULTREG&("INVERT") < 1 Then
        Call REGAPP("INVERT", String$(15, 0) + Chr$(1) + MKI$(0))
        Call CIERRARCH("INVERT")
    End If
    '
    NV& = ULTREG&("INVERT") - 1
    If CVI(Mid$(REGLEIDO$("INVERT", 1&), 17, 2)) <> NV& Then
        Call COMUNI("Posible Inconsistencia (1) !!!\Ejecute Opción de Control.")
        Exit Sub
    End If
    '
    ABRE$ = AJUSTI$(COD$, 15)
    LI& = 1: LS& = NV& + 2
    ENCU% = 0
        '
126 E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 128
    L& = LI& + E&: If L& < 2 Or L& > NV& + 1 Then GoTo 128
    i$ = REGLEIDO$("INVERT", L&): I1$ = Left$(i$, 15)
                                  If I1$ < ABRE$ Then LI& = L&: GoTo 126
                                  If I1$ > ABRE$ Then LS& = L&: GoTo 126
    ENCU% = 1: CI% = CVI(Mid$(i$, 17, 2)): RGI& = L&
               If Mid$(i$, 16, 1) = Chr$(96) Then
                   ENCU% = (-1)
               End If
    '
128 Return


181 FIECO% = FILENBR%("ECOMAST")
    Dim ECM As String * 64
    FICIA% = FILENBR%("CIACTI")
    Dim CCT As String * 2
    '
    CORR% = 0: FIN& = ULTREG&("ECOMAST")
    For IJ& = 1 To FIN&
        Get #FIECO%, IJ&, ECM$
        Get #FICIA%, IJ& + 1, CCT$
        If CORR% > 0 Then
            Put #FIECO%, IJ& - CORR%, ECM$
            Put #FICIA%, IJ& + 1 - CORR%, CCT$
        End If
        If TRIM$(Left$(ECM$, 15)) = TRIM$(COD$) Then CORR% = CORR% + 1
    Next IJ&
    '
    For CRRX% = 1 To CORR%
        ECM$ = String$(64, 0)
        Put #FIECO%, FIN& + 1 - CRRX%, ECM$
        CCT$ = String$(2, 0)
        Put #FICIA%, FIN& + 2 - CRRX%, CCT$
    Next CRRX%
    '
    Get #FICIA%, 1, CCT$
    CCT$ = MKI$(FIN& - CORR%)
    Put #FICIA%, 1, CCT$
    '
    Call CIERRARCH("ECOMAST"): FIECO% = 0
    Call CIERRARCH("CIACTI"): FICIA% = 0
    '
    Call FRESHECHO("ECOMAST")
    '
    Return
    '
End Sub
'


Sub LIESTRU()      ' listado estructuras y enlaces por codigo
    '
    CODACT$ = Formula.Text1.TEXT
    CIACT% = CODINT%(Formula.Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
        Exit Sub
    End If
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Formula.Refresh
    FIECO% = FILENBR%("LIESTRU")
    Dim ECM As String * 16
    '
    Call HEADERS("LIESTRU", "")
    Call HEADERS("LIESTRU", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LIESTRU", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
    '
    Screen.MousePointer = 11
    NESTRU& = ULTREG&("ESTRUNUE")
    j& = 0
    '
    Formula.List2.Clear
    Formula.List2.Visible = True
    For i& = 1 To NESTRU&
      CI% = CVI(Left$(REGLEIDO$("ESTRUNUE", i&), 2))
      ACTIX% = 0
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          CODYY$ = CODEXT$(CI%)
          If Mid$(CODYY$, 16, 1) <> Chr$(96) Then ' SUPRIME ITEMS DE BAJA
            Formula.List2.Refresh
            Formula.List2.AddItem CODYY$ + DESCRIT$(CI%)
            Formula.List2.ListIndex = Formula.List2.ListCount - 1
            If CODYY$ >= COD1$ Then
              If CODYY$ <= COD2$ Then
                Call LEEEXPLO(CODYY$, 1)
                If CAIT% > 0 Then
                  For K% = 1 To CAIT%
                    CI1% = CIJ%(K%)
                    CAN = USOI(K%)
                    PD% = PDI%(K%)
                    If CI1% > 0 Then
                      If CI1% <= NUMAS% Then
                        If Abs(CAN) > 0.0005 Then
                          UM$ = UNIMED$(CI1%)
                          Z$ = REGBLAN$("LIESTRU")
                          Z$ = REPLACE$(Z$, MKI$(CI1%), 3, 2)
                          Z$ = REPLACE$(Z$, UM$, 5, 2)
                          Z$ = REPLACE$(Z$, MKI$(CI%), 1, 2)
                          Z$ = REPLACE$(Z$, MKS$(CAN), 7, 4)
                          Z$ = REPLACE$(Z$, Chr$(PD%), 11, 1)
                          '
                          j& = j& + 1
                          Call GRAREG("LIESTRU", Z$, j&)
                          ACTIX% = 1
                          '
                        End If
                      End If
                    End If
                  Next K%
                End If
              End If
            End If
          End If
        End If
      End If
      '
      If ACTIX% > 0 Then
        Z$ = REGBLAN$("LIESTRU")
        Z$ = REPLACE$(Z$, MKI$(CI%), 1, 2)
        j& = j& + 1
        Call GRAREG("LIESTRU", Z$, j&)
      End If
      '
    Next i&
    '
    For i& = j& + 1 To ULTREG&("LIESTRU")
      Call GRAREG("LIESTRU", String$(16, 0), i&)
    Next i&
    '
    Call CIERRARCH("LIESTRU")
    '
    Screen.MousePointer = 1
    Call FINAL("*SOLIEST")
    Formula.List2.Visible = False
    Formula.Text1.TEXT = CODACT$
    On Error Resume Next
    Formula.Text1.SetFocus
    On Error GoTo 0
    '
End Sub
'

Sub LIMULTI()
    '
    CODACT$ = Formula.Text1.TEXT
    CIACT% = CODINT%(Formula.Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    Formula.Refresh
    Formula.List2.Clear
    Formula.List2.Visible = True
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        x$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(x$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  Formula.List2.Refresh
                  Formula.List2.AddItem CODYY$ + DESCRIT$(CI%)
                  Formula.List2.ListIndex = Formula.List2.ListCount - 1
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(x$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
    Formula.List2.Clear
    '
    For REG& = 1 To NUINV%
        '
        x$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(x$, 16)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(x$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOMUL(CI%)
                If CONTAMUL% > 1 Then
                  '
                  Formula.List2.Refresh
                  Formula.List2.AddItem CODYY$ + DESCRIT$(CI%)
                  Formula.List2.ListIndex = Formula.List2.ListCount - 1
                  '
                  If REGQ& > 0 Then
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIMULTI", String$(128, "="), REGQ&)
                  End If
                  '
                  For i% = 1 To CONTAMUL%
                    '
                    If i% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If CANT(i%) > 0.00005 Then  ' SUPRIME DESARMADO
                       XX$ = REGBLAN$("LIMULTI")
                       Call REPLA(XX$, MKI$(i% - 1), 1, 2)
                       Call REPLA(XX$, CODEXT$(CIK%(i%)), 3, 16)
                       Call REPLA(XX$, DESCRIT$(CIK%(i%)), 17 + 2 * NIVE%(i%), 48)
                       Call REPLA(XX$, CSTRING$(MKS$(CANT(i%)), 4, 10, 5, 2), 75, 10)
                       Call REPLA(XX$, CSTRING$(MKI$(PAD%(i%)), 1, 6, 0, 2), 85, 6)
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If i% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                 Next i%
                  '
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    For REGI& = REGQ& + 1 To ULTREG&("LIMULTI")
        Call GRAREG("LIMULTI", String$(96, 0), REGI&)
    Next REGI&
    Call CIERRARCH("LIMULTI")
    '
    Call HEADERS("LIMULTI", "")
    Screen.MousePointer = 1
    Call FINAL("*LIMULTI")
    'Call FINAL("")
    '
    Formula.List2.Visible = False
    Formula.Text1.TEXT = CODACT$
    'On Error Resume Next
    'FORMULA.Text1.SetFocus
    'On Error GoTo 0
    '
End Sub
'

Sub LIMULIN()
    '
    CODACT$ = Formula.Text1.TEXT
    CIACT% = CODINT%(Formula.Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         CI1% = PRICOD%
         CI2% = ULTCOD%
    End If
    x$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(x$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(x$, 2)): COD1$ = CODEXT$(CI1%): If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(x$, 3, 2)): COD2$ = CODEXT$(CI2%): If TRIM$(COD2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    Formula.Refresh
    Formula.List2.Clear
    Formula.List2.Visible = True
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        x$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(x$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  Formula.List2.Refresh
                  Formula.List2.AddItem CODYY$ + DESCRIT$(CI%)
                  Formula.List2.ListIndex = Formula.List2.ListCount - 1
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(x$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
    Formula.List2.Clear
    '
    j& = 0
    For REG& = 1 To NUINV%
        '
        x$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(x$, 16)
        If CODYY$ >= COD1$ Then
          If CODYY$ <= COD2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(x$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOFAI(CI%, 1, 1)
                If CONTAMUL% > 1 Then
                  '
                  Formula.List2.Refresh
                  Formula.List2.AddItem CODYY$ + DESCRIT$(CI%)
                  Formula.List2.ListIndex = Formula.List2.ListCount - 1
                  '
                  For K% = 1 To CONTAMUL%
                    CI1% = CIK%(K%)
                    CAN = CANT(K%)
                    PD% = 0 'PDI%(K%)
                    If CI1% > 0 Then
                      If CI1% <= NUMAS% Then
                        If Abs(CAN) > 0.0005 Then
                          UM$ = UNIMED$(CI1%)
                          Z$ = REGBLAN$("LIESTRU")
                          Z$ = REPLACE$(Z$, MKI$(CI1%), 3, 2)
                          Z$ = REPLACE$(Z$, UM$, 5, 2)
                          Z$ = REPLACE$(Z$, MKI$(CI%), 1, 2)
                          Z$ = REPLACE$(Z$, MKS$(CAN), 7, 4)
                          Z$ = REPLACE$(Z$, Chr$(PD%), 11, 1)
                          '
                          j& = j& + 1
                          Call GRAREG("LIESTRU", Z$, j&)
                          ACTIX% = 1
                          '
                        End If
                      End If
                    End If
                  Next K%
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    Call SETULTREG("LIESTRU", j&)
    Call CIERRARCH("LIESTRU")
    '
    Call HEADERS("LIESTRU", "")
    Screen.MousePointer = 1
    Call FINAL("*LIESTRU")
    '
    Formula.List2.Visible = False
    Formula.Text1.TEXT = CODACT$
    '
End Sub
'

Sub REVIMAS()                         ' revision del maestro
        '
        Screen.MousePointer = 11
        HAYINCO% = 0
        '
        Call BLOQARCH("MASTER")
        Call BLOQARCH("INVERT")
        Call BLOQARCH("INVENUE")
        '
        Dim CODISEL$(32767), CISEL%(32767)
        '
        If Left$(UCase$(EMPREAC$), 5) = "BOREN" Then
           ' ejecuta conversion de codigos
           FIN& = NUMAS%
           For i& = 1 To FIN&
              Call TRACE(20, i&, FIN&)
              x$ = REGLEIDO$("MASTER", i& + 1)
              CODES$ = Mid$(x$, 77, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(x$, CODIS$, 77, 16)
                 Call GRAREG("MASTER", x$, i& + 1)
              End If
           Next i&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For i& = 1 To FIN&
              Call TRACE(20, i&, FIN&)
              x$ = REGLEIDO$("INVERT", i& + 1)
              CODES$ = Mid$(x$, 1, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(x$, CODIS$, 1, 16)
                 Call GRAREG("INVERT", x$, i& + 1)
              End If
           Next i&
           Call CIERRARCH("INVERT")
           '
        End If
        '
        NM% = ULTREG&("MASTER") - 1
        If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
              x$ = REGLEIDO$("MASTER", 1&)
              Call REPLA(x$, MKI$(NM%), 1, 2)
              Call GRAREG("MASTER", x$, 1&)
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Indice Restaurado.")
              HAYINCO% = 1
          Else
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Control O.K.")
        End If
        '
        Screen.MousePointer = 11
        NV% = ULTREG&("INVERT") - 1
        FIN& = NV%
        For II& = 1 To NV%
            Call TRACE(20, II&, FIN&)
            x$ = REGLEIDO$("INVERT", II& + 1)
            COD$ = Left$(x$, 16): CII% = CVI(Mid$(x$, 17, 2))
            If Mid$(COD$, 16, 1) = Chr$(96) Then
               If CII% > 0 Then
                  If CII% <= NUMAS% Then
                     REGMA& = CII% + 1
                     Y$ = REGLEIDO$("MASTER", REGMA&)
                     If Mid$(Y$, 77, 15) = Left$(COD$, 15) Then
                        Call REPLA(Y$, Chr$(96), 92, 1)
                        Call GRAREG("MASTER", Y$, REGMA&)
                     End If
                  End If
               End If
            End If
        Next II&
        '
        AMASTOK.List1.Clear
        AMASTOK.SORTLIST.Clear
        FIN& = NM%
        For II& = 1 To NM%
            Call TRACE(20, II&, FIN&)
            x$ = REGLEIDO$("MASTER", II& + 1)
            If CVI(Left$(x$, 2)) <> II& Then
                Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                CIX% = II&
                Call REPLA(x$, MKI$(CIX%), 1, 2)
                Call GRAREG("MASTER", x$, II& + 1)
            End If
            COD$ = AJUSTI$(TRIM$(Mid$(x$, 77, 15)), 15) + Mid$(x$, 92, 1)
            ZZ$ = "A" + COD$ + CSTRING$(MKS$(II&), 3, 7, 0, 2)
            AMASTOK.SORTLIST.AddItem ZZ$
        Next II&
        '
        FIN& = AMASTOK.SORTLIST.ListCount
        For II& = 1 To AMASTOK.SORTLIST.ListCount
           Call TRACE(20, II&, FIN&)
           ZZ$ = AMASTOK.SORTLIST.LIST(II& - 1)
           Call AGRELISTA(ZZ$)
        Next II&
        '
        NV% = ULTREG&("INVERT") - 1
        If NV% < 1 Then
             Call REGAPP("INVERT", String$(16, 0) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = AMASTOK.List1.ListCount
        x$ = REGBLAN$("INVERT")
        Call REPLA(x$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", x$, 1)
        '
        j& = 1
        FIN& = NV% - 1
        For II& = 0 To AMASTOK.List1.ListCount - 1
            Call TRACE(20, II&, FIN&)
            AMASTOK.List1.ListIndex = II&
            CODISEL$(II&) = Mid$(AMASTOK.List1.TEXT, 2, 16)
            CISEL%(II&) = Val(Mid$(AMASTOK.List1.TEXT, 18))
            x$ = REGBLAN$("INVERT")
            Call REPLA(x$, CODISEL$(II&), 1, 16)
            Call REPLA(x$, MKI$(CISEL%(II&)), 17, 2)
            Call GRAREG("INVERT", x$, II& + 2)
        Next II&
        '
        Call CIERRARCH("*.*")
        '
        CODA$ = ""
        FIN& = NV%
        For II& = 1 To NV%
           Call TRACE(20, II&, FIN&)
           IK$ = REGLEIDO$("INVERT", II& + 1)
           CODL$ = Left$(IK$, 15)
           If CODL$ = CODA$ Then
                If TRIM$(CODL$) = "" Then GoTo 19
                REPLAI% = REPLAI% + 1
                CODNUE$ = TRIM$(CODL$) + "-" + Chr$(64 + REPLAI%)
                If Left$(CODNUE$, 1) = "-" Then CODNUE$ = "0" + CODNUE$
                RGM& = 1 + CVI(Mid$(IK$, 17, 2))
                M$ = REGLEIDO$("MASTER", RGM&)
                Call REPLA(M$, CODNUE$, 77, 16)
                Call GRAREG("MASTER", M$, RGM&)
                Call REPLA(IK$, CODNUE$, 1, 15)
                Call REPLA(IK$, Chr$(96), 16, 1)
                Call GRAREG("INVERT", IK$, II& + 1)
                Call COMUNI("Codigo '" + TRIM$(CODL$) + "'Duplicado.\  \Reemplazado por '" + TRIM$(CODNUE$) + "'.")
                Screen.MousePointer = 11
                HAYINCO% = 1
              Else
                REPLAI% = 0
           End If
           '
           If REPLAI% = 0 Then CODA$ = CODL$
19      Next II&
        '
        Call CIERRARCH("*.*")
        Call BLANARCH("INVENUE")
        '
17000   'Call BLIMESS("Actualizando Archivos de Validación ...")
        '
        'Call BLOQARCH("ECOMAST")
        FIECO% = FILENBR%("ECOMAST")
        Dim ECM As String * 64
        '
        'Call BLOQARCH("CIACTI")
        FICIA% = FILENBR%("CIACTI")
        Dim CCT As String * 2
        '
        NV% = ULTREG&("INVERT") - 1
        CCT$ = MKI$(NV%)
        Put #FICIA%, 1, CCT$
        '
        K% = 0
        FIN& = NV%
        For XL& = 1 To NV%
             Call TRACE(20, XL&, FIN&)
             IK$ = REGLEIDO$("INVERT", XL& + 1)
             If Mid$(IK$, 16, 1) <> Chr$(96) Then
               If TRIM$(Left$(IK$, 16)) <> "" Then
                 '
                 CI% = CVI(Mid$(IK$, 17, 2))
                 DE$ = DESCRIT$(CI%)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, Left$(IK$, 16), 1, 16)
                 Call REPLA(EC$, DE$, 17, 46)
                 '
                 K% = K% + 1: ECM$ = EC$: Put #FIECO%, K%, ECM$
                              CCT$ = MKI$(CI%): Put #FICIA%, K% + 1, CCT$
                 '
               End If
             End If
        Next XL&
        '
        CCT$ = MKI$(K%): Put #FICIA%, 1, CCT$
        '
        For K1% = K% + 1 To LOF(FIECO%) / 64
           ECM$ = String$(64, 0): Put #FIECO%, K1%, ECM$
           CCT$ = String$(2, 0): Put #FICIA%, K1% + 1, CCT$
        Next K1%
        '
        Call CIERRARCH("ECOMAST")
        Call CIERRARCH("CIACTI")
        FIECO% = 0: FICIA% = 0
        '
        Call GRACONTROL("", "AMASTO", "NORMAL")
        '
        Call CIERRARCH("*.*")
        Call LIBARCH("AMASTO")
        '
        Call FRESHECHO("MASTER")
        Call FRESHECHO("ECOMAST")
        '
        'Call BLIMESS("-")
        '
        Screen.MousePointer = 1
        If HAYINCO% = 0 Then
            Call COMUNI("No se han Detectado Inconsistencias\en la Revisión del Maestro de Items.")
        End If
        Exit Sub
        '
        '
12500   If ULTREG&("INVERT") < 1 Then
             Call REGAPP("INVERT", String$(15, 0) + Chr$(1) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = ULTREG&("INVERT") - 1
        If CVI(Mid$(REGLEIDO$("INVERT", 1&), 17, 2)) <> NV% Then
            Call COMUNI("Posible Inconsistencia (1) !!!\Ejecute Opción de Control.")
            Call FINAL("")    ' GoTo 2000
        End If
        '
        ABRE$ = AJUSTI$(COD$, 15)
        LI& = 1: LS& = NV% + 2
        ENCU% = 0
        '
12510   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 12580
        L& = LI& + E&: If L& < 2 Or L& > NV% + 1 Then GoTo 12580
        IK$ = REGLEIDO$("INVERT", L&): I1$ = Left$(IK$, 15)
                                  If I1$ < ABRE$ Then LI& = L&: GoTo 12510
                                  If I1$ > ABRE$ Then LS& = L&: GoTo 12510
        ENCU% = 1: CI% = CVI(Mid$(IK$, 17, 2)): RGI& = L&
                If Mid$(IK$, 16, 1) = Chr$(96) Then
                    ENCU% = (-1)
                End If
        '
12580   Return
        '
        '
End Sub

Sub LIMASTER(MODO%)
    '
    HOO% = HOY(HO$)
    If MODO% = 0 Then
        MODO% = ALTERNA%("Listado de Códigos y Atributos\Código y Descripción para Inventarios\Stock Mínimo y Lote Reposición")
        If MODO% < 1 Or MODO% > 3 Then
          Exit Sub
        End If
      ElseIf MODO% = 1 Then
        MODO% = ALTERNA%("Listado de Códigos y Atributos\Stock Mínimo y Lote Reposición")
        If MODO% < 1 Or MODO% > 2 Then
          Exit Sub
        End If
        If MODO% = 2 Then MODO% = 3
    End If
    '
    CI1% = PRICOD%
    CI2% = ULTCOD%
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    JJ& = 0
    If MODO% = 1 Then
        'Call BLANARCH("LIMASEL")
        Call HEADERS("LIMASEL", "")
        Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
        Call HEADERS("LIMASEL", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMASEL", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
      ElseIf MODO% = 2 Then
        'Call BLANARCH("LIMAINV")
        Call HEADERS("LIMAINV", "")
        Call HEADERS("LIMAINV", "Tipo de Material: " + TIMAT$)
        Call HEADERS("LIMAINV", "Desde Codigo: " + TRIM$(COD1$) + " - " + DESCRIT$(CI1%))
        Call HEADERS("LIMAINV", "Hasta Codigo: " + TRIM$(COD2$) + " - " + DESCRIT$(CI2%))
        Call HEADERS("LIMAINV", "Stock Nominal al: " + HO$ + " - " + Time$ + " hs.")
      ElseIf MODO% = 3 Then
        'Call BLANARCH("LIMALOT")
        Call HEADERS("LIMALOT", "")
        Call HEADERS("LIMALOT", "Material: " + TRIM$(TIMAT$) + " (" + TRIM$(COD1$) + " - " + TRIM$(COD2$) + ")")
      Else
        Screen.MousePointer = 1
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    j& = 0: FIN& = NUINV%
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        IV$ = REGLEIDO$("INVERT", i& + 1)
        If Mid$(IV$, 16, 1) = "`" Then GoTo 145          ' codigo de baja
        '
        COD$ = Left$(IV$, 16)
        If COD$ < COD1$ Then GoTo 145
        If COD$ > COD2$ Then GoTo 149
        '
        CI% = CVI(Mid$(IV$, 17, 2))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
                REGMAS& = CI% + 1
                M$ = REGLEIDO$("MASTER", REGMAS&)
                If TMT% > 0 Then
                    If Asc(Mid$(M$, 93, 1)) <> TMT% Then GoTo 145
                End If
                '
                If MODO% = 1 Then
                   ML$ = Left$(M$, 2) + Mid$(M$, 93, 36)
                   JJ& = JJ& + 1: Call GRAREG("LIMASEL", ML$, JJ&)
                 ElseIf MODO% = 2 Then
                   ML$ = MKI$(0) + "  " + MKS$(0) + Space$(56)
                   Call REGAPP("LIMAINV", ML$)
                   ML$ = Left$(M$, 2) + Mid$(M$, 109, 2)
                   ML$ = ML$ + MKS$(STOCKACT(CI%))
                   ML$ = ML$ + String$(62, ".")
                   JJ& = JJ& + 1: Call GRAREG("LIMAINV", ML$, JJ&)
                 ElseIf MODO% = 3 Then
                   ML$ = REGBLAN$("LIMALOT")
                   Call REPLA(ML$, Left$(M$, 2), 1, 2)
                   Call REPLA(ML$, MKS$(STOMIN(CI%)), 3, 4)
                   Call REPLA(ML$, MKS$(LOTESTAN(CI%)), 7, 4)
                   Call REPLA(ML$, MKS$(LOREPOMA(CI%)), 11, 4)
                   Call REPLA(ML$, MKS$(COSUNI(CI%)), 15, 4)
                   Call REPLA(ML$, MKI$(FECOSTO%(CI%)), 19, 2)
                   JJ& = JJ& + 1: Call GRAREG("LIMALOT", ML$, JJ&)
                End If
            End If
        End If
145 Next i&
    '
149 FIECO% = 0
    If MODO% = 1 Then
        Call SETULTREG("LIMASEL", JJ&)
        Call CIERRARCH("LIMASEL")
        Call FINAL("*LIMASEL")
      ElseIf MODO% = 2 Then
        Call SETULTREG("LIMAINV", JJ&)
        Call CIERRARCH("LIMAINV")
        Call FINAL("*LIMAINV")
      ElseIf MODO% = 3 Then
        Call SETULTREG("LIMALOT", JJ&)
        Call CIERRARCH("LIMALOT")
        Call FINAL("*LIMALOT")
    End If
    '
    Screen.MousePointer = 1
End Sub
'

Sub MUESTRAPLANO(CIIX%)
   REGPLA& = 0
   If CIIX% > 0 Then
      If CIIX% <= NUMAS% Then
         For U& = 1 To ULTREG&("IPLANOS")
            XX$ = REGLEIDO$("IPLANOS", U&)
            If TRIM$(Mid$(XX$, 3, 63)) <> "" Then
               UTILA$ = TRIM$(Mid$(XX$, 3, 63))
            End If
            If CVI(Left$(XX$, 2)) = CIIX% Then
               util$ = TRIM$(Mid$(XX$, 3, 63))
               Ruta$ = TRIM$(Mid$(XX$, 66, 63))
               REGPLA& = U&
               GoTo 10
            End If
         Next U&
         '
         util$ = UTILA$
         Ruta$ = ""
         '
10       XX$ = REGBLAN$("IPLANOS")
         Call REPLA(XX$, MKI$(CIIX%), 1, 2)
         Call REPLA(XX$, util$, 3, 63)
         Call REPLA(XX$, Ruta$, 66, 63)
         '
         YY$ = OBJPLA$("ACCEPLANOS", XX$)
         If YY$ = "" Then
            Exit Sub
         End If
         '
         util$ = TRIM$(Mid$(YY$, 3, 63))
         If util$ = "" Then GoTo 10
         If EXISTE%(util$) < 1 Then
            Call MENSERR(24, "Utilitario Inexistente")
            GoTo 10
         End If
         Ruta$ = TRIM$(Mid$(YY$, 66, 63))
         If Ruta$ = "" Then GoTo 10
         If EXISTE%(Ruta$) < 1 Then
            Call MENSERR(24, "Archivo no Encontrado")
            GoTo 10
         End If
         '
         On Error GoTo 90
         Shell util$ + " " + Ruta$, vbMaximizedFocus
         On Error GoTo 0
         '
         If REGPLA& > 0 Then
              Call GRAREG("IPLANOS", YY$, REGPLA&)
            Else
              Call REGAPP("IPLANOS", YY$)
         End If
         Call CIERRARCH("IPLANOS")
         '
      End If
   End If
   '
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   Call MENSERR(24, "Plano no Accesible - Consulte")
   '
End Sub

Sub ALTAIT(CODX$)
    '
    If VACONTRO% <> 0 Then Exit Sub
    VACONTRO% = 1
    '
    If TRIM$(CODX$) = "" Then
        Call MENSERR(21, "Formato de Código\no Válido - Repita")
        TANTA = Timer
        ALTAITEM.Text1.SetFocus
        Exit Sub
    End If
    '
    If CODINT%(CODX$) > 0 Then
        Call MENSERR(21, "Código Existente\en la Base de Datos.")
        TANTA = Timer
        ALTAITEM.Text1.SetFocus
        Exit Sub
    End If
    '
    CODD$ = AJUSTI$(CODX$, 15)
    For U& = 1 To AMASTOK.List1.ListCount
       If Mid$(AMASTOK.List1.LIST(U& - 1), 2, 15) = CODD$ Then
          If Mid$(AMASTOK.List1.LIST(U& - 1), 17, 1) = Chr$(96) Then GoTo 10
          Call MENSERR(24, "Alta Imposible\Código Existente.")
          TANTA = Timer
          ALTAITEM.Text1.SetFocus
          Exit Sub
       End If
    Next U&
    '
10  COD1$ = AJUSTI$(CODX$, 15) + Chr$(96)
    If CODINT%(COD1$) > 0 Then
        OPXX% = COMALTER%("Este Código Figura de Baja\en la Base de Datos\  \Puede ahora Revivirlo o Cancelar Operación de Alta\\Revivir\Cancelar")
        If OPXX% = 1 Then
            Call REVIVIR(CODX$)
        End If
        TANTA = Timer
        ALTAITEM.Text1.TEXT = ""
        ALTAITEM.Text1.SetFocus
        Exit Sub
    End If
    '
    VALICOD% = 1
    '
End Sub
'

Sub REVIVIR(CODD$)
    '
    For U& = 1 To AMASTOK.List1.ListCount
       XX$ = AMASTOK.List1.LIST(U& - 1)
       If TRIM$(Mid$(XX$, 2, 15)) = TRIM$(Left$(CODD$, 15)) Then
          Call REPLA(XX$, " ", 17, 1)
          AMASTOK.List1.RemoveItem U& - 1
          'AMASTOK.List1.AddItem XX$
          Call AGRELISTA(XX$)
          GoTo 10
       End If
    Next U&
    Exit Sub
    '
10  CII% = Val(Mid$(XX$, 18))
    If CII% > 0 Then
       If CII% <= NUMAS% Then
          REGMA& = CII% + 1
          x$ = REGLEIDO$("MASTER", REGMA&)
          Call REPLA(x$, " ", 92, 1)
          Call GRAREG("MASTER", x$, REGMA&)
          REMAS$ = x$
       End If
    End If
    '
End Sub
'

Sub SUSTICOX(COD1$, COD2$, SUCCESS%)
     '
     Screen.MousePointer = 11
     '
     Call BLOQARCH("ECOMAST")
     Call BLOQARCH("CIACTI")
     Call BLOQARCH("INVERT")
     Call BLOQARCH("MASTER")
     '
     CI% = CODINT%(COD1$)
     CODNUE$ = COD$
     SUCCESS% = 0
     NV& = NUINV%
     '
     CD1$ = AJUSTI$(COD1$, 16)
     CD2$ = AJUSTI$(COD2$, 16)
     SENTI% = 1: FIN& = NV& + 1
     '
     Screen.MousePointer = 11
     AMASTOK.List1.Clear
     Dim II As String * 18
     N1% = FILENBR%("INVERT")
     FIN& = NUINV%
     For U& = 1 To FIN&
        Call TRACE(1, U&, FIN&)
        Get #N1%, U& + 1, II$
        ZZ$ = "A" + Left$(II$, 16) + CSTRING$(Mid$(II$, 17, 2), 1, 7, 0, 2)
        'AMASTOK.List1.AddItem ZZ$
        Call AGRELISTA(ZZ$)
     Next U&
     '
     For U& = 1 To AMASTOK.List1.ListCount
        XX$ = AMASTOK.List1.LIST(U& - 1)
        If TRIM$(Mid$(XX$, 2, 15)) = TRIM$(CD1$) Then
           Call REPLA(XX$, CD2$, 2, 16)
           AMASTOK.List1.RemoveItem U& - 1
           'AMASTOK.List1.AddItem XX$
           Call AGRELISTA(XX$)
           GoTo 10
       End If
    Next U&
    Exit Sub
    '
10  CII% = Val(Mid$(XX$, 18))
    If CII% > 0 Then
       If CII% <= NUMAS% Then
          REGMA& = CII% + 1
          x$ = REGLEIDO$("MASTER", REGMA&)
          Call REPLA(x$, CD2$, 77, 16)
          Call GRAREG("MASTER", x$, REGMA&)
          REMAS$ = x$
       End If
    End If
    '
    FINVE% = FILENBR%("INVERT")
    Dim VL As String * 18
    FIECO% = FILENBR%("ECOMAST")
    Dim EL As String * 64
    FICIA% = FILENBR%("CIACTI")
    Dim CL As String * 2
    FIMAS% = FILENBR%("MASTER")
    Dim ML As String * 128
    '       '
    J1& = 1: J2& = 0: J3& = 1: ATX% = 0
    FIN& = AMASTOK.List1.ListCount
    For U& = 1 To FIN&
        Call TRACE(1, U&, FIN&)
        CODD$ = Mid$(AMASTOK.List1.LIST(U& - 1), 2, 16)
        CIIX% = Val(Mid$(AMASTOK.List1.LIST(U& - 1), 18))
        If CODD$ <> Left$(VL$, 16) Then ATX% = 1
        If CIIX% < 0 Then ATX% = 1
        J1& = J1& + 1
        If ATX% = 1 Then
           ZZ1$ = Space$(18)
           Call REPLA(ZZ1$, CODD$, 1, 16)
           Call REPLA(ZZ1$, MKI$(Abs(CIIX%)), 17, 2)
           LSet VL$ = ZZ1$: Put #FINVE%, J1&, VL$
        End If
        '
        If Mid$(CODD$, 16, 1) <> Chr$(96) Then
           If TRIM$(CODD$) <> "" Then
              J2& = J2& + 1
              J3& = J3& + 1
              If J2& > LOF(FIECO%) / 64 Then ATX% = 1
              '
              If ATX% = 1 Then
                 CI% = Abs(CIIX%)
                 REGMA& = CI% + 1
                 Get #FIMAS%, REGMA&, ML$
                 DENO$ = Mid$(ML$, 5, 64)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, MKI$(CI%), 63, 2)
                 Call REPLA(EC$, CODD$, 1, 16)
                 Call REPLA(EC$, DENO$, 17, 46)
                 '
                 LSet EL$ = EC$: Put #FIECO%, J2&, EL$
                 LSet CL$ = MKI$(CI%): Put #FICIA%, J3&, CL$
                 '
              End If
           End If
        End If
    Next U&
    '
    Screen.MousePointer = 1
    SUCCESS% = 1
    '
90  Call CIERRARCH("MASTER")
    Call CIERRARCH("INVERT")
    Call CIERRARCH("ECOMAST")
    Call CIERRARCH("CIACTI")
    '
End Sub
' ***************************************************************************
'
Sub AGRELISTA(RAGRE$)
   '
   ULIS& = AMASTOK.List1.ListCount
   If ULIS& < 1 Or Left$(RAGRE$, 17) > AMASTOK.List1.LIST(ULIS& - 1) Then
       AMASTOK.List1.AddItem RAGRE$
       GoTo 10
     Else
       For j& = ULIS& - 1 To 0 Step -1
         If Left$(AMASTOK.List1.LIST(j&), 17) <= Left$(RAGRE$, 17) Then
           AMASTOK.List1.AddItem RAGRE$, j& + 1
           GoTo 10
         End If
       Next j&
       AMASTOK.List1.AddItem RAGRE$, 0
10 End If
   '
End Sub
'
Sub IMPORNEU()
   '
   Dim TXX$(128)
   Dim STC$(8192)
   '
   LU$ = LUDAT$
   N1% = FILEOPEN(LU$ + "TRAIDOS\FGPART.DIF", 1, 0)
   If ULTREG&("MASTER") > 0 Then
        Call MENSERR(24, "Importacion Imposible.\Archivo Maestro existente")
        Exit Sub
      Else
        x$ = MKI$(1) + String$(126, 0)
        Call GRAREG("MASTER", x$, 1)
   End If
   '
10 Line Input #N1%, AA$
   If AA$ = "BOT" Then GoTo 20
   GoTo 10
   '
20 BB$ = "": JJ% = 0
25 If EOF(N1%) Then
      Call CIERRARCH("*.*")
      Call REVIMAS
      GoTo 50
   End If
   Line Input #N1%, AA$
   If AA$ = "BOT" Then
      GoTo 30
   End If
   If Left$(AA$, 1) <> Chr$(34) Then GoTo 25
   If Mid$(AA$, 2, 3) = "ART" Then GoTo 10
   CC$ = REPLACAR$(AA$, Chr$(34), "")
   JJ% = JJ% + 1
   TXX$(JJ%) = CC$
   GoTo 25
   '
30 COD$ = TXX$(1)
   DE$ = TRIM$(TXX$(3)) + " " + TRIM$(TXX$(4)) + TRIM$(TXX$(5))
   UNI$ = TRIM$(TXX$(13))
      If UNI$ = "" Then UNI$ = "U."
      If UNI$ = "1" Then UNI$ = "U."
   COU = Val(TXX$(21))
   If Val(TXX$(41)) > 0 Then COU = Val(TXX$(41))
   If Val(TXX$(15)) > 1 Then COU = COU / Val(TXX$(15))
   If Val(TXX$(12)) = 1 Then COU = COU / 100
   If Val(TXX$(12)) = 2 Then COU = COU / 1000
   MONEI% = 1
   '
   STMII% = 0: If Val(TXX$(27)) < 32000 Then STMII% = Val(TXX$(27))
   LOREI% = 0: If Val(TXX$(29)) < 32000 Then LOREI% = Val(TXX$(29))
   LOTEMAI% = 0: If Val(TXX$(28)) < 32000 Then LOTEMAI% = Val(TXX$(28))
   If LOTEMAI% < LOREI% Then LOTEMAI% = LOREI%
   '
   CI% = ULTREG&("MASTER")
   REMAS$ = REGBLAN$("MASTER")
   Call REPLA(REMAS$, MKI$(CI%), 1, 2)
   Call REPLA(REMAS$, DE$, 5, 64)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
   Call REPLA(REMAS$, String$(7, 0), 96, 7)
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
   Call REPLA(REMAS$, MKS$(COU), 111, 4)
   Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
   Call REPLA(REMAS$, MKS$(PPP), 117, 4)
   Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
   '
   Call REGAPP("MASTER", REMAS$)
   x$ = REGLEIDO$("MASTER", 1&)
   x$ = REPLACE$(x$, MKI$(CI%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   GoTo 20
   '
50 N1% = FILEOPEN(LU$ + "TRAIDOS\FGPEST.DIF", 1, 0)
   '
60 Line Input #N1%, AA$
   If AA$ = "BOT" Then GoTo 70
   GoTo 60
   '
70 BB$ = "": JJ% = 0
75 If EOF(N1%) Then
      Call CIERRARCH("*.*")
      GoTo 90
   End If
   Line Input #N1%, AA$
   If AA$ = "BOT" Then
      GoTo 80
   End If
   If Left$(AA$, 1) <> Chr$(34) Then GoTo 75
   'If Mid$(AA$, 2, 3) = "ART" Then GoTo 10
   CC$ = REPLACAR$(AA$, Chr$(34), "")
   JJ% = JJ% + 1
   TXX$(JJ%) = CC$
   GoTo 75
   '
80 COD1$ = TXX$(1)
   COD2$ = TXX$(2)
   Uso = Val(TXX$(4))
   CI1% = CODINT%(COD1$)
   If CI1% > 0 Then
      CI2% = CODINT%(COD2$)
      If CI2% > 0 Then
         If Abs(Uso) > 0 Then
            STC$(CI1%) = STC$(CI1%) + MKI$(CI2%) + MKS$(Uso) + Chr$(1) + Chr$(0)
         End If
      End If
   End If
   '
   GoTo 70
   '
90 U& = 0
   For i% = 1 To 8192
      If STC$(i%) <> "" Then
         STCX$ = STC$(i%) + String$(400, 0)
         R2$ = MKI$(i%) + Left$(STCX$, 120) + String$(6, 0)
         R4$ = Mid$(STCX$, 121, 256)
         U& = U& + 1
         Call GRAREG("ESTRUNUE", R2$, U&)
         Call GRAREG("ESTRUDOS", R4$, U&)
      End If
   Next i%
   '
   Call CIERRARCH("*.*")
End Sub


Sub IMPORMAI()
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\DESCRIP.PRN", 1, 0)
   Line Input #N1%, AA$
   '
10 If EOF(N1%) Then
      Close #N1%
      GoTo 19
   End If
   '
   Line Input #N1%, AA$
   COD$ = TRIM$(Mid$(AA$, 10, 16))
   DE$ = TRIM$(Mid$(AA$, 32))
   TIMA% = Asc(Left$(AA$, 1)) - 64
   UNID$ = "U."
   '
   CODX$ = TRIM$(COD$)
   For KKI% = 1 To CACOD%
      If CODCAR$(KKI%) = CODX$ Then GoTo 10
   Next KKI%
   CACOD% = CACOD% + 1
   CODCAR$(CACOD%) = CODX$
   '
   REMAS$ = REGBLAN$("MASTER")
   Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
   Call REPLA(REMAS$, DE$, 5, 64)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   Call REPLA(REMAS$, MKI$(0), 103, 2)
   Call REPLA(REMAS$, MKI$(0), 105, 2)
   Call REPLA(REMAS$, MKI$(0), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
        '
   REMA& = CACOD% + 1
   Call GRAREG("MASTER", REMAS$, REMA&)
   GoTo 10
   '
19 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\ARTIC.PRN", 1, 0)
   Line Input #N1%, AA$
   '
20 BB$ = "": JJ% = 0
21 If EOF(N1%) Then
      GoTo 29
   End If
   Line Input #N1%, AA$
23 COD$ = Left$(AA$, 16)
   DE$ = ""
   UNI$ = "U."
   COSUNIX$ = TRIM$(Mid$(AA$, 62, 12))
   Call REPLACAR(COSUNIX$, ",", ".")
   COU = Val(COSUNIX$)
   MONEI% = 1
   '
   STMII% = 0
   LOREI% = Val(Mid$(AA$, 29, 6))
   LOTEMAI% = Val(Mid$(AA$, 21, 6))
   FECO% = HOY(HO$)
   '
   CODX$ = TRIM$(COD$)
   For KKI% = 1 To CACOD%
      If CODCAR$(KKI%) = CODX$ Then GoTo 24
   Next KKI%
   CACOD% = CACOD% + 1
   CODCAR$(CACOD%) = CODX$
   DE$ = ""
   KKI% = CACOD%
   REMAS$ = REGBLAN$("MASTER")
   Call REPLA(REMAS$, MKI$(KKI%), 1, 2)
   Call REPLA(REMAS$, DE$, 5, 64)
   REMA& = KKI% + 1
   Call GRAREG("MASTER", REMAS$, REMA&)
   DE$ = ""
   '
24 REMA& = KKI% + 1
   REMAS$ = REGLEIDO$("MASTER", REMA&)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
   Call REPLA(REMAS$, String$(7, 0), 96, 7)
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
   Call REPLA(REMAS$, MKS$(COU), 111, 4)
   Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
   Call REPLA(REMAS$, MKS$(PPP), 117, 4)
   Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
   '
   Call GRAREG("MASTER", REMAS$, REMA&)
   '
26 GoTo 20
   '
29 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\SUBPR.PRN", 1, 0)
   Line Input #N1%, AA$
   '
30 BB$ = "": JJ% = 0
31 If EOF(N1%) Then
      GoTo 39
   End If
   Line Input #N1%, AA$
   TIMA% = Asc(Left$(AA$, 1)) - 64
33 COD$ = Mid$(AA$, 9, 15)
   DE$ = ""
   UNI$ = "U."
   COSUNIX$ = TRIM$(Mid$(AA$, 78, 12))
   Call REPLACAR(COSUNIX$, ",", ".")
   COU = Val(COSUNIX$)
   MONEI% = 1
   '
   STMII% = 0
   LOREI% = Val(Mid$(AA$, 39, 6))
   If Val(Mid$(AA$, 30, 6)) < 30000 Then
       LOTEMAI% = Val(Mid$(AA$, 30, 6))
     Else
       LOTEMAI% = (-1) * ((Val(Mid$(AA$, 30, 6))) / 100)
   End If
   FECO% = HOY(HO$)
   '
   CODX$ = TRIM$(COD$)
   For KKI% = 1 To CACOD%
      If CODCAR$(KKI%) = CODX$ Then GoTo 34
   Next KKI%
   CACOD% = CACOD% + 1
   CODCAR$(CACOD%) = CODX$
   DE$ = ""
   KKI% = CACOD%
   REMAS$ = REGBLAN$("MASTER")
   Call REPLA(REMAS$, MKI$(KKI%), 1, 2)
   Call REPLA(REMAS$, DE$, 5, 64)
   REMA& = KKI% + 1
   Call GRAREG("MASTER", REMAS$, REMA&)
   DE$ = ""
   '
34 REMA& = KKI% + 1
   REMAS$ = REGLEIDO$("MASTER", REMA&)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
   Call REPLA(REMAS$, String$(7, 0), 96, 7)
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
   Call REPLA(REMAS$, MKS$(COU), 111, 4)
   Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
   Call REPLA(REMAS$, MKS$(PPP), 117, 4)
   Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
   '
   Call GRAREG("MASTER", REMAS$, REMA&)
   '
36 GoTo 30
   '
39 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   '
40 Close #N1%
   REMA& = CACOD% + 1
   Call SETULTREG("MASTER", REMA&)
   Call CIERRARCH("*.*")
   Call REVIMAS
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\ARTICTD.PRN", 1, 0)
   Line Input #N1%, AA$
   '
52 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 55
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   CODPAR1$ = TRIM$(Mid$(AA$, 24, 15))
   COSUNIX$ = TRIM$(Mid$(AA$, 41, 10))
   Call REPLACAR(COSUNIX$, ",", ".")
   Uso = Val(COSUNIX$)
   '
   CI1% = CODINT%(CODCON$)
   CI2% = CODINT%(CODPAR1$)
   If CI1% > 0 Then
     If CI2% > 0 Then
       If Uso > 0 Then
         CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
         YACAR1%(CI1%) = 1
       End If
     End If
   End If
   '
   GoTo 52
   '
55 Close #N1%
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\SUBPRTD.PRN", 1, 0)
   Line Input #N1%, AA$
   '
62 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 65
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   CODPAR1$ = TRIM$(Mid$(AA$, 30, 15))
   COSUNIX$ = TRIM$(Mid$(AA$, 54, 10))
   Call REPLACAR(COSUNIX$, ",", ".")
   Uso = Val(COSUNIX$)
   '
   CI1% = CODINT%(CODCON$)
   CI2% = CODINT%(CODPAR1$)
   If CI1% > 0 Then
     If CI2% > 0 Then
       If Uso > 0 Then
         CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
         YACAR1%(CI1%) = 1
       End If
     End If
   End If
   '
   GoTo 62
   '
65 Close #N1%
   '
   N1% = FILEOPEN(LU$ + "TRAIDOS\ARTTD.PRN", 1, 0)
   Line Input #N1%, AA$
   '
66 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 68
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   CODPAR1$ = TRIM$(Mid$(AA$, 24, 15))
   COSUNIX$ = TRIM$(Mid$(AA$, 45, 10))
   Call REPLACAR(COSUNIX$, ",", ".")
   Uso = Val(COSUNIX$)
   '
   CI1% = CODINT%(CODCON$)
   CI2% = CODINT%(CODPAR1$)
   If CI1% > 0 Then
     If CI2% > 0 Then
       If Uso > 0 Then
         If YACAR1%(CI1%) = 0 Then
           CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
         End If
       End If
     End If
   End If
   '
   GoTo 66
   '
68 Close #N1%
   '
   j& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       x$ = MKI$(KKI%) + CODCAR$(KKI%) + String$(128, 0)
       j& = j& + 1
       Call GRAREG("ESTRUNUE", x$, j&)
     End If
  Next KKI%
  Call SETULTREG("ESTRUNUE", j&)
  Call CIERRARCH("*.*")
  '
71  N1% = FILEOPEN(LU$ + "TRAIDOS\ARTICTP.PRN", 1, 0)
   Line Input #N1%, AA$
   j& = 0
   '
72 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 75
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   NOPE% = 10 * Val(Mid$(AA$, 20, 2))
   DOPE$ = TRIM$(Mid$(AA$, 22, 29))
   MAQUI$ = TRIM$(Mid$(AA$, 51, 6))
   TIE = 60 * Val(Mid$(AA$, 70, 3)) + Val(Mid$(AA$, 73, 2)) + Val(Mid$(AA$, 76, 2)) / 60
   '
   CI1% = CODINT%(CODCON$)
   If CI1% > 0 Then
      x$ = REGBLAN$("SECOPER")
      Call REPLA(x$, MKI$(CI1%), 1, 2)
      Call REPLA(x$, MKI$(NOPE%), 3, 2)
      Call REPLA(x$, DOPE$, 5, 48)
      Call REPLA(x$, MAQUI$, 53, 6)
      Call REPLA(x$, MKS$(TIE), 67, 4)
      Call REPLA(x$, MKS$(1), 83, 4)
      j& = j& + 1
      Call GRAREG("SECOPER", x$, j&)
      YACAR2%(CI1%) = 1
   End If
   GoTo 72
   '
75 Close #N1%
   '
   
81  N1% = FILEOPEN(LU$ + "TRAIDOS\SUBPRTP.PRN", 1, 0)
   Line Input #N1%, AA$
   '
82 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 85
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   NOPE% = 10 * Val(Mid$(AA$, 25, 2))
   DOPE$ = TRIM$(Mid$(AA$, 27, 29))
   MAQUI$ = TRIM$(Mid$(AA$, 57, 6))
   TIE = 60 * Val(Mid$(AA$, 78, 3)) + Val(Mid$(AA$, 81, 2)) + Val(Mid$(AA$, 84, 2)) / 60
   '
   CI1% = CODINT%(CODCON$)
   If CI1% > 0 Then
      x$ = REGBLAN$("SECOPER")
      Call REPLA(x$, MKI$(CI1%), 1, 2)
      Call REPLA(x$, MKI$(NOPE%), 3, 2)
      Call REPLA(x$, DOPE$, 5, 48)
      Call REPLA(x$, MAQUI$, 53, 6)
      Call REPLA(x$, MKS$(TIE), 67, 4)
      Call REPLA(x$, MKS$(1), 83, 4)
      j& = j& + 1
      Call GRAREG("SECOPER", x$, j&)
      YACAR2%(CI1%) = 1
   End If
   GoTo 82
   '
85 Close #N1%
   '
91  N1% = FILEOPEN(LU$ + "TRAIDOS\ARTTP.PRN", 1, 0)
   Line Input #N1%, AA$
   '
92 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 95
   End If
   CODCON$ = TRIM$(Left$(AA$, 15))
   NOPE% = 10 * Val(Mid$(AA$, 21, 2))
   DOPE$ = TRIM$(Mid$(AA$, 23, 26))
   MAQUI$ = TRIM$(Mid$(AA$, 49, 6))
   TIE = 60 * Val(Mid$(AA$, 64, 3)) + Val(Mid$(AA$, 68, 2)) + Val(Mid$(AA$, 71, 2)) / 60
   '
   CI1% = CODINT%(CODCON$)
   If CI1% > 0 Then
     If YACAR2%(CI1%) = 0 Then
       x$ = REGBLAN$("SECOPER")
       Call REPLA(x$, MKI$(CI1%), 1, 2)
       Call REPLA(x$, MKI$(NOPE%), 3, 2)
       Call REPLA(x$, DOPE$, 5, 48)
       Call REPLA(x$, MAQUI$, 53, 6)
       Call REPLA(x$, MKS$(TIE), 67, 4)
       Call REPLA(x$, MKS$(1), 83, 4)
       j& = j& + 1
       Call GRAREG("SECOPER", x$, j&)
     End If
   End If
   GoTo 92
   '
95 Close #N1%
   '
121 N1% = FILEOPEN(LU$ + "TRAIDOS\CENTROP.PRN", 1, 0)
   Line Input #N1%, AA$
   j& = 0
   '
122 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 125
   End If
   CODMAQ$ = TRIM$(Left$(AA$, 6))
   DEMAQ$ = TRIM$(Mid$(AA$, 18, 33))
   COSUNIX$ = TRIM$(Mid$(AA$, 61, 10))
   Call REPLACAR(COSUNIX$, ",", ".")
   COSTO = Val(COSUNIX$)
   x$ = REGBLAN$("PADMAQ")
   Call REPLA(x$, CODMAQ$, 1, 6)
   Call REPLA(x$, DEMAQ$, 7, 30)
   Call REPLA(x$, MKS$(COSTO), 93, 4)
   j& = j& + 1
   Call GRAREG("PADMAQ", x$, j&)
   GoTo 122
   '
125 Call CIERRARCH("*.*")
    Close
    End

End Sub
'

Sub IMPORMAJ()
   '
Exit Sub
   LU$ = LUDAT$
   N1% = FILEOPEN(LU$ + "TRAIDOS\ARTIC.PRN", 1, 0)
   Line Input #N1%, AA$
   '
   Call BLANARCH("LIPRE001")
20 BB$ = "": JJ% = 0
21 If EOF(N1%) Then
      GoTo 29
   End If
   Line Input #N1%, AA$
23 COD$ = Left$(AA$, 16)
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         DE$ = DESCRIT0$(CI%)
         COSUT$ = TRIM$(Mid$(AA$, 62, 12))
         COSUNIX$ = REPLACAR$(COSUT$, ",", ".")
         COU = Val(COSUNIX$)
         PRE# = CDbl((Int(COU * 1.3 * 192 + 0.5)) / 100)
         COU = (Int(COU * 130 + 0.5)) / 100
         '
         REMA& = CI% + 1
         REMAS$ = REGLEIDO$("MASTER", REMA&)
         Call REPLA(REMAS$, MKS$(COU), 111, 4)
         Call GRAREG("MASTER", REMAS$, REMA&)
         '
         Z$ = REGBLAN$("LIPRE001")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, DE$, 3, 42)
         Call REPLA(Z$, MKD$(PRE#), 62, 8)
         Call REGAPP("LIPRE001", Z$)
         '
      End If
   End If
   GoTo 21
   '
29 Close #N1%
   N1% = FILEOPEN(LU$ + "TRAIDOS\SUBPR.PRN", 1, 0)
   Line Input #N1%, AA$
   '
30 BB$ = "": JJ% = 0
31 If EOF(N1%) Then
      GoTo 39
   End If
   Line Input #N1%, AA$
33 COD$ = Mid$(AA$, 9, 15)
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         COSUT$ = TRIM$(Mid$(AA$, 78, 12))
         COSUNIX$ = REPLACAR$(COSUT$, ",", ".")
         COU = Val(COSUNIX$)
         COU = (Int(COU * 130 + 0.5)) / 100
         '
         REMA& = CI% + 1
         REMAS$ = REGLEIDO$("MASTER", REMA&)
         Call REPLA(REMAS$, MKS$(COU), 111, 4)
         Call GRAREG("MASTER", REMAS$, REMA&)
      End If
   End If
   GoTo 31
   '
39 Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   '
End Sub
'

Sub IMPORMAK()
   '
   LU$ = LUDAT$
   N1% = FILEOPEN("A:\2.TXT", 1, 0)
   CI0% = NUMAS%
   Dim PRECIO#(16384)
   '
   j& = 0
   For U& = 1 To ULTREG&("LIPRE001")
      Z$ = REGLEIDO$("LIPRE001", U&)
      CI% = CVI(Left$(Z$, 2))
      If CODEXT$(CI%) <> "" Then
         j& = j& + 1
         Call GRAREG("LIPRE001", Z$, j&)
         If CI% > 0 Then
            If CI% <= 16384 Then
               If PRECIO#(CI%) < 0.005 Then
                  PRECIO#(CI%) = CVD(Mid$(Z$, 62, 8))
               End If
            End If
         End If
      End If
   Next U&
   Call SETULTREG&("LIPRE001", j&)
   '
   While Not EOF(N1%)
   Line Input #N1%, AA$
23 COD$ = TRIM$(Left$(AA$, 13))
   CI% = CODINT%(COD$)
   PRE# = CDbl(Val(Mid$(AA$, 58, 9)))
   If CI% < 1 Then
      CI0% = CI0% + 1
      CI% = CI0%
      DE$ = Mid$(AA$, 16, 41)
      REMA& = CI% + 1
      REMAS$ = String$(128, 0)
      Call REPLA(REMAS$, MKI$(CI%), 1, 2)
      Call REPLA(REMAS$, DE$, 5, 64)
      Call REPLA(REMAS$, String$(8, 0), 69, 8)
      Call REPLA(REMAS$, COD$, 77, 16)
      Call GRAREG("MASTER", REMAS$, REMA&)
   End If
   '
   If PRECIO#(CI%) < 0.005 Then
      PRECIO#(CI%) = PRE#
      Z$ = REGBLAN$("LIPRE001")
      Call REPLA(Z$, MKI$(CI%), 1, 2)
      Call REPLA(Z$, DE$, 3, 42)
      Call REPLA(Z$, MKD$(PRE#), 62, 8)
      Call REGAPP("LIPRE001", Z$)
   End If
   Wend
   '
29 Close #N1%
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   '
   Call REVIMAS
   '
End Sub

Sub IMPORGRI()
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
GoTo 51
   N1% = FILEOPEN(LU$ + "ARTICULO.PRN", 1, 0)
   Line Input #N1%, AA$
   '
   FIN& = 1.1 * LOF(N1%) / Len(AA$)
   IKJ& = 0
10 If EOF(N1%) Then
      Close #N1%
      GoTo 19
   End If
   '
   IKJ& = IKJ& + 1
   Call TRACE(20, IKJ&, FIN&)
   Line Input #N1%, AA$
   COD$ = TRIM$(Left$(AA$, 16))
   DE$ = TRIM$(Mid$(AA$, 20, 30)) + TRIM$(Mid$(AA$, 74))
   TIMA% = 1 ' Asc(Left$(AA$, 1)) - 64
   UNID$ = "U."
   '
   CODX$ = TRIM$(COD$)
   For KKI% = 1 To CACOD%
      If CODCAR$(KKI%) = CODX$ Then GoTo 10
   Next KKI%
   CACOD% = CACOD% + 1
   CODCAR$(CACOD%) = CODX$
   '
   REMAS$ = REGBLAN$("MASTER")
   Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
   Call REPLA(REMAS$, DE$, 5, 64)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   Call REPLA(REMAS$, MKI$(0), 103, 2)
   Call REPLA(REMAS$, MKI$(0), 105, 2)
   Call REPLA(REMAS$, MKI$(0), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
        '
   REMA& = CACOD% + 1
   Call GRAREG("MASTER", REMAS$, REMA&)
   GoTo 10
   '
19 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   '
   N1% = FILEOPEN(LU$ + "FORMULAS.PRN", 1, 0)
   Line Input #N1%, AA$
   '
   FIN& = 1.1 * LOF(N1%) / Len(AA$)
   KLJ& = 0
20 BB$ = "": JJ% = 0
21 If EOF(N1%) Then
      GoTo 29
   End If
   KLJ& = KLJ& + 1
   Call TRACE(20, KLJ&, FIN&)
   Line Input #N1%, AA$
23 COD$ = TRIM$(Mid$(AA$, 56, 16))
   DE$ = TRIM$(Mid$(AA$, 72, 39))
   UNI$ = TRIM$(UCase$(Mid$(AA$, 131, 4)))
   If UNI$ = "UNI" Then UNI$ = "U."
   COSUNIX$ = TRIM$(Mid$(AA$, 111, 9))
   COSUSS$ = REPLACAR$(COSUNIX$, ",", ".")
   COU = Val(COSUSS$)
   MONEI% = 1
   '
   STMII% = 0
   LOREI% = 0
   LOTEMAI% = 0
   FECO% = HOY(HO$)
   '
   CODX$ = TRIM$(COD$)
   For KKI% = 1 To CACOD%
      If CODCAR$(KKI%) = CODX$ Then GoTo 24
   Next KKI%
   GoTo 26    ' SALTEA IMPORTACION
   'CACOD% = CACOD% + 1
   'CODCAR$(CACOD%) = CODX$
   'DE$ = ""
   'KKI% = CACOD%
   'REMAS$ = REGBLAN$("MASTER")
   'Call REPLA(REMAS$, MKI$(KKI%), 1, 2)
   'Call REPLA(REMAS$, DE$, 5, 64)
   'REMA& = KKI% + 1
   'Call GRAREG("MASTER", REMAS$, REMA&)
   'DE$ = ""
   '
24 REMA& = KKI% + 1
   REMAS$ = REGLEIDO$("MASTER", REMA&)
   Call REPLA(REMAS$, String$(8, 0), 69, 8)
   Call REPLA(REMAS$, COD$, 77, 16)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
   Call REPLA(REMAS$, String$(7, 0), 96, 7)
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   Call REPLA(REMAS$, UNID$, 109, 2)
   Call REPLA(REMAS$, MKS$(COU), 111, 4)
   Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
   Call REPLA(REMAS$, MKS$(PPP), 117, 4)
   Call REPLA(REMAS$, MKI$(FULT%), 121, 2)
   Call REPLA(REMAS$, MKI$(NPRO%), 123, 2)
   Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
   Call REPLA(REMAS$, MKI$(FCP%), 127, 2)
   '
   Call GRAREG("MASTER", REMAS$, REMA&)
   '
26 GoTo 20
   '
29 x$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(x$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", x$, 1&)
   '
   REMA& = CACOD% + 1
   Call SETULTREG("MASTER", REMA&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call REVIMAS
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   N1% = FILEOPEN(LU$ + "FORMULAS.PRN", 1, 0)
   Line Input #N1%, AA$
   '
   FIN& = 1.1 * LOF(N1%) / Len(AA$)
   KLJ& = 0
   '
52 Line Input #N1%, AA$
   If EOF(N1%) Then
      Close #N1%
      GoTo 55
   End If
   KLJ& = KLJ& + 1
   Call TRACE(20, KLJ&, FIN&)
   CODCON$ = TRIM$(Left$(AA$, 15))
   CODPAR1$ = TRIM$(Mid$(AA$, 56, 15))
   COSUNIX$ = TRIM$(Mid$(AA$, 121, 7))
   COSUSS$ = REPLACAR$(COSUNIX$, ",", ".")
   Uso = Val(COSUSS$)
   '
   CI1% = CODINT%(CODCON$)
   CI2% = CODINT%(CODPAR1$)
   If CI1% > 0 Then
     If CI2% > 0 Then
       If Uso > 0 Then
         CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(Uso) + String$(2, 0)
         YACAR1%(CI1%) = 1
       End If
     End If
   End If
   '
   GoTo 52
   '
55 Close #N1%
   '
   j& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       x$ = MKI$(KKI%) + CODCAR$(KKI%) + String$(128, 0)
       j& = j& + 1
       Call GRAREG("ESTRUNUE", x$, j&)
     End If
  Next KKI%
  Call SETULTREG("ESTRUNUE", j&)
  Call CIERRARCH("*.*")
  '
65 Call CIERRARCH("*.*")
   Close
   End

End Sub
