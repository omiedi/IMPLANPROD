Attribute VB_Name = "EXPLO_FUNCTIONS"
Public CAIT As Integer
Public CIJ%(512), USOI(512), PDI%(512)
Public EXPLOLIN$(256)
'
Public NIVE%(1024), CIK%(1024), CANT(1024), PAD%(1024)
Public NIVE1%(1024), CII1%(1024), CANT1(1024), PAD1%(1024)
Public CONTAMUL%
'
Public FORMUCON%, SABOKIT%
Public PUNTCI As String
'
Public DESARMA%
Public REVIFORMU
'
Sub EXPLOMUL(CI%)
    '
    CONTAMUL% = 1
    NIVE%(CONTAMUL%) = 0
    CIK%(CONTAMUL%) = CI%
    CANT(CONTAMUL%) = 1
    PAD%(CONTAMUL%) = 0
    '
    NIVESEL% = 0
    '
10  J% = 0
    HAYEXPLO% = 0
    For I% = 1 To CONTAMUL%
        J% = J% + 1
        NIVE1%(J%) = NIVE%(I%)
        CII1%(J%) = CIK%(I%)
        CANT1(J%) = CANT(I%)
        PAD1%(J%) = PAD%(I%)
        If NIVE%(I%) = NIVESEL% Then
            CIBAS% = CIK%(I%)
            CABA = CANT(I%)
            Call AREXPLO(CIBAS%, CABA, 1)
            If CAIT% > 0 Then
                HAYEXPLO% = 1
                For K% = 1 To CAIT%
                    J% = J% + 1
                    NIVE1%(J%) = NIVESEL% + 1
                    CII1%(J%) = CIJ%(K%)
                    CANT1(J%) = CABA * USOI(K%)
                    PAD1%(J%) = PAD%(I%) + PDI%(K%)
                Next K%
            End If
        End If
    Next I%
    '
    CONTAMUL% = J%
    For I% = 1 To CONTAMUL%
        NIVE%(I%) = NIVE1%(I%)
        CIK%(I%) = CII1%(I%)
        CANT(I%) = CANT1(I%)
        PAD%(I%) = PAD1%(I%)
    Next I%
    '
    If HAYEXPLO% = 1 Then
        NIVESEL% = NIVESEL% + 1
        If NIVESEL% <= 9 Then
            GoTo 10
        End If
    End If
    '
End Sub
'

Sub LEEEXPLO(COD$, CABAS)      ' GENERACION AUTOMATICA ARCHIVO DE EXPLOSION LINEAL
    '
    CAIT% = 0
    CABA = CABAS: If CABA = 0 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    If CIBAS% > 0 Then
        Call AREXPLO(CIBAS%, CABA, 1)
    End If
    '
End Sub
'

Sub LEEEXPLOMUL(COD$, CABAS, MODO%)     ' GENERACION AUTOMATICA ARCHIVO DE EXPLOSION LINEAL
    '
    ' MODO%=2 fuerza explosion multilineal para SABO
    '
    Static EPAC$
    If EPAC$ = "" Then EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    If FORMUCON% = 0 Then
       FORMUCON% = 1
       If MODO% = 2 Or CONTROL$("", "FORMUCON") = "MULTILIN" Then
           FORMUCON% = 2
       End If
       SABOKIT% = 0: If MODO% = 2 Then SABOKIT% = 1
    End If
    '
    CAIT% = 0
    CABA = CABAS: If CABA < 0.5 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    If CIBAS% > 0 Then
      If FORMUCON% = 2 Then
          If InStr(EPAC$, "WORK") > 0 Then
              Call EXPLOCOM(CIBAS%, CABA)
              CAIT% = CONTAMUL%
              For KKII% = 1 To CAIT%
                CIJ%(KKII%) = CIK%(KKII%)
                USOI(KKII%) = CANT(KKII%) / CABA
              Next KKII%
            Else
              Call EXPLOFAI(CIBAS%, CABA, SABOKIT%)
          End If
        Else
          Call AREXPLO(CIBAS%, CABA, 1)
      End If
    End If
    '
End Sub
'

Sub EXPLOFAI(CIBASI%, CABASI, SABOKIT%)
    '
    Static CARTI&, ARTIPRO%(16384)
    '
    If CARTI& = 0 Then
        CARTI& = ULTREG&("ARTIPRO")
        If CARTI& < 1 Then
            CARTI& = (-1)
            GoTo 29
        End If
        '
        For JI& = 1 To CARTI&
           XI$ = REGLEIDO$("ARTIPRO", JI&)
           CIXK% = CVI(Left$(XI$, 2))
           If CIXK% > 0 Then
              If CIXK% <= NUMAS% Then
                 ARTIPRO%(CIXK%) = CVI(Mid$(XI$, 3, 2))
              End If
           End If
        Next JI&
    End If
    '
29  CONTAMUL% = 1
    NIVE%(CONTAMUL%) = 0
    CIK%(CONTAMUL%) = CIBASI%
    CANT(CONTAMUL%) = 1
    PAD%(CONTAMUL%) = 0
    '
    If CABASI <= 0 Then CABASI = 1
    '
    'CCCC$ = CODEXT$(CIBASI%)
    NIVESEL% = 0
    '
10  J% = 0
    HAYEXPLO% = 0
    For I% = 1 To CONTAMUL%
        J% = J% + 1
        NIVE1%(J%) = NIVE%(I%)
        CII1%(J%) = CIK%(I%)
        CANT1(J%) = CANT(I%)
        PAD1%(J%) = PAD%(I%)
        If NIVE%(I%) = NIVESEL% Then
          If ARTIPRO%(CIK%(I%)) < 1 Or I% = 1 Then
            CIBAS% = CIK%(I%)
            '
            If SABOKIT% = 1 Then
               CII7% = CIBAS%
               If InStr(UCase$(DESCRIT$(CII7%)), "KIT") <> 1 Then GoTo 19
            End If
            '
            If I% > 1 Then
              If PROVHABI%(CIK%(I%)) > 0 Then GoTo 19
            End If
            '
            CABA = CANT(I%)
            Call AREXPLO(CIBAS%, CABA, 1)
            If CAIT% > 0 Then
                HAYEXPLO% = 1
                For K% = 1 To CAIT%
                    J% = J% + 1
                    NIVE1%(J%) = NIVESEL% + 1
                    CII1%(J%) = CIJ%(K%)
                    CANT1(J%) = CABA * USOI(K%)
                    PAD1%(J%) = PAD%(I%) + PDI%(K%)
                Next K%
            End If
          End If
        End If
19  Next I%
    '
    CONTAMUL% = J%
    For I% = 1 To CONTAMUL%
        NIVE%(I%) = NIVE1%(I%)
        CIK%(I%) = CII1%(I%)
        CANT(I%) = CANT1(I%)
        PAD%(I%) = PAD1%(I%)
    Next I%
    '
    If HAYEXPLO% = 1 Then
        NIVESEL% = NIVESEL% + 1
        If NIVESEL% <= 9 Then
            GoTo 10
        End If
    End If
    '
    J% = 0
    For I% = 2 To CONTAMUL%
       If I% = CONTAMUL% Or NIVE%(I% + 1) <= NIVE%(I%) Then
          J% = J% + 1
          NIVE%(J%) = NIVE%(I%)
          If NIVE%(J%) > 1 Then NIVE%(J%) = 1
          CIK%(J%) = CIK%(I%)
          CANT(J%) = CANT(I%) * CABASI
          PAD%(J%) = PAD%(I%)
          '
          CIJ%(J%) = CIK%(J%)
          USOI(J%) = CANT(J%) / CABASI
          PDI%(J%) = PAD%(J%)
          '
       End If
    Next I%
    '
    CAIT% = J%
    CONTAMUL% = J%
    '
End Sub
'

Sub EXPLOCOM(CIBASI%, CABASI)
   '
   Static DP%(32767), FESTRU#
   '
   CI% = CIBASI%
   If CI% < 1 Then Exit Sub '
   '
   CODD$ = CODEXT$(CI%)
   CICON% = CI%
   '
   NVI% = NUINV%
   NMM% = NUMAS%
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
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
   Call EXPLOMUL(CI%)
   '
   FIN& = CONTAMUL%
   KKJ% = 0
   For KKI% = 1 To CONTAMUL
     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
        '
        KKJ% = KKJ% + 1
        CIK%(KKJ%) = CIK%(KKI%)
        CANT(KKJ%) = CANT(KKI%)
        '
      Else
        '
        If COSUNI(CIK%(KKI%)) > 0 Then
'          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
'             If PROVINT%(PHB%) = 0 Then
               KKJ% = KKJ% + 1
               CIK%(KKJ%) = CIK%(KKI%)
               CANT(KKJ%) = CANT(KKI%)
'             End If
           End If
        End If
        '
     End If
   Next KKI%
   '
   For KKK% = 1 To KKJ%
     For KKL% = 1 To KKK% - 1
       If CIK%(KKL%) = CIK%(KKK%) Then
         CANT(KKL%) = CANT(KKL%) + CANT(KKK%)
         CANT(KKK%) = 0
         GoTo 19
       End If
     Next KKL%
19 Next KKK%
   '
   CONTAMUL% = 0
   For KKK% = 1 To KKJ%
     If CANT(KKK%) >= 0.000005 Then
       CONTAMUL% = CONTAMUL% + 1
       CIK%(CONTAMUL%) = CIK%(KKK%)
       CANT(CONTAMUL%) = CANT(KKK%)
     End If
   Next KKK%
   '
End Sub
'
Sub EXPLOGR(COD$, CABAS)
    '
    CABA = CABAS: If CABA = 0 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    Call AREXPLO(CIBAS%, CABA, 2)
    '
End Sub
'

Sub AREXPLO(CIBAS%, CABA, OPCI%)
    '
    If PUNTCI$ = "" Then
        If ULTREG&("ESTRUNUE") < 1 Then
            If ULTREG&("BKSTRUC") > 0 Then
                Call RECUSTRUC
            End If
        End If
        '
        SMPA = Screen.MousePointer
        Screen.MousePointer = 11
        NM% = NUMAS%
        FIN& = ULTREG&("ESTRUNUE")
        For I& = 1 To FIN&
            X$ = REGLEIDO$("ESTRUNUE", I&)
            CI% = CVI(Left$(X$, 2) + String$(2, 0))
            PUNTCI$ = PUNTCI$ + MKI$(CI%)
        Next I&
        '
        Screen.MousePointer = SMPA
    End If
    '
    ENCU% = 0: PX1% = 1
500 PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
    If PX% < PX1% Then GoTo 600
    If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
    ENCU% = (PX% + 1) / 2
    '
600 If OPCI% = 1 Then
        VUELTX% = 0
        
610     CAIT% = 0
        REG& = ENCU%
        REVIFORMU = 0
        If REG& > 0 Then
            X$ = REGLEIDO$("ESTRUNUE", REG&)
            DPS$ = Mid$(X$, 3, 120)
            REVIFORMU = CVI(Mid$(X$, 123, 2))
            If ULTREG&("ESTRUDOS") >= REG& Then
                 DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
            End If
            '
            RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CIBAS%))
            If Len(RFF$) >= 4 Then
              RELA& = CVS(RFF$)
              If RELA& > 0 Then
                If RELA& <= ULTREG&("ESTRULAR") Then
                  DPS$ = DPS$ + Mid$(REGLEIDO$("ESTRULAR", RELA&), 3)
                End If
              End If
            End If
            '
            While Len(DPS$) Mod 8 <> 0
              DPS$ = DPS$ + Chr$(0)
            Wend
            '
            For K% = 1 To Len(DPS$) Step 8
              DPSX$ = Mid$(DPS$, K%, 8)
              CIXXII% = CVI(Left$(DPSX$, 2))
              If CIXXII% > 0 Then
                If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
                  CODXX$ = CODEXT$(CIXXII%)
                  If TRIM$(CODXX$) <> "" Then
                    CAIT% = CAIT% + 1
                    CIJ%(CAIT%) = CIXXII%
                    USOI(CAIT%) = CVS(Mid$(DPSX$, 3, 4))
                    PDI%(CAIT%) = Asc(Mid$(DPSX$, 7, 1))
                    '
                    DEXX$ = DESCRIT$(CIJ%(CAIT%))
                    UNXX$ = UNIMED$(CIJ%(CAIT%))
                    USON = CABA * USOI(CAIT%)
                    USOX$ = CSTRING$(MKS$(USON), 4, 12, 4, 2)
                    PDX$ = CSTRING$(Mid$(DPSX$, 7, 1), 1, 4, 0, 2)
                    XXX$ = Space$(64)
                    Call REPLA(XXX$, CODXX$, 1, 16)
                    Call REPLA(XXX$, DEXX$, 17, 28)
                    Call REPLA(XXX$, UNXX$, 47, 2)
                    Call REPLA(XXX$, USOX$, 49, 12)
                    Call REPLA(XXX$, PDX$, 61, 4)
                    EXPLOLIN(CAIT%) = XXX$
                  End If
                End If
              End If
            Next K%
        End If
        '
        If CAIT% = 15 Then
           If VUELTX% = 0 Then
              VUELTX% = 1
              For U& = 1 To ULTREG&("BKSTRUC")
                 X$ = REGLEIDO$("BKSTRUC", U&)
                 If CVI(Left$(X$, 2)) = CIBAS% Then
                    X2$ = Mid$(X$, 129, 256)
                    Call GRAREG("ESTRUDOS", X2$, REG&)
                    '
                    X1$ = REGLEIDO$("ESTRUNUE", REG&)
                    X3$ = X1$ + X2$
                    Call GRAREG("BKSTRUC", X3$, U&)
                    Call CIERRARCH("BKSTRUC")
                    '
                    GoTo 610
                 End If
              Next U&
           End If
        End If
        '
        Exit Sub
        '
     ElseIf OPCI% = 2 Then
        '
        For UK& = 1 + ULTREG&("ESTRUDOS") To ULTREG&("ESTRUNUE")
          Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), UK&)
        Next UK&
        '
        DPS$ = ""
        If CAIT% > 111 Then
          Call MENSERR(24, "Imposible Registrar.\Formula Excede Maximo Admisible\de 111 Componentes por Nivel.")
        End If
        '
        For I& = 1 To CAIT%
           XXX$ = EXPLOLIN(I&)
           CODXX$ = Mid$(XXX$, 1, 16)
           USON = Val(Mid$(XXX$, 49, 12)) / CABA
           PDII% = Val(Mid$(XXX$, 61, 4))
           If PDII% > 255 Then PDII% = 0
           CIII% = CODINT%(CODXX$)
           If Abs(USON) > 0 Then
             If CIII% > 0 Then
               If CIII% <= NUMAS% Then
                 DPSX$ = MKI$(CIII%) + MKS$(USON) + Chr$(PDII%) + Chr$(0)
                 DPS$ = DPS$ + DPSX$
               End If
             End If
           End If
           '
        Next I&
        '
        If ENCU% < 1 And Len(DPS$) < 8 Then Exit Sub
        '
        While Len(DPS$) < 376: DPS$ = DPS$ + MKI$(0) + MKS$(0) + Chr$(0) + Chr$(0): Wend
        '
        REG& = ENCU%
           If REG& < 1 Then
               REG& = 1 + ULTREG&("ESTRUNUE")
               Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), REG&)
               PUNTCI$ = PUNTCI$ + MKI$(CIBAS%)
           End If
           '
        DPS0$ = Left$(DPS$, 120) + String$(126, 0)
        RVFF% = HOY(HOS$)
        Call REPLA(DPS0$, MKI$(RVFF%), 121, 2)
        Call GRAREG("ESTRUNUE", MKI$(CIBAS%) + Left$(DPS0$, 126), REG&)
        DPS1$ = Mid$(DPS$, 121, 256)
        DPS3$ = Mid$(DPS$, 377)
        Call REPLA(DPS1$, Chr$(1), 256, 1)
        ATXXX% = 0: REGXXX& = REG&
        If REGLEIDO$("ESTRUDOS", REGXXX&) <> DPS1$ Then ATXXX% = 1
        Call GRAREG("ESTRUDOS", DPS1$, REG&)
        '
        RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CIBAS%))
        If DPS3$ <> "" Or RFF$ <> "" Then
          DPS3$ = MKI$(CIBAS%) + DPS3$ + String$(512, 0)
          RFF$ = RECFILT$("ESTRULAR", 1, MKI$(CIBAS%))
          RELAR& = 1 + ULTREG&("ESTRULAR")
          If Len(RFF$) >= 4 Then
            RELA& = CVS(RFF$)
            If RELA& > 0 Then
              If RELA& < RELAR& Then
                RELAR& = RELA&
              End If
            End If
          End If
          Call GRAREG("ESTRULAR", DPS3$, RELAR&)
          Call CIERRARCH("ESTRULAR")
        End If
        '
        If ATXXX% <> 0 Then
          For U& = 1 To ULTREG&("BKSTRUC")
            X$ = REGLEIDO$("BKSTRUC", U&)
            If CVI(Left$(X$, 2)) = CIBAS% Then
              X1$ = REGLEIDO$("ESTRUNUE", REG&)
              X3$ = X1$ + DPS1$
              Call GRAREG("BKSTRUC", X3$, U&)
              GoTo 999
            End If
          Next U&
        End If
        '
        If ATENTI% <> 12345 Then   ' MARCA PARA NO REORGANIZAR SI VIENE DE IMPORTACION
          If ULTREG&("IMPLOSI") < 1 Then
            Screen.MousePointer = 11
            JJ& = 0
            NESTRU& = ULTREG&("ESTRUNUE")
            VECON$ = "": VELEM$ = ""
            For I& = 1 To NESTRU&
              Call TRACE(20, I&, NESTRU&)
              CIXXI% = CVI(Left$(REGLEIDO$("ESTRUNUE", I&), 2))
              If CIXXI% > 0 Then
                If CIXXI% <= NUMAS% Then
                  Call AREXPLO(CIXXI%, 1, 1)
                  If CAIT% > 0 Then
                    For K% = 1 To CAIT%
                      VECON$ = MKI$(CIXXI%) + MKI$(CIJ%(K%)) + MKS$(USOI(K%))
                      JJ& = JJ& + 1
                      Call GRAREG("IMPLOSI", VECON$, JJ&)
                    Next K%
                  End If
                End If
              End If
            Next I&
            Call CIERRARCH("IMPLOSI")
            Call BAJALDISCO
            Call COMUNI("FLEXOFT ha Re-Organizado\la Base de Datos de Fórmulas.")
            Call CONECRUN(App.EXEName, "")
          End If
        End If
        '
        ' actualiza archivo de implosión IMPLOSI.DAT
        URN& = ULTREG&("IMPLOSI")
        RNC$ = RECFILT$("IMPLOSI", 1, MKI$(CIBAS%))
        KK& = (-3)
        '
        For I& = 1 To CAIT%
           XXX$ = EXPLOLIN(I&)
           CODXX$ = Mid$(XXX$, 1, 16)
           USON = Val(Mid$(XXX$, 49, 12)) / CABA
           CIII% = CODINT%(CODXX$)
           TTYY$ = MKI$(CIBAS%) + MKI$(CIII%) + MKS$(USON)
           KK& = KK& + 4
           If KK& <= Len(RNC$) - 3 Then
                 RENOTA& = CVS(Mid$(RNC$, KK&, 4))
                 XY$ = REGLEIDO$("IMPLOSI", RENOTA&)
                 If CVI(Left$(XY$, 2)) <> CIBAS% Then
                    Call MENSERR(24, "Error de Grabación de Archivo de Implosión.")
                    URN& = URN& + 1
                    RENOTA& = URN&
                 End If
               Else
                 URN& = URN& + 1
                 RENOTA& = URN&
           End If
           Call GRAREG("IMPLOSI", TTYY$, RENOTA&)
        Next I&
        '
        While KK& < Len(RNC$) - 3
           KK& = KK& + 4
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
           Call GRAREG("IMPLOSI", String$(8, 0), RENOTA&)
        Wend
        Call CIERRARCH("IMPLOSI")
        '
    End If
    '
999 End Sub
'
Function MAXIFAB(CIXI%)
   MFB = 0
   If CIXI% > 0 Then
     Call AREXPLO(CIXI%, 1, 1)
     If CAIT% > 0 Then
       MFB = 9999999: HAYEX% = 0
       For KII% = 1 To CAIT%
         If USOI(KII%) > 0.0001 Then
           HAYEX% = 1
           STAAX = STOCKDIS(CIJ%(KII%))
           If STAAX < 0 Then STAAX = 0
           On Error Resume Next
           MFC = STAAX / USOI(KII%)
           On Error GoTo 0
           If MFC < MFB Then MFB = MFC
         End If
       Next KII%
       MFB = MFB * HAYEX%
     End If
   End If
   '
   If Left$(UCase$(UNIMED$(CIXI%)), 1) = "U" Then
     MFB = Int(MFB)
   End If
   MAXIFAB = MFB
   '
End Function

Sub COPYEXPLO()
    '
100 CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
    If Len(CICJ$) < 4 Then
        Exit Sub
    End If
    '
    CODORI% = CVI(Left$(CICJ$, 2))
    CODEST% = CVI(Mid$(CICJ$, 3, 2))
    CICJ0$ = CICJ$
    '
    CODOR$ = TRIM$(CODEXT$(CODORI%))
    CODES$ = TRIM$(CODEXT$(CODEST%))
    '
    If CODOR$ = "" Then
        Call MENSERR(24, "Codigo Origen Inexistente")
        GoTo 100
    End If
    '
    If CODES$ = "" Then
        Call MENSERR(24, "Codigo Destino Inexistente")
        GoTo 100
    End If
    '
    Screen.MousePointer = 11
    Call LEEEXPLO(CODOR$, 1)
    If CAIT% < 1 Then
        Screen.MousePointer = 1
        OPX% = COMALTER%("Copia Imposible.\Código Origen " + TRIM$(CODOR$) + " sin Estructura.\\Repetir\Abandonar")
        If OPX% = 1 Then GoTo 100
        Exit Sub
    End If
    '
    Call LEEEXPLO(CODES$, 1)
    If CAIT% > 0 Then
        Screen.MousePointer = 1
        OPX% = COMALTER%("ATENCION: Código Destino " + TRIM$(CODES$) + " con Estructura.\\Cancelar\Sobre-escribir")
        If OPX% <> 2 Then GoTo 100
    End If
    '
    Screen.MousePointer = 11
    Call CIERRARCH("*.*")
    LU$ = LUDAT$
    '
    J& = 0
    PUNTCI$ = ""
    ESP1$ = "": ESP2$ = ""
    For I& = 1 To ULTREG&("ESTRUNUE")
        EST1$ = REGLEIDO$("ESTRUNUE", I&)
        CIV% = CVI(Left$(EST1$, 2))
        EST2$ = REGLEIDO$("ESTRUDOS", I&)
        If CIV% > 0 Then
           If CIV% <= NUMAS% Then
              If CIV% <> CODEST% Then
                 J& = J& + 1
                 Call GRAREG("ESTRUNUE", EST1$, J&)
                 Call GRAREG("ESTRUDOS", EST2$, J&)
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
    Next I&
    '
    Call REPLA(ESP1$, MKI$(CODEST%), 1, 2)
    EST1$ = ESP1$
    EST2$ = ESP2$
    J& = J& + 1
    Call GRAREG("ESTRUNUE", EST1$, J&)
    Call GRAREG("ESTRUDOS", EST2$, J&)
    PUNTCI$ = PUNTCI$ + MKI$(CODEST%)
    '
    For I& = J& + 1 To ULTREG&("ESTRUNUE")
        Call GRAREG("ESTRUNUE", String$(128, 0), I&)
        Call GRAREG("ESTRUDOS", String$(128, 0), I&)
    Next I&
    '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    FORMULA.Text1.TEXT = CODES$
    FORMULA.Text1.SetFocus
    '
End Sub

Sub BKSTRUC()
    '
    Screen.MousePointer = 11
    Dim CIBK%(16384)
    CABK% = 0
    For U& = 1 To ULTREG&("BKSTRUC")
        X$ = REGLEIDO$("BKSTRUC", U&)
        CIKJ% = CVI(Left$(X$, 2))
        If CIKJ% > 0 Then
            If CIKJ% <= 16384 Then
                CIBK%(CIKJ%) = 1
            End If
        End If
    Next U&
    '
    LOBK% = RECLEN%("BKSTRUC")
    '
    For U& = 1 To ULTREG&("ESTRUNUE")
        X$ = REGLEIDO$("ESTRUNUE", U&)
        If Len(X$) = 128 Then
            CIKJ% = CVI(Left$(X$, 2))
            If CIKJ% > 0 Then
                If CIKJ% <= 16384 Then
                    If CIBK%(CIKJ%) < 1 Then
                        U1& = U&
                        X$ = X$ + REGLEIDO$("ESTRUDOS", U1&)
                        X$ = Left$(X$ + String$(512, 0), 512)
                        Call REGAPP("BKSTRUC", X$)
                        CIBK%(CIKJ%) = 1
                    End If
                End If
            End If
        End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'

Sub RECUSTRUC()
    '
    If ULTREG&("ESTRUNUE") < 1 Then
      Screen.MousePointer = 11
      J& = 0
      For U& = 1 To ULTREG&("BKSTRUC")
        X$ = REGLEIDO$("BKSTRUC", U&)
        X1$ = Left$(X$, 128)
        X2$ = Mid$(X$, 129, 256)
        J& = J& + 1
        Call GRAREG("ESTRUNUE", X1$, J&)
        Call GRAREG("ESTRUDOS", X2$, J&)
      Next U&
      Call CIERRARCH("BKSTRUC")
      Screen.MousePointer = 1
    End If
    '
End Sub
'

Sub RECURSI()
    '
Call CONECRUN("#RECURSI", "Control de Recursividad de Estructuras")
Exit Sub

20  Dim A$(8)
    Dim CII%(1024), CII1%(1024), NIVE%(1024), NIVE1%(1024)
    Dim cix%(32)
    '
    Screen.MousePointer = 11
110 NM% = NUMAS%
    CONT2% = NUINV%
    '
120 NESTRU% = 0
    PADANT% = 0
    PUNTCI$ = ""
    '
    For I& = 1 To ULTREG&("ESTRUNUE")
            X$ = REGLEIDO$("ESTRUNUE", I&)
            NESTRU% = NESTRU% + 1
            PADANT% = CVI(Left$(X$, 2))
            PUNTCI$ = PUNTCI$ + MKI$(PADANT%)
    Next I&
    '
130 FIN% = NESTRU%

        For REG% = 1 To NESTRU%

        '        Call TRAZA(20, REG%, FIN%)

       '         GET #N4%,REG%
                CI% = CVI(R1$)
                If CI% > 0 Then
                    If CI% <= NM% Then
        '                GET #N1%,CI%+1
                        CODES$ = RTrim$(LTrim$(M5$)) + " - " + RTrim$(LTrim$(M3$))
                        CODESX$ = CODES$
                        While Len(CODES$) < 78: CODES$ = " " + CODES$ + " ": Wend
                       ' LOCATE 13, 2: Print Left$(CODES$, 78)
                        '
                        GoSub 16000       ' revisar items de baja
                        '
                        GoSub 17100
                        '
                    End If
                End If
        Next REG%

3900    Call CIERRARCH("*.*")
        'Call FINAL("AMASTOK")
        Close: End


16000 'GET #N4%,REG%:GET #N5%,REG%
      DPS$ = R2$ + R4$
      '
      For K% = 1 To Len(DPS$) Step 8
           DPSX$ = Mid$(DPS$, K%, 8)
           CIY% = CVI(DPSX$)
           If CIY% > 0 Then
              CODY$ = CODEXT$(CIY%)
              CIZ% = CODINT%(CODY$)
              If CIZ% < 1 Or CIZ% > NM% Then
                  TXT$ = "Codigo: " + TRIM$(M5$) + "\Incluye Items de Baja.\-\Revisar Oportunamente."
                  Call COMUNI(TXT$)
              End If
           End If
      Next K%
      '
      Return


17000 ENCU% = 0: PX1% = 1
17005 PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
      If PX% < PX1% Then GoTo 17035
      If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 17005
      L% = (PX% + 1) / 2
      ENCU% = L%
      'GET #N4%,L%
17035 Return
17040 '

17100   CONTA% = 1
        NIVE%(1) = 0
        CII%(1) = CI%
        NIVBAS% = 0
        J% = 0

17105 For II% = 1 To CONTA%
        J% = J% + 1
        NIVE1%(J%) = NIVE%(II%)
        CII1%(J%) = CII%(II%)

        If NIVE%(II%) <> NIVBAS% Then GoTo 17135

                CIBAS% = CII%(II%)
                GoSub 17000
                        If ENCU% < 1 Or ENCU% > NESTRU% Then GoTo 17135
                GoSub 17200    ' AGREGAR ITEMS
17135 Next II%

      If J% <= CONTA% Then GoTo 17180 '   TERMINAR - NO SE AGREGARON ITEMS



17145 CONTA% = J%: RETNIVBAS% = NIVBAS%

        For II% = 1 To CONTA%
                NIVE%(II%) = NIVE1%(II%)
                CII%(II%) = CII1%(II%)
        Next II%

        For RGL% = CONTA% To 1 Step -1

                GoSub 17300
                If VALIDA% < 0 Then

                  '      Print Chr$(7);
                        Call COMUNI("Recursividad en Item\" + TRIM$(CODEXT$(CIBAS%)) + "\" + TRIM$(DESCRIT$(CIBAS%)))

                        GoSub 17400
                   '     Call SETCOL(1, 2)
                        NIVBAS% = NIVE%(RGL%)
                   '     Call VENTANA(8, 7, 10 + NIVBAS%, 74, "=", "")

                                For PQ% = 0 To NIVBAS%
       '                                 GET #N1%,CIX%(PQ%)+1
                                        NIVES$ = "0" + String$(NIVBAS%, ".")
                                        If PQ% > 0 Then NIVES$ = String$(PQ%, ".") + Mid$(Str$(PQ%), 2) + String$(NIVBAS% - PQ%, ".")

                             '           LOCATE 9 + PQ%, 9: Print NIVES$ + "  " + M5$ + " " + Left$(M3$, 45 - NIVBAS%)
                                Next PQ%

                   '     Call SETCOL(1, 1)

                        XI% = COMALTER%("&17.62&Revise la Cadena Recursiva\\Continuar\Abandonar Control")
                        If XI% <> 1 Then GoTo 3900

                    '    Call VENTANA(8, 24, 9 + NIVBAS%, 60, "=", "RESTORE")

                        GoTo 17180

                End If

        Next RGL%

        J% = 0
        NIVBAS% = RETNIVBAS%
        If NIVBAS% < 9 Then NIVBAS% = NIVBAS% + 1: GoTo 17105

17180   Return



17200 'GET #N4%,ENCU%
        DPS$ = R2$

        For K% = 1 To 120 Step 8
                DPSX$ = Mid$(DPS$, K%, 8)
                If CVI(DPSX$) < 1 Or CVI(DPSX$) > NM% Then GoTo 17280

                J% = J% + 1: NIVE1%(J%) = NIVBAS% + 1
                        CII1%(J%) = CVI(DPSX$)
        Next K%

17230 'GET #N5%,ENCU%
        DPS$ = R4$

        For K% = 1 To 256 Step 8
                DPSX$ = Mid$(DPS$, K%, 8)
                If CVI(DPSX$) < 1 Or CVI(DPSX$) > NM% Then GoTo 17280

                J% = J% + 1: NIVE1%(J%) = NIVBAS% + 1
                        CII1%(J%) = CVI(DPSX$)
        Next K%

17280 Return


17300 VALIDA% = 0
        NIVBAS% = NIVE%(RGL%)
        CIBAS% = CII%(RGL%)

        For PQ% = RGL% - 1 To 1 Step -1
                If NIVE%(PQ%) >= NIVBAS% Then GoTo 17325
                If NIVE%(PQ%) < NIVBAS% Then NIVBAS% = NIVE%(PQ%)
                If CII%(PQ%) = CIBAS% Then VALIDA% = -1: GoTo 17335
17325   Next PQ%
17330 '
17335 Return
17340 '

17400   NIVBAS% = NIVE%(RGL%): LOCAD% = NIVBAS%
        For PQ% = 0 To 10: cix%(PQ%) = 0: Next PQ%
        cix%(NIVBAS%) = CII%(RGL%)

        For PQ% = RGL% - 1 To 1 Step -1
                If NIVE%(PQ%) >= NIVBAS% Then GoTo 17425
                If NIVE%(PQ%) < NIVBAS% Then
                        NIVBAS% = NIVE%(PQ%)
                        cix%(NIVBAS%) = CII%(PQ%)
                        End If
17425   Next PQ%
17430 '
17435 Return
17440 '
    
End Sub
'
