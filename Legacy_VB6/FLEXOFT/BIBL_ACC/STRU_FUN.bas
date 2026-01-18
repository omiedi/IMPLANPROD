Attribute VB_Name = "STRU_FUN"
Public CAIT%
Public CIJ%(512), USOI(512), PDI%(512), EXPLOLIN$(512), REFCOMP$(512)
Public HABIPORMI%, PORMILLAR%, NOCHECK%
'
Public CONTAMUL%
Public NIVE%(1024), CIK%(1024), CANT(1024), PAD%(1024), REFEX$(1024)
Public NIVE1%(1024), CII1%(1024), CANT1(1024), PAD1%(1024), REFEX1$(1024)
'
Public FORMUCON%, SABOKIT%
Public PUNTCI As String
'
Public DESARMA%
Public REVIFORMU
'
Sub AREXPLO(CIBAS%, CABA, OPCI%)
    '
    Call APERBASDAT("STRUC")
600 If OPCI% = 1 Or OPCI% = 3 Then
       Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas WHERE FORMULAS!CODICONJ=" & CIBAS%, dbOpenSnapshot, dbReadOnly)
615    CAIT% = 0
       On Error Resume Next
       FORMULASTEMP.MoveFirst
       If Err Then Exit Sub
       '
       Do Until FORMULASTEMP.EOF = True
         CAIT% = CAIT% + 1
         CIJ%(CAIT%) = FORMULASTEMP!CODIELEM
         USOI(CAIT%) = FORMULASTEMP!USOBRUTO
         PDI%(CAIT%) = FORMULASTEMP!PLADISPO
         '
         USOUNI = 0: CAUNIDS = 0: LARCORTE = 0
         On Error Resume Next
         CAUNIDS = FORMULASTEMP!UNIS_FORM
         LARCORTE = FORMULASTEMP!LAR_CORTE
         On Error GoTo 0
         PESOMET = PESMETRO(CIJ%(CAIT%))
         If PESOMET > 0 Then
           If LARCORTE > 0 Then
             If CAUNIDS > 0 Then
               USOUNI = PESOMET * LARCORTE / 1000
               USOI(CAIT%) = CAUNIDS * USOUNI
             End If
           End If
         End If
         '
         REFCOMP$(CAIT%) = ""
         On Error Resume Next
           REFCOMP$(CAIT%) = FORMULASTEMP!RefeComp
         On Error GoTo 0
         '
         If OPCI% = 1 Then
           CODXX$ = CODEXT$(CIJ%(CAIT%))
           DEXX$ = DESCRIT$(CIJ%(CAIT%))
           UNXX$ = UNIMED$(CIJ%(CAIT%))
           USON = CABA * USOI(CAIT%)
           If PORMILLAR% < 1 Then
             If HABIPORMI% > 0 Then
               If USON > 0 Then
                 If USON < 0.005 Then
                   PORMILLAR% = 1
                   GoTo 615
                 End If
               End If
             End If
           End If
           If PORMILLAR% = 1 Then USON = 1000 * USON
           USOX$ = CSTRING$(MKS$(USON), 4, 12, 4, 2)
           If CAUNIDS = 0 Then CAUNIDS = USON
           PDX$ = CSTRING$(MKI$(PDI%(CAIT%)), 1, 4, 0, 2)
           XXX$ = Space$(128)
           Call REPLA(XXX$, CODXX$, 1, 16)
           Call REPLA(XXX$, DEXX$, 17, 28)
           Call REPLA(XXX$, UNXX$, 47, 2)
           Call REPLA(XXX$, USOX$, 49, 12)
           Call REPLA(XXX$, PDX$, 61, 4)
           Call REPLA(XXX$, FORMATNUM(CAUNIDS, "F9.4"), 89, 9)
           Call REPLA(XXX$, FORMATNUM(PESOMET, "F7.3"), 101, 7)
           Call REPLA(XXX$, FORMATNUM(LARCORTE, "F6.1"), 111, 6)
           Call REPLA(XXX$, FORMATNUM(USOUNI, "F8.4"), 119, 8)
           XXX$ = XXX$ + REFCOMP$(CAIT%)
           EXPLOLIN(CAIT%) = XXX$
         End If
         '
         FORMULASTEMP.MoveNext
       Loop
       Exit Sub
       '
     ElseIf OPCI% = 2 Then
       '
       Call APERBASDAT("STRUC")
       YAPRIRE% = 0
       On Error Resume Next
       Formulas.MoveFirst
       If Err > 0 Then GoTo 625
       On Error GoTo 0
       '
620    If YAPRIRE% = 0 Then
           Formulas.FindFirst "CODICONJ =" & CIBAS%
           YAPRIRE% = 1
         Else
           Formulas.FindNext "CODICONJ =" & CIBAS%
       End If
       If Formulas.NoMatch = False Then
         Formulas.Edit
         Formulas!MARBORRA = 1
         Formulas.Update
         GoTo 620
       End If
       '
625    On Error GoTo 0
       For I& = 1 To CAIT%
         XXX$ = EXPLOLIN(I&)
         CODXX$ = Mid$(XXX$, 1, 16)
         USON = XVALO(Mid$(XXX$, 49, 12)) / CABA
           If PORMILLAR% = 1 Then
             USON = USON / 1000
           End If
         PDII% = XVALO(Mid$(XXX$, 61, 4))
         If PDII% > 255 Then PDII% = 0
         CIII% = Abs(CODINT%(CODXX$))
         REFCO$ = TRIM$(Mid$(XXX$, 129))
         If Abs(USON) > 0 Then
           If CIII% > 0 Then
             If CIII% <= NUMAS% Then
               Formulas.AddNew   ' YA CONVERTIDO
                       ON ERROR RESUME NEXT
               Formulas!CodiEmpr = Codiemp%
                       ON ERROR GOTO 0
               Formulas!CODICONJ = CIBAS%
               Formulas!CODXCONJ = CODEXT$(CIBAS%)
               Formulas!CODIELEM = CIII%
               Formulas!CODXELEM = CODEXT$(CIII%)
                 On Error Resume Next
               Formulas!UNIS_FORM = XVALO(Mid$(XXX$, 89, 9))
               Formulas!LAR_CORTE = XVALO(Mid$(XXX$, 111, 6))
                 On Error GoTo 0
               Formulas!USOBRUTO = USON
               Formulas!PLADISPO = PDII%
               On Error Resume Next
                 Formulas!RefeComp = REFCO$
               On Error GoTo 0
               Formulas!MARBORRA = 0
               Formulas.Update
             End If
           End If
         End If
       Next I&
       '
       YAPRIRE% = 0
630    If YAPRIRE% = 0 Then
           Formulas.FindFirst "CODICONJ =" & CIBAS% & " AND MARBORRA = 1"
           YAPRIRE% = 1
         Else
           Formulas.FindNext "CODICONJ =" & CIBAS% & " AND MARBORRA = 1"
       End If
       If Formulas.NoMatch = False Then
         'If Formulas!MARBORRA = 1 Then
           Formulas.Delete
           GoTo 630
         'End If
       End If
       '
       Formulas.Close
    End If
    '
999 End Sub
'
Sub AREXPLOant(CIBAS%, CABA, OPCI%)
    '
600 If OPCI% = 1 Or OPCI% = 3 Then
       Set FORMULASTEMP = Articulos.OpenRecordset("SELECT * FROM Formulas WHERE FORMULAS!CODICONJ=" & CIBAS%, dbOpenSnapshot, dbReadOnly)
       CAIT% = 0
       On Error Resume Next
       FORMULASTEMP.MoveLast
       If Err Then Exit Sub
       CAREC% = FORMULASTEMP.RecordCount
       If CAREC% < 1 Then Exit Sub
       '
       FORMULASTEMP.MoveFirst
       ACTIREC% = 1
610    If FORMULASTEMP.NoMatch = False Then
          CAIT% = CAIT% + 1
          CIJ%(CAIT%) = FORMULASTEMP!CODIELEM
          USOI(CAIT%) = FORMULASTEMP!USOBRUTO
          PDI%(CAIT%) = FORMULASTEMP!PLADISPO
          '
          If OPCI% = 1 Then
            CODXX$ = CODEXT$(CIJ%(CAIT%))
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
          '
          If ACTIREC% < CAREC% Then
            ACTIREC% = ACTIREC% + 1
            FORMULASTEMP.MoveNext
            GoTo 610
          End If
       End If
       '
       Exit Sub
       '
     ElseIf OPCI% = 2 Then
       '
       Call APERBASDAT("STRUC")
       Formulas.MoveFirst
       Formulas.FindFirst "CODICONJ =" & CIBAS%
       If Formulas.NoMatch = False Then
          Formulas.Edit
          Formulas!MARBORRA = 1
          Formulas.Update
       End If
       '
620    Formulas.FindNext "CODICONJ =" & CIBAS%
       If Formulas.NoMatch = False Then
          Formulas.Edit
          Formulas!MARBORRA = 1
          Formulas.Update
          GoTo 620
       End If
       '
       For I& = 1 To CAIT%
         XXX$ = EXPLOLIN(I&)
         CODXX$ = Mid$(XXX$, 1, 16)
         USON = XVALO(Mid$(XXX$, 49, 12)) / CABA
         PDII% = XVALO(Mid$(XXX$, 61, 4))
         REFCO$ = Mid$(XXX$, 129)
         If PDII% > 255 Then PDII% = 0
         CIII% = CODINT%(CODXX$)
         If Abs(USON) > 0 Then
           If CIII% > 0 Then
             If CIII% <= NUMAS% Then
               Formulas.AddNew   ' YA CONVERTIDO
                       ON ERROR RESUME NEXT
               Formulas!CodiEmpr = Codiemp%
                       ON ERROR GOTO 0
               Formulas!CODICONJ = CIBAS%
               Formulas!CODXCONJ = CODEXT$(CIBAS%)
               Formulas!CODIELEM = CIII%
               Formulas!CODXELEM = CODEXT$(CIII%)
               Formulas!USOBRUTO = USON
               Formulas!PLADISPO = PDII%
               On Error Resume Next
                 Formulas!RefeComp = REFCO$
               On Error GoTo 0
               Formulas!MARBORRA = 0
               Formulas.Update
             End If
           End If
         End If
       Next I&
       '
       ' RVFF% = HOY(HOS$)
       ' Call REPLA(DPS0$, MKI$(RVFF%), 121, 2)
       ' Call GRAREG("ESTRUNUE", MKI$(CIBAS%) + Left$(DPS0$, 126), REG&)
        '
625    Formulas.FindFirst "CODICONJ =" & CIBAS%
       If Formulas.NoMatch = False Then
         If Formulas!MARBORRA = 1 Then
           Formulas.Delete
           GoTo 625
         End If
       End If
       '
630    Formulas.FindNext "CODICONJ =" & CIBAS%
       If Formulas.NoMatch = False Then
         If Formulas!MARBORRA = 1 Then
           Formulas.Delete
           GoTo 630
         End If
       End If
       '
       Formulas.Close
    End If
    '
999 End Sub
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

Sub LEEEXPLOMUL(COD$, CABAS, modo%)     ' GENERACION AUTOMATICA ARCHIVO DE EXPLOSION LINEAL
    '
    ' MODO%=2 fuerza explosion multilineal para SABO
    '
    If FORMUCON% = 0 Then
       FORMUCON% = 1
       If modo% = 2 Or CONTROL$("", "FORMUCON") = "MULTILIN" Then
           FORMUCON% = 2
       End If
       SABOKIT% = 0: If modo% = 2 Then SABOKIT% = 1
    End If
    '
    CAIT% = 0
    CABA = CABAS: If CABA = 0 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    If CIBAS% > 0 Then
       If FORMUCON% = 2 Then
            Call EXPLOFAI(CIBAS%, CABA, SABOKIT%)
          Else
            Call AREXPLO(CIBAS%, CABA, 1)
       End If
    End If
    '
End Sub
'
Sub EXPLOMUL(CI%)
    '
    CONTAMUL% = 1
    NIVE%(CONTAMUL%) = 0
    CIK%(CONTAMUL%) = CI%
    CANT(CONTAMUL%) = 1
    PAD%(CONTAMUL%) = 0
    REFEX$(CONTAMUL%) = ""
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
      REFEX1$(J%) = REFEX$(J%)
      If NIVE%(I%) = NIVESEL% Then
        CIBAS% = CIK%(I%)
        CABA = CANT(I%)
        Call AREXPLO(CIBAS%, CABA, 3)
        If CAIT% > 0 Then
          HAYEXPLO% = 1
          For K% = 1 To CAIT%
            J% = J% + 1
            NIVE1%(J%) = NIVESEL% + 1
            CII1%(J%) = CIJ%(K%)
            CANT1(J%) = CABA * USOI(K%)
            PAD1%(J%) = PAD%(I%) + PDI%(K%)
            REFEX1$(J%) = REFCOMP$(K%)
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
      REFEX$(I%) = REFEX1$(I%)
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
Function CANIVEXPLO%(CI%)   ' DETERMINA CANTIDAD DE NIVELES DE EXPLOSIóN
                            ' SI ES CERO ES UNA MATERIA PRIMA
    CANIVI% = 0
    Call EXPLOMUL(CI%)
    For KKI% = 1 To CONTAMUL%
      If NIVE%(KKI%) > CANIVI% Then CANIVI% = NIVE%(KKI%)
    Next KKI%
    '
    CANIVEXPLO% = CANIVI%
    '
End Function
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
   Call EXPLOMUL(CIBASI%)
   FIN& = CONTAMUL%
   KKJ% = 0
   '
   For KKI% = 2 To CONTAMUL
     If KKI% >= CONTAMUL% Then GoTo 15
     If NIVE%(KKI% + 1) <= NIVE%(KKI%) Then GoTo 15
     If PROVHABI%(CIK%(KKI%)) > 0 Then GoTo 15
     GoTo 19
     '
15   KKJ% = KKJ% + 1
     CIK%(KKJ%) = CIK%(KKI%)
     CANT(KKJ%) = CANT(KKI%)
19 Next KKI%
   '
   For KKK% = 1 To KKJ%
     For KKL% = 1 To KKK% - 1
       If CIK%(KKL%) = CIK%(KKK%) Then
         CANT(KKL%) = CANT(KKL%) + CANT(KKK%)
         CANT(KKK%) = 0
         GoTo 29
       End If
     Next KKL%
29 Next KKK%
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
    Call CIERRARCH("*.*")
    Call COPYFORMU(CODORI%, CODEST%)
    '
End Sub

Sub COPYFORMU(CODORI%, CODEST%)
   '
   Screen.MousePointer = 11
   Call AREXPLO(CODORI%, 1, 1)
   Call AREXPLO(CODEST%, 1, 2)
   Screen.MousePointer = 1
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
                CI% = CVI(R1$)
                If CI% > 0 Then
                    If CI% <= NM% Then
                        CODES$ = RTrim$(LTrim$(M5$)) + " - " + RTrim$(LTrim$(M3$))
                        CODESX$ = CODES$
                        While Len(CODES$) < 78: CODES$ = " " + CODES$ + " ": Wend
                        GoSub 16000       ' revisar items de baja
                        GoSub 17100
                        '
                    End If
                End If
        Next REG%

3900    Call CIERRARCH("*.*")
        Close: End


16000 DPS$ = R2$ + R4$
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

