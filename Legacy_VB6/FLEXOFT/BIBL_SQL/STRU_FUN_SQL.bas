Attribute VB_Name = "STRU_FUN_SQL"
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
Function CANCOMPO%(CIBAS%)
    CONDI$ = "CODICONJ = " & CIBAS%
    CANCOMPO% = CantRegistros&("FORMULAS", CONDI$)
End Function

Sub AREXPLO(CIBAS%, CABA, OPCI%, Optional COMPLETO%, Optional LEADTIME%)
    ' SE AGREGA PARAMETRO COMPLETO% PARA INCLUIR (1) O EXCLUIR (0) LOS COMPONENTES DE BAJA
    '
    
    
    Dim VectArticulos As New articulos
    VectArticulos.cargarvector
    Call APERBASDAT("STRUC")
600 If OPCI% = 1 Or OPCI% = 3 Then
       '
       Dim FORMULASTEMP As ADODB.Recordset
       Set FORMULASTEMP = New ADODB.Recordset
       STRSQL = "SELECT * FROM Formulas WITH(NOLOCK) "
       STRSQL = STRSQL + " WHERE CODICONJ = " & CIBAS%
       STRSQL = STRSQL + " ORDER BY NuorIt ASC"
       If COMPLETO% = 0 Then ' OBVIA LOS COMPONENTES QUE ESTAN DADOS DE BAJA
            STRSQL = "SELECT  MASTER.STATUS, MASTER.PESMETRO, MASTER.LETICOM, FORMULAS.CODIELEM, FORMULAS.USOBRUTO, FORMULAS.PLADISPO, "
            STRSQL = STRSQL + " FORMULAS.UNIS_FORM, FORMULAS.LAR_CORTE, FORMULAS.RefeComp "
            STRSQL = STRSQL + " FROM FORMULAS WITH(NOLOCK) "
            STRSQL = STRSQL + " INNER JOIN MASTER WITH(NOLOCK) "
            STRSQL = STRSQL + " ON FORMULAS.CodIELEM = MASTER.CODINT "
            STRSQL = STRSQL + " WHERE MASTER.STATUS > 0 AND FORMULAS.CODICONJ = " & CIBAS%
            STRSQL = STRSQL + " ORDER BY FORMULAS.NuorIt ASC"
       End If
       'Set FORMULASTEMP = READSET(CStr(strsql))
       'Set FORMULASTEMP = FLEXCONN.Execute(FILTSQL$(strsql))
       FORMULASTEMP.CursorType = adOpenForwardOnly
       FORMULASTEMP.LockType = adLockReadOnly
       FORMULASTEMP.CursorLocation = adUseClient
       FORMULASTEMP.Open STRSQL, FLEXCONN, , , adCmdText
615 CAIT% = 0
       On Error Resume Next
       FORMULASTEMP.MoveFirst
       If Err Then
          Exit Sub
       End If
       On Error GoTo 0
       '
       Do Until FORMULASTEMP.EOF = True
         CAIT% = CAIT% + 1
         CIJ%(CAIT%) = XVALO(FORMULASTEMP!codielem)
         USOI(CAIT%) = XVALO(FORMULASTEMP!usobruto)
         PDI%(CAIT%) = XVALO(FORMULASTEMP!PLADISPO)
         ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
         If LEADTIME% > 0 Then
            LT% = XVALO(FORMULASTEMP!LETICOM)
            If LT% > 0 Then
               LTI_SEMANAS% = Int(LT% / 7) ' LETICOM ESTA EXPRESADO EN DIAS. AL DIVIDIRLO EN 7 LO LLEVO A SEMANAS
               If (LT% Mod 7) > 0 Then LTI_SEMANAS% = LTI_SEMANAS% + 1 ' REDONDEA AL ENTERO SUPERIOR
               PDI%(CAIT%) = LTI_SEMANAS%
            End If
         End If
         '
         USOUNI = 0: CAUNIDS = 0: LARCORTE = 0
         CAUNIDS = XVALO(FORMULASTEMP!UNIS_FORM)
         LARCORTE = XVALO(FORMULASTEMP!LAR_CORTE)
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
         REFCOMP$(CAIT%) = SAFETEXT$(FORMULASTEMP!RefeComp)
         '
         If OPCI% = 1 Then
           CODXX$ = CODEXT$(CIJ%(CAIT%))
           DEXX$ = DESCRIT$(CIJ%(CAIT%))
           CODXX$ = VectArticulos.Codigo(CIJ%(CAIT%))
           DEXX$ = VectArticulos.Descripcion(CIJ%(CAIT%))
           
            'INCLUYE LEYENDA BAJA PARA LOS COMPONENTES QUE ESTAN INACTIVOS
            STAT% = XVALO(VALOBADA("MASTER", "STATUS", "CODINT=" & CIJ%(CAIT%)))
            If STAT% < 0 And COMPLETO% = 1 Then
                'DEXX$ = "(Baja) " & DESCRIT$(CIJ%(CAIT%))
                DEXX$ = "(Baja) " & VectArticulos.Descripcion(CIJ%(CAIT%))
            End If
           'UNXX$ = Unimed$(CIJ%(CAIT%))
           UNXX$ = VectArticulos.Unimed(CIJ%(CAIT%))
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
           If PORMILLAR% = 1 Then
               USON = 1000 * USON
             Else
               ' CONVERSION DE FORMULA
               Call CONVERUNID(USON, UNXX$)
           End If
           '
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
619      FORMULASTEMP.MoveNext
       Loop
       '
       FORMULASTEMP.Close
       Set FORMULASTEMP = Nothing
       Exit Sub
       '
     ElseIf OPCI% = 2 Then
       '
       Call APERBASDAT("STRUC")
       Call CREACAMPO("", "FORMULAS", "UMED_FORM", 10, 8)
       'sql
       
       STRSQ$ = "Update Formulas set marborra=1 where CODICONJ=" & CIBAS%
       Call EJECUTACOMANDO(STRSQ$)
       '
       Dim RFORMULAS As ADODB.Recordset
       Set RFORMULAS = New ADODB.Recordset
       'RFORMULAS.CursorType = adOpenDynamic
       RFORMULAS.CursorType = adOpenKeyset
       RFORMULAS.LockType = adLockPessimistic
       STRSQL = ("SELECT * FROM Formulas WHERE CODICONJ=" & CIBAS%)
25     On Error Resume Next
       RFORMULAS.Open FILTSQL$(STRSQL), FLEXCONN, , , adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       MAXCAIT& = CAIT%
       For i& = 1 To CAIT%
         Call TRACE(20, i&, MAXCAIT&)
         XXX$ = EXPLOLIN(i&)
         CODXX$ = Mid$(XXX$, 1, 16)
         USON = XVALO(Mid$(XXX$, 49, 12)) / CABA
           If PORMILLAR% = 1 Then
               USON = USON / 1000
             Else
               CIELE% = CODINT%(CODXX$)
               UNIFOR$ = TRIM$(Mid$(XXX$, 47, 2))
               UNISTO$ = Unimed$(CIELE%)
               UNISTO$ = VectArticulos.Unimed(CIELE%)
               If UNIFOR$ <> UNISTO$ Then
                  FACO = FACONVERT(UNISTO$, UNIFOR$)
                  If FACO > 0 Then USON = USON / FACO
               End If
           End If
         PDII% = XVALO(Mid$(XXX$, 61, 4))
         If PDII% > 255 Then PDII% = 0
         CIII% = Abs(CODINT%(CODXX$))
         REFCO$ = TRIM$(Mid$(XXX$, 129))
         If Abs(USON) > 0 Then
           If CIII% > 0 Then
             If CIII% <= NUMAS% Then
               RFORMULAS.AddNew   ' YA CONVERTIDO
                        On Error Resume Next
               RFORMULAS!CodiEmpr = CodiEmp%
                        On Error GoTo 0
               RFORMULAS!CODICONJ = CIBAS%
               'RFORMULAS!CODXCONJ = CODEXT$(CIBAS%)
               RFORMULAS!CODXCONJ = VectArticulos.Codigo(CIBAS%)
               RFORMULAS!codielem = CIII%
               'RFORMULAS!codxelem = CODEXT$(CIII%)
               RFORMULAS!CodXElem = VectArticulos.Codigo(CIII%)
               RFORMULAS!NUORIT = i&
                 On Error Resume Next
               RFORMULAS!UNIS_FORM = XVALO(Mid$(XXX$, 89, 9))
               RFORMULAS!UMED_FORM = TRIM$(Mid$(XXX$, 47, 4))
               RFORMULAS!LAR_CORTE = XVALO(Mid$(XXX$, 111, 6))
                 On Error GoTo 0
               RFORMULAS!usobruto = USON
               RFORMULAS!PLADISPO = PDII%
               On Error Resume Next
                 RFORMULAS!RefeComp = REFCO$
               On Error GoTo 0
               RFORMULAS!MARBORRA = 0
               RFORMULAS.Update
             End If
           End If
         End If
       Next i&
       '
       RFORMULAS.Close
       Set RFORMULAS = Nothing
       
       STRSQL = "Delete Formulas where marborra=1 and CODICONJ=" & CIBAS%
       FLEXCONN.Execute STRSQL
       On Error Resume Next
       '
    End If
    '
999 End Sub
'
Sub LEEEXPLO(COD$, CABAS, Optional COMPLETO%)      ' GENERACION AUTOMATICA ARCHIVO DE EXPLOSION LINEAL
    'SE AGREGA PARAMETRO COMPLETO% PARA INCLUIR (1) O EXCLUIR (0) LOS COMPONENTES DE BAJA
    '
    CAIT% = 0
    CABA = CABAS: If CABA = 0 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    If CIBAS% > 0 Then
        Call AREXPLO(CIBAS%, CABA, 1, COMPLETO%)
    End If
    '
End Sub
'

Sub LEEEXPLOMUL(COD$, CABAS, MODO%, Optional COMPLETO%, Optional LEADTIME%)     ' GENERACION AUTOMATICA ARCHIVO DE EXPLOSION LINEAL
    'SE AGREGA PARAMETRO COMPLETO% PARA INCLUIR (1) O EXCLUIR (0) LOS COMPONENTES DE BAJA
    '
    ' MODO%=2 fuerza explosion multilineal para SABO
    ' MODO%=3 fuerza explosion multinivel pero sin la característica de SABOKIT% = 1
    
    If FORMUCON% = 0 Then
       FORMUCON% = 1
       If MODO% = 2 Or Control$("", "FORMUCON") = "MULTILIN" Then
           FORMUCON% = 2
       End If
       If MODO% = 3 Then FORMUCON% = 2 '
       SABOKIT% = 0: If MODO% = 2 Then SABOKIT% = 1
    End If
    '
    CAIT% = 0
    CABA = CABAS: If CABA = 0 Then CABA = 1
    CIBAS% = CODINT%(COD$)
    If CIBAS% > 0 Then
       If FORMUCON% = 2 Then
            ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
            Call EXPLOFAI(CIBAS%, CABA, SABOKIT%, LEADTIME%)
          Else
            ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
            Call AREXPLO(CIBAS%, CABA, 1, COMPLETO%, LEADTIME%)
       End If
    End If
    '
End Sub
'
Sub EXPLOMUL(Ci%, Optional LEADTIME%)
    '
    CONTAMUL% = 1
    NIVE%(CONTAMUL%) = 0
    CIK%(CONTAMUL%) = Ci%
    CANT(CONTAMUL%) = 1
    PAD%(CONTAMUL%) = 0
    REFEX$(CONTAMUL%) = ""
    '
    NIVESEL% = 0
    '
10  J% = 0
    HAYEXPLO% = 0
    For i% = 1 To CONTAMUL%
      J% = J% + 1
      NIVE1%(J%) = NIVE%(i%)
      CII1%(J%) = CIK%(i%)
      CANT1(J%) = CANT(i%)
      PAD1%(J%) = PAD%(i%)
      REFEX1$(J%) = REFEX$(J%)
      If NIVE%(i%) = NIVESEL% Then
        CIBAS% = CIK%(i%)
        CABA = CANT(i%)
        Call AREXPLO(CIBAS%, CABA, 3, , LEADTIME%)
        If CAIT% > 0 Then
          HAYEXPLO% = 1
          For k% = 1 To CAIT%
            J% = J% + 1
            NIVE1%(J%) = NIVESEL% + 1
            CII1%(J%) = CIJ%(k%)
            CANT1(J%) = CABA * USOI(k%)
            PAD1%(J%) = PAD%(i%) + PDI%(k%)
            REFEX1$(J%) = REFCOMP$(k%)
          Next k%
        End If
      End If
    Next i%
    '
    CONTAMUL% = J%
    For i% = 1 To CONTAMUL%
      NIVE%(i%) = NIVE1%(i%)
      CIK%(i%) = CII1%(i%)
      CANT(i%) = CANT1(i%)
      PAD%(i%) = PAD1%(i%)
      REFEX$(i%) = REFEX1$(i%)
    Next i%
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
Function CANIVEXPLO%(Ci%)   ' DETERMINA CANTIDAD DE NIVELES DE EXPLOSIóN
                            ' SI ES CERO ES UNA MATERIA PRIMA
    CANIVI% = 0
    Call EXPLOMUL(Ci%)
    For KKI% = 1 To CONTAMUL%
      If NIVE%(KKI%) > CANIVI% Then CANIVI% = NIVE%(KKI%)
    Next KKI%
    '
    CANIVEXPLO% = CANIVI%
    '
End Function
'
Sub EXPLOFAI(CIBASI%, CABASI, SABOKIT%, Optional LEADTIME%)
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
    For i% = 1 To CONTAMUL%
        J% = J% + 1
        NIVE1%(J%) = NIVE%(i%)
        CII1%(J%) = CIK%(i%)
        CANT1(J%) = CANT(i%)
        PAD1%(J%) = PAD%(i%)
        If NIVE%(i%) = NIVESEL% Then
          If ARTIPRO%(CIK%(i%)) < 1 Or i% = 1 Then
            CIBAS% = CIK%(i%)
            '
            If SABOKIT% = 1 Then
               CII7% = CIBAS%
               If InStr(UCase$(DESCRIT$(CII7%)), "KIT") <> 1 Then GoTo 19
            End If
            '
            If i% > 1 Then
              If PROVHABI%(CIK%(i%)) > 0 Then GoTo 19
            End If
            '
            CABA = CANT(i%)
            ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
            Call AREXPLO(CIBAS%, CABA, 1, , LEADTIME%)
            If CAIT% > 0 Then
                HAYEXPLO% = 1
                For k% = 1 To CAIT%
                    J% = J% + 1
                    NIVE1%(J%) = NIVESEL% + 1
                    CII1%(J%) = CIJ%(k%)
                    CANT1(J%) = CABA * USOI(k%)
                    PAD1%(J%) = PAD%(i%) + PDI%(k%)
                Next k%
            End If
          End If
        End If
19  Next i%
    '
    CONTAMUL% = J%
    For i% = 1 To CONTAMUL%
        NIVE%(i%) = NIVE1%(i%)
        CIK%(i%) = CII1%(i%)
        CANT(i%) = CANT1(i%)
        PAD%(i%) = PAD1%(i%)
    Next i%
    '
    If HAYEXPLO% = 1 Then
        NIVESEL% = NIVESEL% + 1
        If NIVESEL% <= 9 Then
            GoTo 10
        End If
    End If
    '
    J% = 0
    For i% = 2 To CONTAMUL%
       If i% = CONTAMUL% Or NIVE%(i% + 1) <= NIVE%(i%) Then
          J% = J% + 1
          NIVE%(J%) = NIVE%(i%)
          If NIVE%(J%) > 1 Then NIVE%(J%) = 1
          CIK%(J%) = CIK%(i%)
          CANT(J%) = CANT(i%) * CABASI
          PAD%(J%) = PAD%(i%)
          '
          CIJ%(J%) = CIK%(J%)
          USOI(J%) = CANT(J%) / CABASI
          PDI%(J%) = PAD%(J%)
          '
       End If
    Next i%
    '
    CAIT% = J%
    CONTAMUL% = J%
    '
End Sub
'

Sub EXPLOCOM(CIBASI%, CABASI)
   '
   CI1% = CIBASI%
   Call EXPLOMUL(CI1%)
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
     CANTU = CANT(KKK%)
     UME$ = Unimed$(CIK%(KKK%))
     Call CONVERUNID(CANTU, UME$)
     If Abs(CANTU) >= 0.000005 Then
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
    RFORMX = FETEXCOR1$(Int(REVIFORMU))
    STRSQL = ("Update Master set Revi_Formu = '" & RFORMX & "' where codint = " & CIBAS%)
    On Error Resume Next
     FLEXCONN.Execute (FILTSQL$(STRSQL))
    On Error GoTo 0
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
   If Left$(UCase$(Unimed$(CIXI%)), 1) = "U" Then
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
   '
   If Control("MASTER", "FORPATRO") = "SI" Then
      Call COPYFORPAT(CODORI%, CODEST%)
   End If
   '
   Screen.MousePointer = 1
   '
End Sub
Sub COPYFORPAT(CODORI%, CODEST%)
    'Se selecciona la formula patron del codigo origen
    Dim PATORG As ADODB.Recordset
    Set PATORG = New ADODB.Recordset
    SQLX$ = "SELECT * FROM FORMUPAT WHERE CODICONJ =" & CODORI%
    Set PATORG = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    'Se marca para borrar los items de la formula destino
    SQLX1$ = "CODICONJ =" & CODEST%
    Call ACTUREGISTRO("FORMUPAT", "MARBORRA", SQLX1$, 1, REG&)
    
    If PATORG.BOF And PATORG.EOF Then GoTo 99
    
    Dim PATDEST As ADODB.Recordset
    Set PATDEST = New ADODB.Recordset
    SQLX1$ = "SELECT * FROM FORMUPAT WHERE CODICONJ =" & CODEST%
25  On Error Resume Next
    PATDEST.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenKeyset, adLockPessimistic
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    
    
    
    'Se copian los items de la formula patron ORIGEN A LA DESTINO
    PATORG.MoveFirst
    Do While Not (PATORG.EOF)
       PATDEST.AddNew   ' YA CONVERTIDO
              On Error Resume Next
       PATDEST!CodiEmpr = CodiEmp%
              On Error GoTo 0
       PATDEST!CODICONJ = CODEST%
       PATDEST!CODXCONJ = CODEXT$(CODEST%)
       PATDEST!codielem = PATORG!codielem
       PATDEST!CodXElem = PATORG!CodXElem
       PATDEST!usobruto = PATORG!usobruto
       PATDEST!Operacion = PATORG!Operacion
       PATDEST!NUORIT = PATORG!NUORIT
       PATDEST!MARBORRA = 0
       PATDEST.Update
       PATORG.MoveNext
    Loop
    PATDEST.Close
    Set PATDEST = Nothing
    PATORG.Close
    Set PATORG = Nothing
    
    'SE BORRAN LOS ITEMS DE CODIGO DESTINO Q TIENEN MARBORRA = 1
    SQLX$ = "CODICONJ = " & CODEST% & " AND MARBORRA = 1"
    Call BORRAREGISTROS("FORMUPAT", SQLX$, REG&)
    LOREPOX# = VALOBADA("Master", "lorepos", "Codint=" & CODORI%)
    Call ACTUREGISTRO("MASTER", "LOREPOS", "CODINT = " & CODEST%, LOREPOX#, REG&)
    
99
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
    Dim CIX%(32)
    '
    Screen.MousePointer = 11
110 NM% = NUMAS%
    CONT2% = NUINV%
    '
120 NESTRU% = 0
    PADANT% = 0
    PUNTCI$ = ""
    '
    For i& = 1 To ULTREG&("ESTRUNUE")
            X$ = REGLEIDO$("ESTRUNUE", i&)
            NESTRU% = NESTRU% + 1
            PADANT% = CVI(Left$(X$, 2))
            PUNTCI$ = PUNTCI$ + MKI$(PADANT%)
    Next i&
    '
130 FIN% = NESTRU%

        For REG% = 1 To NESTRU%

        '        Call TRAZA(20, REG%, FIN%)

       '         GET #N4%,REG%
                Ci% = CVI(R1$)
                If Ci% > 0 Then
                    If Ci% <= NM% Then
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
      For k% = 1 To Len(DPS$) Step 8
           DPSX$ = Mid$(DPS$, k%, 8)
           CIY% = CVI(DPSX$)
           If CIY% > 0 Then
              CODY$ = CODEXT$(CIY%)
              CIZ% = CODINT%(CODY$)
              If CIZ% < 1 Or CIZ% > NM% Then
                  TXT$ = "Codigo: " + TRIM$(M5$) + "\Incluye Items de Baja.\-\Revisar Oportunamente."
                  Call COMUNI(TXT$)
              End If
           End If
      Next k%
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
        CII%(1) = Ci%
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

        For k% = 1 To 120 Step 8
                DPSX$ = Mid$(DPS$, k%, 8)
                If CVI(DPSX$) < 1 Or CVI(DPSX$) > NM% Then GoTo 17280

                J% = J% + 1: NIVE1%(J%) = NIVBAS% + 1
                        CII1%(J%) = CVI(DPSX$)
        Next k%

17230 'GET #N5%,ENCU%
        DPS$ = R4$

        For k% = 1 To 256 Step 8
                DPSX$ = Mid$(DPS$, k%, 8)
                If CVI(DPSX$) < 1 Or CVI(DPSX$) > NM% Then GoTo 17280

                J% = J% + 1: NIVE1%(J%) = NIVBAS% + 1
                        CII1%(J%) = CVI(DPSX$)
        Next k%

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
        For PQ% = 0 To 10: CIX%(PQ%) = 0: Next PQ%
        CIX%(NIVBAS%) = CII%(RGL%)

        For PQ% = RGL% - 1 To 1 Step -1
                If NIVE%(PQ%) >= NIVBAS% Then GoTo 17425
                If NIVE%(PQ%) < NIVBAS% Then
                        NIVBAS% = NIVE%(PQ%)
                        CIX%(NIVBAS%) = CII%(PQ%)
                        End If
17425   Next PQ%
17430 '
17435 Return
17440 '
    
End Sub
'
Function USOUNITARIO#(codiconjunto, codielemento)
    USOUNITARIO# = XVALO(VALOBADA("FORMULAS", "USOBRUTO", "CODICONJ = " & codiconjunto & " and CODIELEM = " & codielemento))
End Function
