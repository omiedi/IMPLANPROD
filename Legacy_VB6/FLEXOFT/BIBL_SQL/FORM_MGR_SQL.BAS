Attribute VB_Name = "FORM_MGR_SQL"
Public ARFOR$
Public CARGAIMAGENCOMOLINK%
'
Public FOR_DESC$, FOR_ORIE$, FOR_MGIZ, FOR_MGSU, AUTO_PRN$, PAGI_ANT$, PAGI_SIG$, LOCK_DSG$, FORM_OPX$, PRINTGRI$
Public RTEXT1, RCOMBO2, RCOMBO3, RTEXT7
'
Public IOBJ%, CANOBJ%, TIOBJ$(512), ONAME$(512), MULIN%(512), ISEL%(512)
Public H1(512), H2(512), V1(512), V2(512), CONMARCA%(512)
Public Texto(512), FUENTE$(512), TAFUEN$(512), ORITEXT99%(512), ORITEXT99M%(512), ORILAB99%(512)
Public CARMAN%(512), ARDOCLI$(512)
Public IOBJELB%(512), IOBJETX%(512), IOBJETXM%(512)
Public NEGRI(512), ITALI(512), SUBRA(512), INVIS(512), ENCAS(512)
Public FORPA$(512), FOREPLA$(512), TASEL$(512), CASEL$(512), CONDSEL$(512), LOCKFI$(512), OBLIGA%(512), VAMINI$(512), VAMAXI$(512), TABVALI$(512)
'
Public ORDHOJA&, TOTHOJA&, AUTOPAGE%  ' NUMERO DE HOJA ACTIVA Y TOTAL DE HOJAS DEL BATCH RECORD
'
Public CAPARINV%, COMPAR$(128)
Public CANREGRID%(128)
'
Public POVERTI
'
Sub CARGADATFOR(CODFOR$, Optional CANOBJANT%)
    '
    IOBJ% = XVALO(CANOBJANT%)
    HAYGRA% = 0
    For KKI% = CANOBJANT% + 1 To 512
       TIOBJ$(KKI) = ""
       ONAME$(KKI%) = ""
       H1(KKI%) = 0
       H2(KKI%) = 0
       V1(KKI%) = 0
       V2(KKI%) = 0
       CONMARCA%(KKI%) = 0
       Texto(KKI%) = ""
       ARDOCLI$(KKI%) = ""
       FUENTE$(KKI%) = ""
       TAFUEN$(KKI%) = ""
       NEGRI(KKI%) = 0
       ITALI(KKI%) = 0
       SUBRA(KKI%) = 0
       INVIS(KKI%) = 0
       ENCAS(KKI%) = 0
       MULIN(KKI%) = 0
       FORPA$(KKI%) = ""
       FOREPLA$(KKI%) = ""
       TASEL$(KKI%) = ""
       CASEL$(KKI%) = ""
       CONDSEL$(KKI%) = ""
       OBLIGA%(KKI%) = 1
       LOCKFI$(KKI%) = ""
       VAMINI$(KKI%) = ""
       VAMAXI$(KKI%) = ""
       TABVALI$(KKI%) = ""
       IOBJETX%(KKI%) = 0
    Next KKI%
    CANOBJ% = CANOBJANT%
    '
    ARFOR$ = LUDAT$ + CODFOR$ + ".SQR"
    If Left$(CODFOR$, 1) = "!" Then
       ARFOR$ = LUCOM$ + Mid$(CODFOR$, 2) + ".SQX"
    End If
    '
    If EXISTE%(ARFOR$) > 0 Then
       n1% = FILEOPEN%(ARFOR$, 1, 128)
       Do While Not EOF(n1%)
          Line Input #n1%, TTXX$
          If Left$(TTXX$, 8) = ">>>>>>>>" Then GoTo 25         'EMPIEZA GRAFICOS
          '
          If InStr(ARFOR$, ".SQX") > 0 Then
17           If Right$(TTXX$, 1) <> Chr$(124) Then
                On Error Resume Next
                Line Input #n1%, TTXX1$
                If Err Then GoTo 18
                On Error GoTo 0
                TTXX$ = TTXX$ + vbCrLf + TTXX1$
                GoTo 17                ' ESTO ES PARA CARGAR TEXTOS MULTILINEA
             End If
          End If
          '
18        On Error GoTo 0
          If UCase$(Left$(TTXX$, 6)) = "CONFIG" Then
             FOR_DESC$ = BUSCAVALOR(TTXX$, "FOR_DESC =")
             FOR_ORIE$ = UCase$(BUSCAVALOR(TTXX$, "FOR_ORIE ="))
             FOR_MGIZ = XVALO(BUSCAVALOR(TTXX$, "FOR_MGIZ ="))
             FOR_MGSU = XVALO(BUSCAVALOR(TTXX$, "FOR_MGSU ="))
             AUTO_PRN$ = BUSCAVALOR(TTXX$, "AUTO_PRN =")
             PAGI_ANT$ = BUSCAVALOR(TTXX$, "PAGI_ANT =")
             PAGI_SIG$ = BUSCAVALOR(TTXX$, "PAGI_SIG =")
             LOCK_DSG$ = BUSCAVALOR(TTXX$, "LOCK_DSG =")
             FORM_OPX$ = BUSCAVALOR(TTXX$, "FORM_OPX =")
             PRINTGRI$ = BUSCAVALOR(TTXX$, "PRINTGRI =")
             GoTo 19
          End If
          '
          TIOBJ$(CANOBJ% + 1) = BUSCAVALOR$(TTXX$, "TIPO =")
          ONAME$(CANOBJ% + 1) = BUSCAVALOR$(TTXX$, "NAME =")
          H1(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "H1 ="))
          V1(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "V1 ="))
          H2(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "H2 ="))
          V2(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "V2 ="))
          '
          CONMARCA%(CANOBJ% + 1) = 0
          '
          Texto(CANOBJ% + 1) = LTrim(RTrim(BUSCAVALOR$(TTXX$, "TEXT =")))
          FUENTE$(CANOBJ% + 1) = TRIM$(BUSCAVALOR$(TTXX$, "FONT ="))
          TAFUEN$(CANOBJ% + 1) = TRIM$(BUSCAVALOR$(TTXX$, "SIZE ="))
          FORPA$(CANOBJ% + 1) = TRIM$(BUSCAVALOR$(TTXX$, "FORM ="))
          ARDOCLI$(CANOBJ% + 1) = TRIM$(BUSCAVALOR$(TTXX$, "ADOC ="))
          '
          NEGRI(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "NEGRI ="))
          ITALI(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "ITALI ="))
          SUBRA(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "SUBRA ="))
          INVIS(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "INVIS ="))
          ENCAS(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "ENCAS ="))
          MULIN(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "MULIN ="))
          '
          FOREPLA$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR$(TTXX$, "FOREPLA =")))
          TASEL$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR$(TTXX$, "TASEL =")))
          CASEL$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR$(TTXX$, "CASEL =")))
          CONDSEL$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR$(TTXX$, "CONDSEL =")))
          OBLIGA%(CANOBJ% + 1) = XVALO(BUSCAVALOR$(TTXX$, "OBLIGA ="))
          LOCKFI$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR(TTXX$, "LOCKFI =")))
          VAMINI$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR(TTXX$, "VAMINI =")))
          VAMAXI$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR(TTXX$, "VAMAXI =")))
          TABVALI$(CANOBJ% + 1) = LTrim$(RTrim$(BUSCAVALOR(TTXX$, "TABVALI =")))
          '
          If VAMINI$(CANOBJ% + 1) = VAMAXI$(CANOBJ% + 1) Then
            If Left$(UCase$(FORPA$(CANOBJ% + 1)), 1) = "A" Then
              If TABVALI$(CANOBJ% + 1) = "" Then
                TABVALI$(CANOBJ% + 1) = VAMINI$(CANOBJ% + 1)
                VAMINI$(CANOBJ% + 1) = ""
                VAMAXI$(CANOBJ% + 1) = ""
              End If
            End If
          End If
          '
          Select Case TIOBJ$(CANOBJ% + 1)
          Case "LABEL"
             FOREPLA$(CANOBJ% + 1) = ""
             If V2(CANOBJ% + 1) < V1(CANOBJ% + 1) + 2 Then V2(CANOBJ% + 1) = V1(CANOBJ% + 1) + 2
             If H2(CANOBJ% + 1) < H1(CANOBJ% + 1) + 4 Then H2(CANOBJ% + 1) = H1(CANOBJ% + 1) + 4
             If Texto(CANOBJ% + 1) <> "" Then
               If XVALO(TAFUEN$(CANOBJ% + 1)) > 0 Then
                 CANOBJ% = CANOBJ% + 1
               End If
             End If
          Case "PARAM"
             If V2(CANOBJ% + 1) < V1(CANOBJ% + 1) + 2 Then V2(CANOBJ% + 1) = V1(CANOBJ% + 1) + 2
             If H2(CANOBJ% + 1) < H1(CANOBJ% + 1) + 4 Then H2(CANOBJ% + 1) = H1(CANOBJ% + 1) + 4
             If FORPA$(CANOBJ% + 1) <> "" Then
               If XVALO(TAFUEN$(CANOBJ% + 1)) > 0 Then
                 CANOBJ% = CANOBJ% + 1
                 If InStr(ARFOR$, ".SQR") > 0 Then Texto(CANOBJ%) = ""
                 If UCase$(Left$(App.EXEName, 6)) = "DSGFRM" Then
                    Texto(CANOBJ%) = TMASCARA$(FORPA$(CANOBJ%))
                    If H2(CANOBJ%) < H1(CANOBJ%) Then H2(CANOBJ%) = H1(CANOBJ%) + 10
                    If V2(CANOBJ%) < V1(CANOBJ%) Then V2(CANOBJ%) = V1(CANOBJ%) + 3
                 End If
               End If
             End If
          Case "TEXT"
             If V2(CANOBJ% + 1) < V1(CANOBJ% + 1) + 2 Then V2(CANOBJ% + 1) = V1(CANOBJ% + 1) + 2
             If H2(CANOBJ% + 1) < H1(CANOBJ% + 1) + 4 Then H2(CANOBJ% + 1) = H1(CANOBJ% + 1) + 4
             If FORPA$(CANOBJ% + 1) <> "" Then
               If XVALO(TAFUEN$(CANOBJ% + 1)) > 0 Then
                 CANOBJ% = CANOBJ% + 1
                 If InStr(ARFOR$, ".SQR") > 0 Then Texto(CANOBJ%) = ""
                 If UCase$(Left$(App.EXEName, 6)) = "DSGFRM" Then
                    Texto(CANOBJ%) = TMASCARA$(FORPA$(CANOBJ%))
                    If H2(CANOBJ%) < H1(CANOBJ%) Then H2(CANOBJ%) = H1(CANOBJ%) + 10
                    If V2(CANOBJ%) < V1(CANOBJ%) Then V2(CANOBJ%) = V1(CANOBJ%) + 3
                 End If
               End If
             End If
          Case "LINE"
             FOREPLA$(CANOBJ% + 1) = ""
             CANOBJ% = CANOBJ% + 1
          Case "BOX"
             FOREPLA$(CANOBJ% + 1) = ""
             CANOBJ% = CANOBJ% + 1
          Case "GRAPH"
             'ESTO PARA CODIGO QR
'             PARAXM$ = FOREPLA$(CANOBJ% + 1)
'             IDENT_IMAG$ = TRIM$(UCase$(Left(TEXREPLA$(PARAXM$), 10)))
'             If IDENT_IMAG$ = "@IMAGENQR" Then
'             'FIN CODIGO QR NO HACE NADA
'             Else 'LIMPIA
'                FOREPLA$(CANOBJ% + 1) = ""
'             End If
             HAYGRA% = 1
             FORPA$(CANOBJ% + 1) = "A512"
             CANOBJ% = CANOBJ% + 1
          Case "GRID"
             TEGRI$ = BUSCAVALOR(TTXX$, "TEGRID =")
             If TEGRI$ <> "" Then Texto(CANOBJ% + 1) = Texto(CANOBJ% + 1) + "///TEGRID = " + TEGRI$
             CANOBJ% = CANOBJ% + 1
          End Select
19     Loop
       '
25     Close #n1%
       '
       If HAYGRA% > 0 Then
         ' COMIENZA CARGA DE GRAFICOS
         TARCH$ = LOADFILE$(ARFOR$)
         For II% = 1 To CANOBJ%
           If TIOBJ$(II%) = "GRAPH" Then
             '****AGREGAR ESTO****
             AA$ = FOREPLA$(II%)
             IDENT_IMAG$ = UCase$(Left(TEXREPLA$(AA$), 10))
             If TRIM$(IDENT_IMAG$) = "@IMAGEN_QR" Then
               RUTAARCH$ = Mid$(TEXREPLA$(AA$), 11)
               ARX$ = LOADFILE$(RUTAARCH$)
               SGRA$ = ARX$
               GoTo 40
             End If
             '*******HASTA ACA ***** Y TAMBIEN EL 40 DE MAS ABAJO
             NGRA$ = Texto(II%)
             PPXX% = InStr(UCase$(NGRA$), "/LOCKED")
             If PPXX% > 1 Then NGRA$ = Left$(NGRA$, PPXX% - 1)
             PINI& = InStr(TARCH$, ">>>>>>>>" + NGRA$)
             If PINI& > 0 Then
               PFIN& = InStr(TARCH$, NGRA$ + "<<<<<<<<")
               If PFIN& > PINI& + 16 Then
                 SGRA$ = Mid$(TARCH$, PINI& + 16, PFIN - PINI& - 16)
40               AGRA$ = RANDSTRING$(-8)
                 Call SAVEFILE(LUCOM$ + AGRA$ + ".TMP", SGRA$)
                 Texto(II%) = LUCOM$ + AGRA$ + ".TMP"
               End If
             End If
           End If
         Next II%
       End If
    End If
    '
    ' ORDENAMIENTO
'    For II% = 1 To CANOBJ%
'       For JJ% = II% + 1 To CANOBJ%
'          If H1(JJ%) >= H2(II%) And V1(JJ%) > V1(II%) - 3 And V1(JJ%) < V1(II%) + 3 Then GoTo 89 ' ESTAN EN ORDEN CORRECTO
'          HII = V2(II%) - V1(II%): If HII > 2 Then HII = 2    ' PARA NO CONFUNDIR CON OBJETOS ALTOS
'          If V1(JJ%) < V1(II%) - HII / 2 Then
'                GoSub INTERCAM
'                GoTo 89
'             ElseIf V1(JJ%) > V1(II%) + HII / 2 Then
'                GoTo 89
'             Else
'                H1II = Int(H1(II%))
'                H1JJ% = Int(H1(JJ%))
'                If H1II > H1JJ Then
'                   GoSub INTERCAM
'                End If
'          End If
'89     Next JJ%
'    Next II%
    GoTo 99
    '
    '
INTERCAM:
    Call SWAP(TIOBJ$(II%), TIOBJ$(jj%))
    Call SWAP(ONAME$(II%), ONAME$(jj%))
    Call SWAP(H1(II%), H1(jj%))
    Call SWAP(H2(II%), H2(jj%))
    Call SWAP(V1(II%), V1(jj%))
    Call SWAP(V2(II%), V2(jj%))
    Call SWAP(CONMARCA%(II%), CONMARCA%(jj%))
    Call SWAP(Texto(II%), Texto(jj%))
    Call SWAP(ARDOCLI$(II%), ARDOCLI$(jj%))
    Call SWAP(FUENTE$(II%), FUENTE$(jj%))
    Call SWAP(TAFUEN$(II%), TAFUEN$(jj%))
    Call SWAP(NEGRI(II%), NEGRI(jj%))
    Call SWAP(ITALI(II%), ITALI(jj%))
    Call SWAP(SUBRA(II%), SUBRA(jj%))
    Call SWAP(INVIS(II%), INVIS(jj%))
    Call SWAP(ENCAS(II%), ENCAS(jj%))
    Call SWAP(MULIN(II%), MULIN(jj%))
    Call SWAP(FORPA$(II%), FORPA$(jj%))
    Call SWAP(FOREPLA$(II%), FOREPLA$(jj%))
    Call SWAP(TASEL$(II%), TASEL$(jj%))
    Call SWAP(CASEL$(II%), CASEL$(jj%))
    Call SWAP(CONDSEL$(II%), CONDSEL$(jj%))
    Call SWAP(OBLIGA%(II%), OBLIGA%(jj%))
    Call SWAP(LOCKFI$(II%), LOCKFI$(jj%))
    Call SWAP(VAMINI$(II%), VAMINI$(jj%))
    Call SWAP(VAMAXI$(II%), VAMAXI$(jj%))
    Call SWAP(TABVALI$(II%), TABVALI$(jj%))
    Call SWAP(IOBJETX%(II%), IOBJETX%(jj%))
Return
    '
     
99 End Sub

Sub STDFORM(FORNA$, FILTX$, Optional MODO$, Optional TITUFORM$, Optional ENVIAMAILFACTURA%, Optional PARAM_CANCOPIAS&)
    '
    On Error Resume Next
    Unload STDSCREEN
    On Error GoTo 0
    '
    
    If PARAM_CANCOPIAS& > 0 Then CANTICOPIAS& = PARAM_CANCOPIAS& - 1 'POR QUE SI ES UNA SE IMPRIME POR DEFECTO
    
    ARFOR$ = LUDAT$ + TRIM$(FORNA$) + ".SQR"
    FORNB$ = FORNA$
    If Left$(TRIM$(FORNA$), 1) = "!" Then
       FORNB$ = Mid$(TRIM$(FORNA$), 2)
       ARFOR$ = LUCOM$ + Mid$(TRIM$(FORNA$), 2) + ".SQX"
    End If
    '
    If EXISTE%(ARFOR$) < 1 Then
       RUTAX$ = REPLACAR$(ARFOR$, "\", "/")
       Call COMUNI("Falta Archivo '" & RUTAX$ & "'\Correspondiente al Formulario de Pantalla.")
       Exit Sub
    End If
    '
    DESCRI$ = TRIM$(ECOARCH$("INDISQR", FORNB$))
    If TITUFORM$ <> "" Then DESCRI$ = TITUFORM$
    STDSCREEN.Caption = DESCRI$
    STDSCREEN.Label1 = FORNA$ + "(" + FILTX$ + ")"
    If UCase$(TRIM$(MODO$)) = "PRINT" Then
        Call STDSCREEN.Command4_Click
        ' ENVIO POR MAIL DE FACTURA, NC Y ND
        If ENVIAMAILFACTURA% > 0 Then
            POS1% = InStr(TESTCAE09.FILTX, ";")
            NumeroComprobante = TRIM$(Mid$(TESTCAE09.FILTX, 1, POS1% - 1))
            POS1% = 1 + InStr(TESTCAE09.FILTX, ";")
            Cliente& = XVALO(Mid$(TESTCAE09.FILTX, POS1%))
            If Cliente& > 0 Then
              CONDI$ = "Nume_Cli =" & CStr(Cliente&)
              DireccionCorreo$ = TRIM$(FINFACTUR07.Label1(14))
              '
              DESTINATARIO$ = VALOBADA("DEUDOR12", "Ctact_Cob", CONDI$, "NOMESS")
              STDMAIL.Text3 = CStr(Cliente&)
              STDMAIL.Text1 = DESTINATARIO$
              STDMAIL.Text6 = "Documento Electrónico: " & NumeroComprobante
              STDMAIL.Text2 = DireccionCorreo$
              TXTX$ = "": If TRIM$(DESTINATARIO$) <> "" Then TXTX$ = "Estimado/a: " & DESTINATARIO$ & vbCrLf & vbCrLf
              TXTX$ = TXTX$ & "Adjunto Estamos Enviando Comprobante Electrónico: " & NumeroComprobante & vbCrLf & vbCrLf
              TXTX$ = TXTX$ & TEXTO_MAIL_FACTURA$ & vbCrLf 'FUNCION QUE DEVUELVE EL TEXTO CONFIGURADO PARA LA FACTURA POR MAIL
              STDMAIL.Text5 = TXTX$
              '
              STDMAIL.NombreDeArchivo = NumeroComprobante
              STDMAIL.Command14_Click
            End If
        End If
50        Unload STDSCREEN
      ElseIf UCase$(TRIM$(MODO$)) = "READONLY" Then
        STDSCREEN.Command5.Enabled = False
        STDSCREEN.Show 1
      Else
        STDSCREEN.Show 1
    End If
End Sub

Function TEXTO_MAIL_FACTURA$()
      '
      'DEVUELVE EL TEXTO CARGADO EN SETUP DE VENTAS PARA EL MAIL DE FACTURAS QUE SE
      'ENVIA DE MANERA AUTOMATICA A LOS CLIENTES QUE POSEEN UN MAIL DE CONTACTO DE COBRANZA
      'QUE SE GUARDA EN ARCHIVO DE CLIENTES NUEVA SOLAPA
      '
      TEXTO_MAIL_FACTURA$ = ""
      '
      Call COPYSTRU("OBSERVOF", "TXMAIFAC")
      For U& = 1 To ULTREG&("!TXMAIFAC")
           TX$ = REGLEIDO$("!TXMAIFAC", U&)
           TEXTOBS$ = TEXTOBS$ & TRIM$(TX$) & vbCrLf
      Next U&
      '
      If TRIM$(TEXTOBS$) <> "" Then
         TEXTO_MAIL_FACTURA$ = TEXTOBS$
      End If
      '
      Call CIERRARCH("!TXMAIFAC")
      Call CIERRARCH("OBSERVOF")
      '
End Function
Function VALOFORM(CAMPO$)
    '
    If STDSCREEN.APRO.Caption = "OK" Then
      For KKI% = 1 To CANOBJ%
        If TRIM$(UCase$(ONAME$(KKI%))) = TRIM$(UCase$(CAMPO$)) Then
          VALOFORM = Texto(KKI%)
          Exit Function
        End If
      Next KKI%
    End If
    '
    VALOFORM = ""
    '
End Function

Sub MUESTRAFOR()
   '
   If TOTHOJA& < 1 Then TOTHOJA& = 1
   If ORDHOJA& < 1 Then ORDHOJA& = 1
   '
10 ICATEX% = -1: HUBOCAI% = 0
   '
   For i% = 1 To CANOBJ%
      '
      MODOX$ = TIOBJ$(i%)
      OFXX = 100
      If 56.7 * H1(i%) > STDSCREEN.Picture1.Width - 20 Then
          STDSCREEN.Picture1.Width = 56.7 * H1(i%) + OFXX
      End If
      If 56.7 * H2(i%) > STDSCREEN.Picture1.Width - OFXX Then
          STDSCREEN.Picture1.Width = 56.7 * H2(i%) + OFXX
      End If
      If 56.7 * V1(i%) > STDSCREEN.Picture1.Height - OFXX Then STDSCREEN.Picture1.Height = 56.7 * V1(i%) + OFXX
      If 56.7 * V2(i%) > STDSCREEN.Picture1.Height - OFXX Then STDSCREEN.Picture1.Height = 56.7 * V2(i%) + OFXX
      '
      If STDSCREEN.Picture1.Width > 32000 Or STDSCREEN.Picture1.Height > 32000 Then
         Call MENSERR(24, "Objeto Número " & i% & " Fuera de Límites de la Hoja. Corrija.")
         Call FINAL("")
      End If
      '
      ' si está presentando el formulario pre-cargado, mantiene los valores
      If InStr(ARFOR$, ".SQX") < 1 Then
         TEXAN = Texto(i%)
         Call CARGATEXTO(i%, TOOLTI$)
         If Texto(i%) <> TEXAN Then HUBOCAI% = 1
      End If
      CARMAN%(i%) = 0
      '
      On Error Resume Next
      '
      BLOQ% = 0
      PPXX% = InStr(Texto(i%), "/LOCKED")
      If PPXX% > 0 Then
         BLOQ% = 1
         Texto(i%) = Left$(Texto(i%), PPXX% - 1)
      End If
      '
      Select Case MODOX$
         Case "LABEL"
             STDSCREEN.Picture1.CurrentY = 56.7 * V1(i%)
             STDSCREEN.Picture1.CurrentX = 56.7 * H1(i%)
             STDSCREEN.Picture1.Font = FUENTE$(i%)
             STDSCREEN.Picture1.FontSize = TAFUEN(i%)
             STDSCREEN.Picture1.FontBold = NEGRI(i%)
             STDSCREEN.Picture1.FontItalic = ITALI(i%)
             STDSCREEN.Picture1.FontUnderline = SUBRA(i%)
             If INVIS(i%) <> 1 Then
                If MULIN(i%) <> 1 Then
                     STDSCREEN.Picture1.Print Texto(i%)
                  Else
                     GoTo 17         ' LO TRATA COMO UN PARAMETRO
                End If
             End If
             '
          Case "PARAM"
17           NUPIC% = STDSCREEN.Label99.Count
             Load STDSCREEN.Label99(NUPIC%)
             STDSCREEN.Label99(NUPIC%).Top = 56.7 * V1(i%) - 50
             STDSCREEN.Label99(NUPIC%).Left = 56.7 * H1(i%)
             STDSCREEN.Label99(NUPIC%).Width = 56.7 * (H2(i%) - H1(i%))
             STDSCREEN.Label99(NUPIC%).Height = 56.7 * (V2(i%) - V1(i%)) ' * 0.85
             STDSCREEN.Label99(NUPIC%).Font = FUENTE$(i%)
             STDSCREEN.Label99(NUPIC%).FontSize = TAFUEN(i%)
             STDSCREEN.Label99(NUPIC%).FontBold = NEGRI(i%)
             STDSCREEN.Label99(NUPIC%).FontItalic = ITALI(i%)
             STDSCREEN.Label99(NUPIC%).FontUnderline = SUBRA(i%)
             STDSCREEN.Label99(NUPIC%).BorderStyle = ENCAS(i%)
             STDSCREEN.Label99(NUPIC%).BackStyle = ENCAS(i%)
             STDSCREEN.Label99(NUPIC%).Alignment = 0 ' IZQUIERDA
                LETFO$ = UCase$(Left$(FORPA$(i%), 1))
                If LETFO$ = "F" Or LETFO$ = "G" Or LETFO$ = "$" Or LETFO$ = "I" Then
                    STDSCREEN.Label99(NUPIC%).Alignment = 1 ' DERECHA
                  ElseIf LETFO$ = "D" Then
                    STDSCREEN.Label99(NUPIC%).Alignment = 2 ' DERECHA
                End If
             STDSCREEN.Label99(NUPIC%).ForeColor = RGB(0, 0, 0)
             ORILAB99%(NUPIC%) = i%                         ' movido aqui para que lo reconozca al suceso change
             IOBJELB%(i%) = NUPIC%
             If MULIN(i%) <> 1 Then
                   If InStr(FORPA$(i%), "^") > 0 Then
                      STDSCREEN.Label99(NUPIC%).Alignment = 2
                   End If
                   STDSCREEN.Label99(NUPIC%).Caption = FORMATNUM$(Texto(i%), FORPA$(i%))
                Else
                   STDSCREEN.Label99(NUPIC%).Caption = Texto(i%)
             End If
             If Left$(UCase$(App.EXEName), 6) = "DSGFRM" Then
                STDSCREEN.Label99(NUPIC%).ToolTipText = TOOLTI$
             End If
             STDSCREEN.Label99(NUPIC%).Visible = True
             If INVIS(i%) = 1 Then STDSCREEN.Label99(NUPIC%).Visible = False
             '
            'ACA MODIFICO -- DESCUENTO**
            If Left$(UCase$(FOREPLA$(i%)), 10) = "@DESCUENTO" Then
                textox$ = Mid$(FOREPLA$(i%), 2)
                A = Split(textox$, ">") 'FUNCION DEVUELVE UN VECTOR DE A CON LOS VALORES
                ENCONTRO1% = 0
                For LL = 1 To UBound(A)
                   NAMEPARAM$ = A(LL) 'EJ.:A(1)= {PARAM-2} , A(2)= {PARAM-1}
                   For jj% = 0 To STDSCREEN.Label99.Count
                      N_Objeto$ = "{" + ONAME$(jj%) + "}" ' LE DOY EL MISMO FORMATO
                      If N_Objeto$ = NAMEPARAM$ Then 'COMPARO SI ES EL MISMO NOMBRE PARA SABER EL INDICE REAL DE LABEL99
                         If ENCONTRO1% < 1 Then 'LEO EL PRIMER PARAMETRO (SIEMPRE DEBE IR PRIMERO EL IMPORTE
                           IMPORTEBRUTO# = XVALO(STDSCREEN.Label99(jj%).Caption)
                           ENCONTRO1% = 1
                           Exit For
                         Else
                           'LEO EL SEGUNDO PARAMETRO QUE VA SER EL TEXTO DEL DESCUENTO EJ.: 10+15
                           DESC$ = STDSCREEN.Label99(jj%).Caption
                           Exit For
                         End If
                      End If
                   Next jj%
                Next LL
                B = ImporteConDescuento#(IMPORTEBRUTO#, DESC$) 'ASIGNO A B EL IMPORTE CON EL DESCUENTO
                STDSCREEN.Label99(NUPIC%).Caption = FORMATNUM$(B, FORPA$(i%))
                A = COEFDESCU(SAFETEXT$(Texto(i%)))
            End If

          Case "TEXT"
             If MULIN(i%) <> 1 Then
                  NUPIC% = STDSCREEN.Text99.Count
                  Load STDSCREEN.Text99(NUPIC%)
                     Load STDSCREEN.Command99(NUPIC%)
                     Load STDSCREEN.ECO99(NUPIC%)
                  ICATEX% = ICATEX% + 1
                  STDSCREEN.Text99(NUPIC%).TabIndex = ICATEX%
                  STDSCREEN.Text99(NUPIC%).Top = 56.7 * V1(i%) - 50
                     STDSCREEN.Command99(NUPIC%).Top = STDSCREEN.Text99(NUPIC%).Top
                     STDSCREEN.ECO99(NUPIC%).Top = STDSCREEN.Text99(NUPIC%).Top
                  STDSCREEN.Text99(NUPIC%).Left = 56.7 * H1(i%)
                  STDSCREEN.Text99(NUPIC%).Width = 56.7 * (H2(i%) - H1(i%))
                     STDSCREEN.Command99(NUPIC%).Left = STDSCREEN.Text99(NUPIC%).Left + STDSCREEN.Text99(NUPIC%).Width
                     STDSCREEN.ECO99(NUPIC%).Left = STDSCREEN.Command99(NUPIC%).Left + STDSCREEN.Command99(NUPIC%).Width + 100
                  STDSCREEN.Text99(NUPIC%).FontSize = 1
                     STDSCREEN.Command99(NUPIC%).FontSize = 1
                  STDSCREEN.Text99(NUPIC%).Height = 56.7 * (V2(i%) - V1(i%)) * 0.85
                     STDSCREEN.Command99(NUPIC%).Height = STDSCREEN.Text99(NUPIC%).Height
                     STDSCREEN.ECO99(NUPIC%).Height = STDSCREEN.Text99(NUPIC%).Height
                  STDSCREEN.Text99(NUPIC%).Font = FUENTE$(i%)
                     STDSCREEN.ECO99(NUPIC%).Font = FUENTE$(i%)
                  STDSCREEN.Text99(NUPIC%).FontSize = TAFUEN(i%)
                     STDSCREEN.ECO99(NUPIC%).FontSize = TAFUEN(i%)
                     STDSCREEN.Command99(NUPIC%).FontSize = 5 + TAFUEN(i%)
                  STDSCREEN.Text99(NUPIC%).FontBold = NEGRI(i%)
                     STDSCREEN.ECO99(NUPIC%).FontBold = NEGRI(i%)
                     STDSCREEN.Command99(NUPIC%).FontBold = True
                  STDSCREEN.Text99(NUPIC%).FontItalic = ITALI(i%)
                     STDSCREEN.ECO99(NUPIC%).FontItalic = ITALI(i%)
                  STDSCREEN.Text99(NUPIC%).FontUnderline = SUBRA(i%)
                  STDSCREEN.Text99(NUPIC%).ForeColor = RGB(0, 0, 0)
                  If Left$(UCase$(FOREPLA$(i%)), 9) = "@PASSWORD" Then
                       STDSCREEN.Text99(NUPIC%).PasswordChar = "*"
                     ElseIf Left$(FOREPLA$(i%), 1) = "#" Then
                       STDSCREEN.Command99(NUPIC%).Visible = True
                       STDSCREEN.Command99(NUPIC%).ZOrder 0
                       If InStr(FOREPLA$(i%), ">") > 1 Then
                          STDSCREEN.ECO99(NUPIC%).Visible = True
                          STDSCREEN.ECO99(NUPIC%).ZOrder 0
                       End If
                     ElseIf Mid$(FOREPLA$(i%), 1, 8) = "FUNCTION" Then
                        STDSCREEN.ECO99(NUPIC%).Visible = True
                        STDSCREEN.ECO99(NUPIC%).ZOrder 0
                  End If
                     ' si está presentando el formulario pre-cargado, mantiene los valores
                     If InStr(ARFOR$, ".SQX") < 1 Then Call CARGATEXTO(i%, TOOLTI$)
                  ORITEXT99%(NUPIC%) = i%
                  IOBJETX%(i%) = NUPIC%
                  STDSCREEN.Text99(NUPIC%).TEXT = FORMATNUM(Texto(i%), FORPA$(i%))
                  If Left$(UCase$(App.EXEName), 6) = "DSGFRM" Then
                     STDSCREEN.Text99(NUPIC%).ToolTipText = TOOLTI$
                  End If
                  STDSCREEN.Text99(NUPIC%).Visible = True
                  If INVIS(i%) = 1 Then STDSCREEN.Text99(NUPIC%).Visible = False
                  If BLOQ% > 0 Then
                      STDSCREEN.Text99(NUPIC%).Enabled = False
                    Else
                      If InStr(Texto(i%), "/") > 0 Then
                          STDSCREEN.Text99(NUPIC%).TEXT = ""     ' LOS CONMUTANTES SE PRESENTAN EN BLANCO
                        ElseIf Left$(FORPA$(i%), 1) = "D" Or Left$(FORPA$(i%), 1) = "H" Or Left$(FORPA$(i%), 1) = "T" Then
                           If Texto(i%) = TMASCARA$(FORPA$(i%)) Then
                              STDSCREEN.Text99(NUPIC%).TEXT = ""     ' FECHA Y HORA SE PRESENTAN EN BLANCO
                           End If
                      End If
                  End If
                Else
                  NUPIC% = STDSCREEN.Text99M.Count
                  Load STDSCREEN.Text99M(NUPIC%)
                  STDSCREEN.Text99M(NUPIC%).Top = 56.7 * V1(i%) - 50
                  STDSCREEN.Text99M(NUPIC%).Left = 56.7 * H1(i%)
                  STDSCREEN.Text99M(NUPIC%).Width = 56.7 * (H2(i%) - H1(i%))
                  STDSCREEN.Text99M(NUPIC%).FontSize = 1
                  STDSCREEN.Text99M(NUPIC%).Height = 56.7 * (V2(i%) - V1(i%))
                  STDSCREEN.Text99M(NUPIC%).Font = FUENTE$(i%)
                  STDSCREEN.Text99M(NUPIC%).FontSize = TAFUEN(i%)
                  STDSCREEN.Text99M(NUPIC%).FontBold = NEGRI(i%)
                  STDSCREEN.Text99M(NUPIC%).FontItalic = ITALI(i%)
                  STDSCREEN.Text99M(NUPIC%).FontUnderline = SUBRA(i%)
                  STDSCREEN.Text99M(NUPIC%).ForeColor = RGB(0, 0, 0)
                     ' si está presentando el formulario pre-cargado, mantiene los valores
                     If InStr(ARFOR$, ".SQX") < 1 Then Call CARGATEXTO(i%, TOOLTI$)
                     TTXX$ = Texto(i%)
                     If Left$(FORPA$(i%), 1) = "X" Then
                        Moneda$ = DEMONECO$(MONEMI%)
                        TTXX$ = TEXTFORM$(TTXX$, FORPA$(i%), Moneda$)
                     ElseIf Left$(FORPA$(i%), 1) = "Z" Then
                        'MONEDA$ = DEMONECO$(MONEMI%)
                        TTXX$ = TEXTFORM$(TTXX$, FORPA$(i%), "", "", -1) 'FORMATO Z NO IMPRIME LA MONEDA EN EL IMPORTE EN LETRAS
                     End If
                     
                  STDSCREEN.Text99M(NUPIC%).TEXT = RTrim(LTrim(TTXX$))    ' TEXTO(I%)
                  ORITEXT99M%(NUPIC%) = i%
                  IOBJETXM%(i%) = NUPIC%
                  STDSCREEN.Text99M(NUPIC%).Visible = True
                  If INVIS(i%) = 1 Then STDSCREEN.Text99M(NUPIC%).Visible = False
                  If BLOQ% > 0 Then STDSCREEN.Text99M(NUPIC%).Enabled = False
             End If
          Case "LINE"
             STDSCREEN.Picture1.DrawStyle = 0
             STDSCREEN.Picture1.DrawWidth = XVALO(TAFUEN$(i%))
             If INVIS(i%) <> 1 Then STDSCREEN.Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%))
          Case "BOX"
             STDSCREEN.Picture1.DrawStyle = 0
             STDSCREEN.Picture1.DrawWidth = XVALO(TAFUEN$(i%))
             If INVIS(i%) <> 1 Then STDSCREEN.Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%)), , B
          Case "GRAPH"
             If InStr(ARFOR$, ".SQX") < 1 Then
                Call CARGATEXTO(i%, TOOLTI$) ' CARGA LA IMAGEN A PARTIR DE LA CONSULTA SQL
             End If
             If INVIS(i%) <> 1 Then
                If Texto(i%) <> "" Then
                   STDSCREEN.Picture1.DrawStyle = 0
                   STDSCREEN.Picture1.DrawWidth = 1
                   If ENCAS(i%) > 0 Then
                       STDSCREEN.Picture1.Line (56.7 * H1(i%), 56.7 * V1(i%))-(56.7 * H2(i%), 56.7 * V2(i%)), , B
                   End If
                   TTXZ$ = TRIM$(UCase$(Texto(i%)))
                   ' *** @CAPTURA ES IMAGEN A PARTIR DE WEBCAM
                   If TTXZ$ = "@IMAGEN" Or TTXZ$ = "@FOTO" Or TTXZ$ = "@PICTURE" Or Mid$(TTXZ$, 1, 8) = "@CAPTURA" Then
                        STDSCREEN.Picture1.Font = "Arial"
                        STDSCREEN.Picture1.FontSize = 8
                        STDSCREEN.Picture1.FontBold = True
                        STDSCREEN.Picture1.FontItalic = False
                        STDSCREEN.Picture1.FontUnderline = False
                        STDSCREEN.Picture1.CurrentY = 56.7 * ((V1(i%) + V2(i%)) / 2 - 4)
                        STDSCREEN.Picture1.CurrentX = 56.7 * ((H1(i%) + H2(i%)) / 2 - 8)
                        STDSCREEN.Picture1.Print "Doble Click"
                        STDSCREEN.Picture1.CurrentY = 56.7 * ((V1(i%) + V2(i%)) / 2)
                        STDSCREEN.Picture1.CurrentX = 56.7 * ((H1(i%) + H2(i%)) / 2 - 10)
                        STDSCREEN.Picture1.Print "Para Capturar"
                        STDSCREEN.Picture1.CurrentY = 56.7 * ((V1(i%) + V2(i%)) / 2 + 4)
                        STDSCREEN.Picture1.CurrentX = 56.7 * ((H1(i%) + H2(i%)) / 2 - 5)
                        STDSCREEN.Picture1.Print "Imagen"
                      Else
                        ' *** IMAGEN A UNA RUTA GUARDADA EN UNA TABLA
                        ' *** LA RUTA SE OBTIENE EN CARGATEXTO()
                        ARIMAG$ = Texto(i%)
                        If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = App.Path + "\" + Texto(i%)
                        If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = Texto(i%)
                        If EXISTE%(ARIMAG$) > 0 Then
                           Call DIBUIMAG(STDSCREEN.Picture1, ARIMAG$, H1(i%), H2(i%), V1(i%), V2(i%))
                           STDSCREEN.btnQuitaFotoCapturada.Visible = True
                           STDSCREEN.btnQuitaFotoCapturada.Left = H1(i%) * 56.7
                           STDSCREEN.btnQuitaFotoCapturada.Top = V1(i%) * 56.7
                        End If
                   End If
                End If
             End If
          Case "GRID"
             STDSCREEN.Picture1.DrawStyle = 4
             STDSCREEN.Picture1.DrawWidth = 1
             If INVIS(i%) <> 1 Then
                NUPIC% = STDSCREEN.GRID99.Count
                Load STDSCREEN.GRID99(NUPIC%)
                STDSCREEN.GRID99(NUPIC%).Visible = True
                STDSCREEN.GRID99(NUPIC%).Top = 56.7 * V1(i%) - 50
                STDSCREEN.GRID99(NUPIC%).Left = 56.7 * H1(i%)
                STDSCREEN.GRID99(NUPIC%).Width = 56.7 * (H2(i%) - H1(i%))
                STDSCREEN.GRID99(NUPIC%).Height = 56.7 * (V2(i%) - V1(i%))
                STDSCREEN.Refresh
                '
                APLINVO$ = Texto(i%)
                If TRIM$(APLINVO$) <> "" Then
                   APLIXI$ = APLINVO$
                   TEGRI$ = ""
                   PPXX% = InStr(APLINVO$, "///TEGRID =")
                   If PPXX% > 0 Then
                      APLIXI$ = Left$(APLINVO$, PPXX% - 1)
                      TEGRI$ = Mid$(APLINVO$, PPXX% + 11)
                   End If
                   '
                   FILTXI$ = ""
                   PPXX% = InStr(APLIXI$, "(")
                   If PPXX% > 0 Then
                      PPYY% = InStr(PPXX%, APLIXI$, ")")
                      If PPYY% <= PPXX% Then PPYY% = 1 + Len(APLIXI$)
                      FILTXI$ = Mid$(APLIXI$, PPXX% + 1, PPYY% - PPXX% - 1)
                      APLIXI$ = Left$(APLIXI$, PPXX% - 1) + Mid$(APLIXI$, PPYY% + 1)
                   End If
                   '
                   Call ABRECONEXION
                   On Error GoTo 0
                   On Error Resume Next
                   SHOWREP07.Label2 = ""
                   If Err Then
                      On Error GoTo 0
                      Call MENSERR(24, "Falta Formulario 'SHOWREP07.FRM' en Proyecto '" + App.EXEName + "'.")
                      Call FINAL("")
                   End If
                   On Error GoTo 0
                   '
                   SHOWREP07.GRID.Font = FUENTE$(i%)
                      If XVALO(TAFUEN(i%)) < 4 Then TAFUEN(i%) = 4
                   SHOWREP07.GRID.FontSize = TAFUEN(i%)
                   SHOWREP07.GRID.FontBold = NEGRI(i%)
                   SHOWREP07.GRID.FontItalic = ITALI(i%)
                   SHOWREP07.GRID.FontUnderline = SUBRA(i%)
                   '
                   STDSCREEN.GRID99(NUPIC%).Font = FUENTE$(i%)
                   STDSCREEN.GRID99(NUPIC%).FontSize = TAFUEN(i%)
                   STDSCREEN.GRID99(NUPIC%).FontBold = NEGRI(i%)
                   STDSCREEN.GRID99(NUPIC%).FontItalic = ITALI(i%)
                   STDSCREEN.GRID99(NUPIC%).FontUnderline = SUBRA(i%)

                   '
                   CARGAIMAGENCOMOLINK% = 1
                   FILTX$ = TEXREPLA(FILTXI$)
                   Call CARGA_CONSULTA(APLIXI$)
                   SHOWREP07.Label2 = APLIXI$
                   SHOWREP07.Label3 = DESQUERY$     ' ESTO HAY QUE REVISARLO
                   CARGAIMAGENCOMOLINK% = 0
                   STDSCREEN.GRID99(NUPIC%).Clear
                   STDSCREEN.GRID99(NUPIC%).Cols = SHOWREP07.GRID.Cols
                   
                   ' ajusta altura de la primera fila
                   STDSCREEN.GRID99(NUPIC%).RowHeight(0) = 450
                   '
                   ' ajusta cantidad de filas
                   STDSCREEN.GRID99(NUPIC%).Rows = SHOWREP07.GRID.Rows
                   ALTUTO = STDSCREEN.GRID99(NUPIC%).RowHeight(0)
                   For ROWI& = 2 To STDSCREEN.GRID99(NUPIC%).Rows
                      ALTUTO = ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1)
                   Next ROWI&
                   '
                   If STDSCREEN.GRID99(NUPIC%).Rows > 1 Then
                      Do While ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1) < STDSCREEN.GRID99(NUPIC%).Height
                         STDSCREEN.GRID99(NUPIC%).Rows = STDSCREEN.GRID99(NUPIC%).Rows + 1
                         ALTUTO = ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1)
                      Loop
                   End If
                   '
                   ' ajusta al ancho disponible
                   ANCOLU = 0
                   ANDISPO = STDSCREEN.GRID99(NUPIC%).Width - 110
                   For COLI& = 1 To STDSCREEN.GRID99(NUPIC%).Cols
                      STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1) = SHOWREP07.GRID.ColWidth(COLI& - 1)
                      STDSCREEN.GRID99(NUPIC%).ColAlignment(COLI& - 1) = SHOWREP07.GRID.ColAlignment(COLI& - 1)
                      ANCOLU = ANCOLU + STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1)
                   Next COLI&
                   ALTUGRID = STDSCREEN.GRID99(NUPIC%).Height
                   If ALTUTO > ALTUGRID Then
                      ANDISPO = ANDISPO - 300
                   End If
                   COEF = ANDISPO / ANCOLU
                   'se valida por que si coef es negativo genera error 30014 el ancho de columna no es valido
                   If XVALO(COEF) < 0 Then
                      N_Objeto$ = ONAME$(i%)
                      N_FORMUL$ = REPLACAR$(ARFOR$, "\", "/")
                      Call MENSERR(24, "Error de Dimensionamiento del Objeto " & N_Objeto$ & "\ En el Formulario " & N_FORMUL$)
                      GoTo 70
                   End If
                   '
                   For COLI& = 1 To STDSCREEN.GRID99(NUPIC%).Cols
                      STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1) = COEF * STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1)
'                      AA1 = STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1)
                   Next COLI&
                   '
                   ' CARGA LOS VALORES
                   CANREGRID%(NUPIC%) = 1
                   If TEGRI$ = "" Then
                       For ROWI& = 1 To SHOWREP07.GRID.Rows
                         If ROWI& > CANREGRID%(NUPIC%) Then CANREGRID%(NUPIC%) = ROWI&
                         ALTUREMAX = 0
                         For COLI& = 1 To SHOWREP07.GRID.Cols
                           Call CARGALAIMAGEN(NUPIC%, ROWI& - 1, COLI& - 1)
                           If ALTUREMAX > STDSCREEN.GRID99(NUPIC%).RowHeight(ROWI& - 1) Then
                              STDSCREEN.GRID99(NUPIC%).RowHeight(ROWI& - 1) = ALTUREMAX
                           End If
                           STDSCREEN.GRID99(NUPIC%).TextMatrix(ROWI& - 1, COLI& - 1) = SHOWREP07.GRID.TextMatrix(ROWI& - 1, COLI& - 1)
                         Next COLI&
                       Next ROWI&
                     Else
                       Call CARGATEGRI$(TEGRI$, NUPIC%)
                   End If
                   '
                   ' SE AGREGA PUNTUALMENTE PARA EL DETALLE DE UNA FACTURA, APLICABLE CUANDO SE HA DEPURADO LA TABLA DE MOVIMIENTOS
                   ' DE STOCK
                   If CONSIDERA_MOVIVIE = 1 Then
                       '
                       ROWIX& = SHOWREP07.GRID.Rows + 1
                       FILTX$ = TEXREPLA(FILTXI$)
                       Call CARGA_CONSULTA("DETFACA1")
                       SHOWREP07.GRID.Clear
                       SHOWREP07.Label2 = "DETFACA1"
                       SHOWREP07.Label3 = DESQUERY$     ' ESTO HAY QUE REVISARLO
                       CARGAIMAGENCOMOLINK% = 0
                       '
                       ' ajusta altura de la primera fila
                       STDSCREEN.GRID99(NUPIC%).RowHeight(0) = 450
                       '
                       ' ajusta cantidad de filas
                       STDSCREEN.GRID99(NUPIC%).Rows = STDSCREEN.GRID99(NUPIC%).Rows + SHOWREP07.GRID.Rows
                       ALTUTO = STDSCREEN.GRID99(NUPIC%).RowHeight(0)
                       For ROWI& = 2 To STDSCREEN.GRID99(NUPIC%).Rows
                          ALTUTO = ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1)
                       Next ROWI&
                       '
                       If STDSCREEN.GRID99(NUPIC%).Rows > 1 Then
                          Do While ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1) < STDSCREEN.GRID99(NUPIC%).Height
                             STDSCREEN.GRID99(NUPIC%).Rows = STDSCREEN.GRID99(NUPIC%).Rows + 1
                             ALTUTO = ALTUTO + 10 + STDSCREEN.GRID99(NUPIC%).RowHeight(1)
                          Loop
                       End If
                       '
                       ' ajusta al ancho disponible
                       ANCOLU = 0
                       ANDISPO = STDSCREEN.GRID99(NUPIC%).Width - 110
                       For COLI& = 1 To STDSCREEN.GRID99(NUPIC%).Cols
                          STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1) = SHOWREP07.GRID.ColWidth(COLI& - 1)
                          STDSCREEN.GRID99(NUPIC%).ColAlignment(COLI& - 1) = SHOWREP07.GRID.ColAlignment(COLI& - 1)
                          ANCOLU = ANCOLU + STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1)
                       Next COLI&
                       ALTUGRID = STDSCREEN.GRID99(NUPIC%).Height
                       If ALTUTO > ALTUGRID Then
                          ANDISPO = ANDISPO - 300
                       End If
                       COEF = ANDISPO / ANCOLU
                       'se valida por que si coef es negativo genera error 30014 el ancho de columna no es valido
                       If XVALO(COEF) < 0 Then
                          N_Objeto$ = ONAME$(i%)
                          N_FORMUL$ = REPLACAR$(ARFOR$, "\", "/")
                          Call MENSERR(24, "Error de Dimensionamiento del Objeto " & N_Objeto$ & "\ En el Formulario " & N_FORMUL$)
                          GoTo 70
                       End If
                       '
                       For COLI& = 1 To STDSCREEN.GRID99(NUPIC%).Cols
                          STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1) = COEF * STDSCREEN.GRID99(NUPIC%).ColWidth(COLI& - 1)
                       Next COLI&
                       For ROWI& = 2 To SHOWREP07.GRID.Rows
                         If ROWI& > CANREGRID%(NUPIC%) Then CANREGRID%(NUPIC%) = ROWI&
                         ALTUREMAX = 0
                         For COLI& = 1 To SHOWREP07.GRID.Cols
                           Call CARGALAIMAGEN(NUPIC%, ROWIX& - 1, COLI& - 1)
                           If ALTUREMAX > STDSCREEN.GRID99(NUPIC%).RowHeight(ROWIX& - 1) Then
                              STDSCREEN.GRID99(NUPIC%).RowHeight(ROWIX& - 1) = ALTUREMAX
                           End If
                           STDSCREEN.GRID99(NUPIC%).TextMatrix(ROWIX& - 1, COLI& - 1) = SHOWREP07.GRID.TextMatrix(ROWI& - 1, COLI& - 1)
                         Next COLI&
                         ROWIX& = ROWIX& + 1
                       Next ROWI&
                   End If
                   '
                   On Error Resume Next
                   Unload SHOWREP07         ' PARA QUE NO APAREZCA VISIBLE AL OCULTAR EL STDSCREEN
                   On Error GoTo 0
                   '
                End If
                STDSCREEN.Refresh
             End If
      End Select
      STDSCREEN.Picture1.Refresh
      '
70 Next i%
   '
   ' If HUBOCAI% > 0 Then GoTo 10
   ' IVUELTA% = IVUELTA% + 1
   ' If IVUELTA% < 2 Then GoTo 10
   '
   HSTAN = STDSCREEN.Picture1.Height
   If PAGI_ANT$ = "SI" Then
      STDSCREEN.PAGANT.Visible = True
      STDSCREEN.PAGANT1.Visible = True
      STDSCREEN.PAGANT.Top = HSTAN + 150
      STDSCREEN.PAGANT.Left = STDSCREEN.Picture1.Width / 2 - 100 - STDSCREEN.PAGANT.Width
      STDSCREEN.Picture1.Height = HSTAN + 150 + STDSCREEN.PAGANT.Height + 150
      STDSCREEN.Picture1.DrawStyle = 0
      STDSCREEN.Picture1.DrawWidth = 2
      STDSCREEN.Picture1.Line (0, HSTAN + 50)-(32767, HSTAN + 50)
   End If
   '
   If PAGI_SIG$ = "SI" Or AUTOPAGE% > 0 Then
      STDSCREEN.PAGSIG.Visible = True
      STDSCREEN.PAGSIG1.Visible = True
      STDSCREEN.PAGSIG.Top = HSTAN + 150
      STDSCREEN.PAGSIG.Left = STDSCREEN.Picture1.Width / 2 + 100
      STDSCREEN.Picture1.Height = HSTAN + 150 + STDSCREEN.PAGSIG.Height + 150
      STDSCREEN.Picture1.DrawStyle = 0
      STDSCREEN.Picture1.DrawWidth = 2
      STDSCREEN.Picture1.Line (0, HSTAN + 50)-(32767, HSTAN + 50)
   End If
   '
   
   AFORMU = STDSCREEN.Picture2.Left + STDSCREEN.Picture1.Width + STDSCREEN.VScroll1.Width + 70
   If AFORMU > Screen.Width - 100 Then AFORMU = Screen.Width - 100
   STDSCREEN.Width = AFORMU + 30
   STDSCREEN.Picture2.Width = STDSCREEN.ScaleWidth - STDSCREEN.Picture2.Left
   '
   HFORMU = STDSCREEN.Picture2.Top + STDSCREEN.Picture1.Height + STDSCREEN.HScroll1.Height
   'STDSCREEN.Picture2.Height = HFORMU
   If HFORMU > Screen.Height - 1000 Then HFORMU = Screen.Height - 1000
   STDSCREEN.Height = HFORMU + 500
   STDSCREEN.Picture2.Height = STDSCREEN.ScaleHeight - STDSCREEN.Picture2.Top
   '
   STDSCREEN.VScroll1.Top = 0    ' STDSCREEN.Picture1.Top
   STDSCREEN.VScroll1.Left = STDSCREEN.Picture2.Width - STDSCREEN.VScroll1.Width - 30
   STDSCREEN.VScroll1.Height = STDSCREEN.Picture2.Height - STDSCREEN.HScroll1.Height
   STDSCREEN.VScroll1.Value = 0
   STDSCREEN.VScroll1.Max = 0
   If STDSCREEN.Picture1.Height > STDSCREEN.VScroll1.Height Then
      STDSCREEN.VScroll1.Max = STDSCREEN.Picture1.Height - STDSCREEN.VScroll1.Height
   End If
   If STDSCREEN.VScroll1.Max = 0 Then STDSCREEN.Width = STDSCREEN.Width - STDSCREEN.VScroll1.Width - 40
   '
   On Error Resume Next
      STDSCREEN.VScroll1.Value = -30 - STDSCREEN.Picture1.Top + STDSCREEN.PINVER
   On Error GoTo 0
   '
   If POVERTI > 0 Then
      If POVERTI > STDSCREEN.VScroll1.Max Then POVERTI = STDSCREEN.VScroll1.Max
      STDSCREEN.VScroll1.Value = POVERTI
      Call STDSCREEN.VScroll1_Scroll
   End If
   '
   STDSCREEN.HScroll1.Left = STDSCREEN.Picture1.Left
   STDSCREEN.HScroll1.Top = STDSCREEN.Picture2.Height - STDSCREEN.HScroll1.Height - 30 ' - 100
   STDSCREEN.HScroll1.Width = STDSCREEN.Picture2.Width - STDSCREEN.VScroll1.Width
   STDSCREEN.HScroll1.Value = 0
   STDSCREEN.HScroll1.Max = 0
   If STDSCREEN.Picture1.Width > STDSCREEN.HScroll1.Width Then
      STDSCREEN.HScroll1.Max = STDSCREEN.Picture1.Width - STDSCREEN.HScroll1.Width
   End If
   If STDSCREEN.HScroll1.Max = 0 Then STDSCREEN.Height = STDSCREEN.Height - STDSCREEN.HScroll1.Height - 100
   '
   STDSCREEN.Picture8.Left = STDSCREEN.ScaleWidth - STDSCREEN.Picture8.Width
   STDSCREEN.Picture8.BackColor = STDSCREEN.BackColor
   '
   Call CENTRAFORM(STDSCREEN)
   On Error GoTo 0
   '
End Sub

Sub CARGALAIMAGEN(NPC%, RGX&, CLX&)
    '
    TXTXZ$ = SHOWREP07.GRID.TextMatrix(RGX&, CLX&)
    If Left$(TXTXZ$, 1024) <> Space$(1024) Then Exit Sub
    '
    PPXX1% = InStr(TXTXZ$, "///TXT =")
    If PPXX1% > 0 Then
       Call FRATEXTO(Mid$(TXTXZ$, PPXX1% + 8), 128)
       Call SHOWREP07.PONETEXMULINENCELDA(STDSCREEN.GRID99(NPC%), RGX&, CLX&)
       Exit Sub
    End If
    '
    If InStr(TXTXZ$, ".JPG") > 0 Then
       ARIMAG$ = TRIM$(SHOWREP07.GRID.TextMatrix(RGX&, CLX&))
       Call SHOWREP07.PONEIMAGENENCELDA(STDSCREEN.GRID99(NPC%), RGX&, CLX&, ARIMAG$, "")
    End If
    '
End Sub

Function VALHOJANTE$(TTXXZZ$)
   '
   Call ABREORFAB
   OF$ = TEXREPLA$("{COMPAR-0}")
   IDHOJA$ = Mid$(TTXXZZ$, 2)
   PPXX% = InStr(IDHOJA$, "%"): If PPXX% < 1 Then Exit Function
   IdCampo$ = TRIM$(Mid$(IDHOJA$, PPXX% + 1))
   IDHOJA$ = Left$(IDHOJA$, PPXX% - 1)
   '
   CONDI$ = "IdSubOr = '" & OF$ & "' "
   CIXI% = XVALO(VALOBADA("OPERFAB", "Cod_Inte", CONDI$, "NOMESS"))
   '
   CONDI$ = "CODICONJ = " & CIXI% & " AND HoProFab = '" & IDHOJA$ & "'"
   NOPEA% = XVALO(VALOBADA("SECOPER", "NUMEOPER", CONDI$, "NOMESS"))
   '
   VAGUARDB$ = ""
   CONDI$ = "IdSubOr = '" & OF$ & "' AND NumeOper = " & NOPEA%
   VAGUARDA$ = VALOBADA("OPERFAB", "HoBatchRec", CONDI$, "NOMESS")
   IDCAMPX$ = TRIM$(REPLACAR$(IdCampo$, "{", ""))
   IDCAMPY$ = TRIM$(REPLACAR$(IDCAMPX$, "}", ""))
   PPXX% = InStr(VAGUARDA, "NAME = " & IDCAMPY$ + Chr$(124))
   If PPXX% > 0 Then
      PPYY% = InStr(PPXX% + 1, VAGUARDA$, "TEXT = ")
      PPTT% = InStr(PPXX% + 1, VAGUARDA$, "TIPO = ")
         If PPTT% < PPXX% + 1 Then PPTT% = 1 + Len(VAGUARDA$)
      If PPYY% > PPXX% Then
        If PPYY% < PPTT% Then
          PPZZ% = InStr(PPYY% + 1, VAGUARDA$, Chr$(124))
          If PPZZ% > PPYY% Then
            VAGUARDB$ = TRIM$(Mid$(VAGUARDA$, PPYY% + 6, PPZZ% - PPYY% - 6))
            PPTT% = InStr(UCase$(VAGUARDB$), "/LOCKED")
            If PPTT% > 0 Then VAGUARDB$ = Left$(VAGUARDB$, PPTT% - 1)
          End If
        End If
      End If
   End If
   '
   VALHOJANTE$ = VAGUARDB$
   '
End Function
'
Function TEXREPLA$(TEXORI$, Optional ENTRECOMI%)
     '
'SE DESHABILITA PROVISORIAMENTE A VER SI QUEDA CON LA VERSION ANTERIOR
     If Left$(TEXORI$, 1) = "%" Then
        TREP$ = VALHOJANTE$(TEXORI$)
        GoTo 90
     End If
'     '
     TREP$ = TEXORI$
     If UCase$(TEXORI$) = "@USERNAME" Then
          TREP$ = USERNAME$: GoTo 90
     End If
     '
     If UCase$(TEXORI$) = "@HOY" Then
          H0II% = HOY(TREP$): GoTo 90
     End If
     '
     If Left$(TEXORI$, 1) = "@" Then
          TREP$ = ""
          GoTo 90
     End If
    '
10   PPXX1% = InStr(TREP$, "{")
     If PPXX1% > 0 Then      ' BUSCA ENCIERRO ENTRE LLAVES
        PPXX2% = InStr(TREP$, "}"): If PPXX2% <= PPXX1% Then PPXX2% = 1 + Len(TREP$)
        IZQUI$ = Left$(TREP$, PPXX1% - 1)
        DEREX$ = Mid$(TREP$, PPXX2% + 1)
        VAREPLA$ = Mid$(TREP, PPXX1% + 1, PPXX2% - PPXX1% - 1)
          If UCase$(Left(VAREPLA$, 7)) = "COMMPAR" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 8)
          If UCase$(Left$(VAREPLA$, 6)) = "PARCOM" Then VAREPLA$ = "COMPAR" + Mid$(VAREPLA$, 7)
          If VAREPLA$ = "COMPAR" Then VAREPLA$ = "COMPAR-0"
        For jj& = 1 To CANOBJ%
           If UCase$(ONAME$(jj&)) = UCase$(VAREPLA$) Then
              VAREPLA = Texto(jj&)
              ' *** SE LLEVA EL FORMATO 'G' AL 'F', PARA LUEGO PODER HACER OPERACIONES ALGEBRAICAS
              If Left$(FORPA$(jj&), 1) = "G" Then
                VAREPLA = REPLACAR$(VAREPLA, ".", "")
                VAREPLA = REPLACAR$(VAREPLA, ",", ".")
              End If
              If VAREPLA = TMASCARA$(FORPA$(jj&)) Then VAREPLA = "" ' Agregado por EPB para suprimir mascara en test
              If VAREPLA$ = "" Then
                 FORIXX$ = TRIM$(Left$(UCase$(FORPA$(jj&)), 1))
                 If FORIXX$ = "F" Or FORIXX$ = "G" Or FORIXX$ = "$" Or FORIXX$ = "I" Then
                    VAREPLA$ = "0"
                 End If
              End If
              GoTo 29
           End If
        Next jj&
        '
29      comilla$ = "": If ENTRECOMI% = 1 Then comilla$ = "'"
        TREP$ = IZQUI$ + comilla$ + VAREPLA$ + comilla$ + DEREX$
        GoTo 10
     End If
     '
     PPXX% = InStr(TEXORI$, "(")
     If PPXX% > 0 Then
       PPXX% = InStr(TREP$, "(")
       If PPXX% > 0 Then
         XYFUN$ = UCase$(TRIM$(Left$(TREP$, PPXX% - 1)))
         AYFUN$ = TRIM$(REPLACAR$(Mid$(TREP$, PPXX% + 1), ")", ""))
           ' funciones de clientes
         If XYFUN$ = "RASOCLI" Then
             TREP$ = rasocli$(XVALO(AYFUN))
           ElseIf XYFUN$ = "TELECLI" Then
             TREP$ = TEL1CLI$(XVALO(AYFUN))
           ElseIf XYFUN$ = "EMAILCLI" Then
             TREP$ = EMAILCLI$(XVALO(AYFUN))
             '
             ' funciones de empleados
           ElseIf XYFUN$ = "NULEGAJO" Or XYFUN$ = "LEGAJO" Then
             TREP$ = TRIM$(Str$(LEGA_OPER&(AYFUN)))
             If AYFUN$ <> "" And TREP$ = "" Then
                Call MENSERR(24, "No se ha Encontrado 'LEGAJO' para Operador '" & AYFUN & "'.")
             End If
           ElseIf XYFUN$ = "APEYNOM" Or XYFUN$ = "NOMBRE" Or XYFUN$ = "APENOM" Or XYFUN$ = "APYNOM" Then
             TREP$ = NOME_OPER$(AYFUN)
             If AYFUN$ <> "" And TREP$ = "" Then
                Call MENSERR(24, "No se ha Encontrado 'APELLIDO Y NOMBRE' para Operador '" & AYFUN & "'.")
             End If
           ElseIf XYFUN$ = "COCATEG" Then
             TREP$ = CCAT_OPER$(AYFUN)
             AYFX$ = AYFUN$
             PPXY% = InStr(AYFUN$, ">"): If PPXY% > 0 Then AYFX$ = Left$(AYFUN$, PPXY% - 1)
             If AYFX$ <> "" And TREP$ = "" Then
                Call MENSERR(24, "No se ha Encontrado 'CODIGO-CATEGORIA' para Operador '" & AYFX & "'.")
             End If
           ElseIf XYFUN$ = "DECATEG" Or XYFUN$ = "CATEGO" Then
             TREP$ = DCAT_OPER$(AYFUN)
             If AYFUN$ <> "" And TREP$ = "" Then
                Call MENSERR(24, "No se ha Encontrado 'CATEGORIA' para Operador '" & AYFUN & "'.")
             End If
             '
           Else
             ' ESTO DEBE QUEDAR COMENTADO ... SI NO DA UN FALSO MENSAJE DE ERROR
             ' Call MENSERR(24, "Referencia a función Inexistente '" + XYFUN$ + "'.")
         End If
       End If
     End If
     '
90   TEXREPLA = TREP$
     '
End Function
'
Sub CARGATEXTO(i%, TOOLTI$)
   '
   TOOLTI$ = ""
   ' *** @CAPTURA ES IMAGEN A PARTIR DE WEBCAM
   If TIOBJ$(i%) = "PARAM" Or TIOBJ$(i%) = "TEXT" Or (TIOBJ$(i%) = "GRAPH" And Mid$(Texto(i%), 1, 8) = "@CAPTURA") Then
      ' primero trata de levantar de tabla destino
      If TASEL$(i%) <> "" Then
         CONDIX$ = TEXREPLA$(CONDSEL$(i%), 1)
         VAREPLA = VALOBADA(TASEL$(i%), CASEL$(i%), CONDIX$)
         '
         ' ESTO QUE SIGUE SE AGREGO EL 6 DE OCTUBRE DE 2011 PARA QUE EN LA BL140
         ' DE LIMPIEZA NO TRAIGA EL NUMERO DE LOTE CORRESPONDIENTE A UNA SF GENE-
         ' RADA SIN NUMERO DE ORDEN DE PRODUCCION.
         CONDIXU$ = UCase$(REPLACAR$(CONDIX$, " ", ""))
         If Right$(CONDIXU$, 3) = "=''" And InStr(CONDIXU$, "AND") < 1 Then
            VAREPLA = ""
         End If
         
         TOOLTI$ = "SELECT " + CASEL$(i%) + " FROM " + TASEL$(i%) + " WHERE " + CONDIX$
         If VarType(VAREPLA) = 7 Then
              VAREPLA = CVINT(VAREPLA)         ' CASO FECHA
            Else
              TVAR$ = UCase$(Left$(FORPA$(i%), 1))
              If TVAR$ = "I" Or TVAR$ = "F" Or TVAR$ = "G" Or TVAR$ = "$" Then
                 VAREPLA = FORMATNUM$(VAREPLA, FORPA$(i%))
              End If
         End If
         ' *** IMAGEN A PARTIR DE WEBCAM
         ' *** SE AGREGA PARA QUE VAREPLA = '' NO PISE LA BANDERA @CAPTURA AL APLICAR MAS ADELANTE Texto(i%) = VAREPLA
         If TIOBJ$(i%) = "GRAPH" And Mid$(Texto(i%), 1, 8) = "@CAPTURA" And VAREPLA = "" Then
            VAREPLA = Texto(i%)
         ElseIf TIOBJ$(i%) = "GRAPH" And Mid$(Texto(i%), 1, 8) = "@CAPTURA" And VAREPLA <> "" Then
            STDSCREEN.btnQuitaFotoCapturada.Tag = CStr(i%) & ";" & Texto(i%) ' GUARDA LOS DATOS NECESARIOS PARA ELIMINAR LA IMAGEN
         End If
         
         Texto(i%) = VAREPLA    ' FORMATNUM$(VAREPLA, FORPA$(I%))
         If TIOBJ$(i%) = "GRAPH" And TRIM$(VAREPLA) = "" Then
            Texto(i%) = "@IMAGEN"
         End If
         If Texto(i%) <> "" Then Exit Sub
      End If
      '
      ' si no encontro el valor toma del origen
      FOREP$ = LTrim$(RTrim$(FOREPLA$(i%)))
      If FOREP$ = "@HOY" Then FOREP$ = "#FECHA"
      If FOREP$ <> "" Then
         TOOLTI$ = FOREP$
         ' numeracion de hojas
         If InStr(UCase$(FOREP$), "ORDHOJA") > 0 Then FOREP$ = TRIM$(Str$(ORDHOJA&)): GoTo 29
         If InStr(UCase$(FOREP$), "TOTHOJA") > 0 Then FOREP$ = TRIM$(Str$(TOTHOJA&)): GoTo 29
         '
         If InStr(FOREP$, "/") > 0 Then
           If TEXREPLA$(FOREP$) = FOREP$ Then
              Exit Sub         ' PARA CONMUTAR NO ES AUTOMATICO
           End If
         End If
         If Left$(FOREP$, 1) = "@" Then Exit Sub
         If Left$(FOREP$, 1) = "#" Then
            If UCase$(FOREP$) = "#FECHA" Then
               HOII% = HOY(FOREP$): GoTo 29
            End If
            Exit Sub
         End If
         '
         ' TRABAJA CON TEXTO REEMPLAZADO
         FOREP$ = TEXREPLA$(FOREP$)
         TVAR$ = UCase$(Left$(FORPA$(i%), 1))
         If TVAR$ = "I" Or TVAR$ = "F" Or TVAR$ = "G" Or TVAR$ = "$" Or TVAR = "H" Then
            FOREP$ = FORMATNUM$(RESUCALC#(FOREP$), FORPA$(i%))
         End If
29       Texto(i%) = FOREP$    ' FORMATNUM$(FOREP$, FORPA$(I%))    ' ASIGNA COMO TEXTO EL PARAMETRO
         Exit Sub
      End If
      '
   End If
   '
End Sub
'
Sub CARGATEGRI(TEGRI$, NUPIC%)
   '
   TEGRJ$ = TRIM$(TEGRI$)
   Do While TEGRJ$ <> ""
      On Error Resume Next
      PPXX% = InStr(TEGRJ$, "(")
      If PPXX% < 1 Then Exit Sub
      PPYY% = InStr(PPXX% + 1, TEGRJ$, ")")
      If PPYY% <= PPXX% Then PPYY% = 1 + Len(TEGRJ$)
      TTZZ$ = Mid$(TEGRJ$, PPXX% + 1, PPYY% - PPXX% - 1)
      TEGRJ$ = Mid$(TEGRJ$, PPYY% + 1)
      '
      PPXX% = InStr(TTZZ$, ";")
      ROWI& = XVALO(Left$(TTZZ$, PPXX%))
         If ROWI& > CANREGRID%(NUPIC%) Then CANREGRID%(NUPIC%) = ROWI&
      COLI& = XVALO(Mid$(TTZZ$, PPXX% + 1))
      PPXX% = InStr(TTZZ$, "VALUE =")
      PPYY% = InStr(PPXX% + 1, TTZZ$, ";")
      If PPYY% <= PPXX% Then PPYY% = 1 + Len(TTZZ$)
      valor = Mid$(TTZZ$, PPXX% + 7, PPYY% - PPXX% - 7)
      '
      STDSCREEN.GRID99(NUPIC%).TextMatrix(ROWI& - 1, COLI& - 1) = valor
      On Error GoTo 0
   Loop
   '
End Sub

Function GETPICTURE$()
    '
    If HABILAPLI%("GETPICTURE") < 0 Then
       GETPICTURE$ = ""
       Exit Function
    End If
    '
    ADIR$ = IDENTER$
    If ADIR$ = "" Then GoTo 10
    '
    APLIFO$ = TRIM$(VALCONTROL(0, ADIR$, "APLICFOT"))
    RUTAFO$ = TRIM$(VALCONTROL(0, ADIR$, "ARCHIFOT"))
      If Right$(RUTAFO$, 1) <> "\" Then RUTAFO$ = RUTAFO$ + "\"
      '
    If APLIFO$ <> "" Then
       If EXISTE%(RUTAFO$ + ".") < 1 Then
          MsgBox "No Existe Carpeta '" + RUTAFO$ + "'"
          GETPICTURE$ = ""
          GoTo 99
       End If
       '
       FORMALTER.Label3 = FORMATNUM$(CDbl(Now), "F15.7")
       XYZ = Shell(APLIFO$, vbNormalFocus)
       FORMALTER.Picture1.Top = 0
       FORMALTER.Picture1.Left = 0
       FORMALTER.Width = 500 + FORMALTER.Frame2.Width
       FORMALTER.Frame2.Left = (FORMALTER.ScaleWidth - FORMALTER.Frame2.Width) / 2
       FORMALTER.Frame2.Top = 150
       FORMALTER.Height = 500 + FORMALTER.Frame2.Height
       FORMALTER.Timer1.Enabled = True
       FORMALTER.Picture1.Visible = True
       FORMALTER.Label2 = RUTAFO$
       FORMALTER.Label4 = ""
       FORMALTER.Show 1
       ACAPTU$ = FORMALTER.Label4
       Unload FORMALTER
       If ACAPTU$ <> "" Then
         If EXISTE%(ACAPTU$) > 0 Then
           GETPICTURE$ = ACAPTU$
           GoTo 99
         End If
       End If
    End If
    '
10  STDSCREEN.CommonDialog1.CancelError = True
    STDSCREEN.CommonDialog1.DialogTitle = "Archivo Origen de Imagen"
         ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    STDSCREEN.CommonDialog1.FILTER = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
         "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
         ' Especificar el filtro predeterminado
    STDSCREEN.CommonDialog1.FilterIndex = 2
    On Error GoTo 99
    STDSCREEN.CommonDialog1.ShowOpen
    On Error GoTo 0
    '
    DoEvents
    GETPICTURE$ = UCase$(STDSCREEN.CommonDialog1.FileName)
    '
99 End Function

Function COTIFECHA_HISTODOL#(FE%, MON%)
    CONDI$ = " FECHA = '" & FETEXCOR1$(FE%) & "' AND ID = " & MON%
    COTIFECHA_HISTODOL# = XVALO(VALOBADA("HISTOCAMBIO", "COTIZA_DOLAR", CONDI$, "NOMESS"))
    'CONSIDERO SI ES FIN DE SEMANA QUE TOME LA DEL DIA ANTERIOR HABIL.
    If COTIFECHA_HISTODOL# <= 0 Then
50          If DIASEM(FE%) >= 6 Then
               FE% = DIASANTES(FE%, 1)
               GoTo 50
            End If
    End If
    
    CONDI$ = " FECHA = '" & FETEXCOR1$(FE%) & "' AND ID = " & MON%
    COTIFECHA_HISTODOL# = XVALO(VALOBADA("HISTOCAMBIO", "COTIZA_DOLAR", CONDI$, "NOMESS"))
End Function
Sub BLOQUEAR_CONTROLES(FORMULARIO As Form, HABIL_DESHABIL%, BOTONES%, Optional TXTES%, Optional MSFX%, Optional BLOQUEARCIERRE%, Optional LABELX%, Optional LISTX%)
  '1 HABILITA
  '2 DESHBILITA
  
  Dim CONTROLX As Control
  
  
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is CommandButton And BOTONES% > 0 Then
           If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
           If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
        ElseIf TypeOf CONTROLX Is TextBox And TXTES% > 0 Then
           If HABIL_DESHABIL% = 1 Then CONTROLX.Locked = False
           If HABIL_DESHABIL% = 2 Then CONTROLX.Locked = True
        ElseIf TypeOf CONTROLX Is msFlexGrid And MSFX% > 0 Then
           If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
           If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
        ElseIf TypeOf CONTROLX Is Label And LABELX% > 0 Then
           If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
           If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
        ElseIf TypeOf CONTROLX Is ListBox And LISTX% > 0 Then
           If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
           If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
        End If
  Next
End Sub
Sub BLOQUEAR_CONTROLES_ENCONTENEDOR(FORMULARIO As Form, HABIL_DESHABIL%, BOTONES%, Optional TXTES%, Optional MSFX%, Optional BLOQUEARCIERRE%, Optional LABELX%, Optional LISTX%, Optional OBJETOCONTENEDOR As Object)
  '1 HABILITA
  '2 DESHBILITA
  
  Dim CONTROLX As Control
  
  'HABILITA O DESHABILITA TODOS LOS CONTROLES QUE ESTEN DENTRO DE UN CONTENEDOR PASADO COMO ULTIMO PARAMETRO.
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        A = OBJETOCONTENEDOR.Name
        On Error Resume Next
        B = CONTROLX.Container.Name
        If Err <> 0 Then GoTo 30
        If A = B Then
            
            If TypeOf CONTROLX Is CommandButton And BOTONES% > 0 Then
               If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
               If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
'               If CONTROLX.Name = "Command3" Then
'                  AB = 1
'               End If

            ElseIf TypeOf CONTROLX Is TextBox And TXTES% > 0 Then
               If HABIL_DESHABIL% = 1 Then CONTROLX.Locked = False
               If HABIL_DESHABIL% = 2 Then CONTROLX.Locked = True
            ElseIf TypeOf CONTROLX Is msFlexGrid And MSFX% > 0 Then
               If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
               If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
            ElseIf TypeOf CONTROLX Is Label And LABELX% > 0 Then
               If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
               If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
            ElseIf TypeOf CONTROLX Is ListBox And LISTX% > 0 Then
               If HABIL_DESHABIL% = 1 Then CONTROLX.Enabled = True
               If HABIL_DESHABIL% = 2 Then CONTROLX.Enabled = False
            End If
        End If
30  Next

End Sub

Sub BLOQUEAR_CONTROLES_ENCONTENEDOR_Y_COLOR(FORMULARIO As Form, HABIL_DESHABIL%, Optional TXTES%, Optional OBJETOCONTENEDOR As Object, Optional COLOR$)
  '1 HABILITA
  '2 DESHBILITA
  
  Dim CONTROLX As Control
  
  'HABILITA O DESHABILITA TODOS LOS CONTROLES QUE ESTEN DENTRO DE UN CONTENEDOR PASADO COMO ULTIMO PARAMETRO.
  For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        A = OBJETOCONTENEDOR.Name
        On Error Resume Next
        B = CONTROLX.Container.Name
        If Err <> 0 Then GoTo 30
        If A = B Or TRIM$(A) = "" Then
            If TypeOf CONTROLX Is TextBox And TXTES% > 0 Then
              If TRIM$(COLOR$) = "" Then
                If HABIL_DESHABIL% = 1 Then CONTROLX.Locked = False
                If HABIL_DESHABIL% = 2 Then CONTROLX.Locked = True
              ElseIf COLOR$ = CONTROLX.BackColor Then
                If HABIL_DESHABIL% = 1 Then CONTROLX.Locked = False
                If HABIL_DESHABIL% = 2 Then CONTROLX.Locked = True
              End If
            End If
        End If
30  Next

End Sub


Sub GRABA_CONSULTA_SQL(CODMO$, DESMO$, TIPMO%, TEXMO$, OK%)

        OK% = -1
        Dim rs1 As New ADODB.Recordset
        SQLX$ = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
        rs1.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With rs1
          If .EOF Then
            .AddNew
            !COD_QUERY = CODMO$
            !CODIEMPR = CodiEmp%
          End If
            !TIP_QUERY = TIPMO%
            !DES_QUERY = DESMO$
            !PAR_QUERY = TEXMO$
            .Update
        End With
        rs1.Close
        Set rs1 = Nothing
        OK% = 1
End Sub
Sub CargaObservacionesPredeterminadas()
   Call CREACAMPO("", "TAB_TEXTO", "CODINT_OBS", 3, 0, 1)
   Call CREACAMPO("", "TAB_TEXTO", "TITU_OBS", 10, 64)
   Call CREACAMPO("", "TAB_TEXTO", "TEXTO_OBS", 12, 0)
   Call TRALOCAL("TEXTOBS2", "")
2  OPX% = COMALTER%("Carga de Textos Para Observaciones\\Ingresar Nuevo Texto\Editar Texto Existente\Eliminar")
   
   If OPX% = 1 Then ' NUEVO TEXTO DE OBSERVACION
5     TEXTO_TITULO$ = TRIM$(InputBox("Título", "Ingreso Título de Texto de Observaciones", TITUTEXT$))
      If TEXTO_TITULO$ = "" Then Exit Sub
      If CantRegistros&("TAB_TEXTO", "TITU_OBS  ='" & TEXTO_TITULO$ & "'", "NOMESS") > 0 Then
         Call COMUNI("Título Existente")
         GoTo 5
      End If
      '
      Call SETULTREG("!TEXTOBS2", 0)
8     VTB% = VENTABU%("TEXTOBS2/TITUPAN=" + TEXTO_TITULO$)
      If VTB% >= 0 Then
        EPAC$ = TRIM$(EMPREAC$)
        If InStr(UCase$(EPAC$), "AHP") > 0 Then
            CODFOR$ = TRIM$(Control$("", "FORPRERE"))
            If CODFOR$ = "PRRS-AHP" Then
                FORMATOX$ = TRIM$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO"))
                POS1% = InStr(FORMATOX$, "/")
                SEPASO% = 0
                If POS1% > 0 Then
                   CANTIRECORD% = XVALO(Mid$(FORMATOX$, POS1% + 1))
                End If
                TEXTX$ = ""
                For U& = 1 To ULTREG&("!TEXTOBS2")
                  XX$ = REGLEIDO$("!TEXTOBS2", U&)
                  XX1$ = Mid$(XX$, 3, 96)
                  TEXTX$ = TEXTX$ + XX1$ + vbCrLf
                  If CANTIRECORD% < U& Then SEPASO% = 1
                Next U&
                If SEPASO% > 0 Then
                   OPX% = COMALTER%("El Texto Ingresado Supera el Máximo de Renglones\Configurado Para las Observaciones (" & CANTIRECORD% & ")\\Modificar\Continuar")
                   If OPX% = 1 Then GoTo 8
                End If
            End If
        Else
                For U& = 1 To ULTREG&("!TEXTOBS2")
                  XX$ = REGLEIDO$("!TEXTOBS2", U&)
                  XX1$ = Mid$(XX$, 3, 96)
                  TEXTX$ = TEXTX$ + XX1$ + vbCrLf
                Next U&

        End If
        '
10      RNDX$ = RANDSTRING$(10)
        CONDI$ = "TITU_OBS = '" & RNDX$ & "'"
        If CantRegistros("TAB_TEXTO", "TITU_OBS = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
        Call CreaRegistro("TAB_TEXTO", "TITU_OBS", RNDX$, "NOMESS")
          PROXIMOOBS% = 1 + XVALO(VALOBADA("TAB_TEXTO", "MAX(CODINT_OBS)", "CODINT_OBS > 0", "NOMESS"))
        Call ACTUREGISTRO("TAB_TEXTO", "CODINT_OBS", CONDI$, PROXIMOOBS%, REGAF&, "NOMESS")
        Call ACTUREGISTRO("TAB_TEXTO", "TEXTO_OBS", CONDI$, TEXTX$, REGAF&, "NOMESS")
        Call ACTUREGISTRO("TAB_TEXTO", "TITU_OBS", CONDI$, TEXTO_TITULO$, REGAF&, "NOMESS")
        GoTo 2
      End If
      '
    ElseIf OPX% = 2 Then
      'EDICION DE OBSERVACION EXISTENTE
      CONDI$ = "CODINT_OBS > 0 ORDER BY CODINT_OBS"
      Call CARGALISEL("LISTTITU", "TAB_TEXTO", "CODINT_OBS/F5;TITU_OBS/A48", CONDI$, "NOMESS/DISTINCT")
      Call SELECHO("LISTTITU/TITUTLOS DE TEXTO DE OBSERVACIONES")
      If VALACT1$("LISTTITU") = "" Then GoTo 2 ' COMALTER
      '
      CODITITU% = XVALO(VALACT1$("LISTTITU"))
      CONDI$ = "CODINT_OBS = " & CODITITU%
      textox$ = VALOBADA("TAB_TEXTO", "TEXTO_OBS", CONDI$, "NOMESS")
      'SEPARO EL TEXTO SEGUN LOS GOLPES DE ENTER
      X = TEXTOLOTES$(textox$, Chr(13))
      TXX1$ = Space(96)
      Call SETULTREG("!TEXTOBS2", 0)
      For J% = 1 To UBound(CADENATEX$)
         TROZO$ = CADENATEX$(J%)
         TXX1$ = TROZO$
         Call REPLA(TXX1$, MKI$(J%), 1, 2)
         Call REPLA(TXX1$, TROZO$, 3, 96)
         Call REGAPP("!TEXTOBS2", TXX1$)
      Next J%
      VTB% = VENTABU%("TEXTOBS2/TITUPAN=" + VALACT2$("LISTTITU"))
      TXXT$ = ""
      If VTB% >= 0 Then
        For i& = 1 To ULTREG&("!TEXTOBS2")
          XX$ = REGLEIDO$("!TEXTOBS2", i&)
          TXX1$ = TRIM$(Mid$(XX$, 3, 96))
          TXXT$ = TXXT$ & TXX1$ + vbCrLf
        Next i&
        Call ACTUREGISTRO("TAB_TEXTO", "TEXTO_OBS", CONDI$, TXXT$, REGAF&, "NOMESS")
        GoTo 2
      End If
   ElseIf OPX% = 3 Then
      'ELIMINAR COMPLETAMENTE UN TEXTO DE OBSERVACION EXISTENTE.
      CONDI$ = "CODINT_OBS > 0 ORDER BY CODINT_OBS"
      Call CARGALISEL("LISTTITU", "TAB_TEXTO", "CODINT_OBS/F5;TITU_OBS/A48", CONDI$, "NOMESS/DISTINCT")
      Call SELECHO("LISTTITU/TITULO DE TEXTO DE OBSERVACIONES")
      If VALACT1$("LISTTITU") = "" Then GoTo 2 ' COMALTER
      '
      CODITITU% = XVALO(VALACT1$("LISTTITU"))
      Call BORRAREGISTROS("TAB_TEXTO", "CODINT_OBS = " & CODITITU%, REGAF&, "NOMESS")
      Call COMUNI("Texto de Observación Seleccionado 'Eliminado'")
      GoTo 2
   End If
   '
   Call TRACENTRAL("TEXTOBS2", "")

End Sub

Sub HAB_O_DESCOMMAND(FRM As Form, DESHABILITAR%)
  Dim CONTROLX As Control
  For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is CommandButton Or TypeOf CONTROLX Is comboBox Then
          If DESHABILITAR% > 0 Then
           CONTROLX.Enabled = False
          Else
           CONTROLX.Enabled = True
          End If
        End If
  Next
End Sub
Public Function Conmutar(ByRef valor As Boolean) As Boolean
valor = Not valor
Conmutar = valor
'EJEMPLO DE USO:
'Private Sub Check3_Click()
'    Dim estadoVisible As Boolean
'    estadoVisible = Frame3.Visible
'    ' Llamar a la función Conmutar para cambiar el valor entre true o false segun posea al momento
'     Frame3.Visible = Conmutar(estadoVisible)
'End Sub

End Function

