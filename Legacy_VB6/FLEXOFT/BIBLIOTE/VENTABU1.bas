Attribute VB_Name = "VENTABU1"
Dim RGFJ%(32), TFF%(32)
Dim VFMIN$(32), VFMAX$(32)
'
Dim VDS$(32, 16), VDX#(32, 16), VD%(32, 16), VFS$(32), VFSA$(32), VFX#(32), VF%(32), FI$(32)
Dim TVAL%(20), CHE$(20), BAN$(20), IMPCHE#(20), FECHE%(20)
Dim VDP%(32), VDPX#(32), VDPS$(32), VDA$(32), VDAA%(32), VDAB#(32)
Dim SUMA#(32, 2), SUPAR#(32)
'
Dim ATRIBU$

Function VENTABU%(A$)
    '
    TIFORMED% = 1
    PX% = InStr(A$, "/")
    If PX% > 0 Then ATRIBU$ = Mid$(A$, PX%): A$ = Left$(A$, PX% - 1)
    FORMA$ = TRIM$(A$) + ".PRF"
700 Call CARDATVEN
    Call DIBUVEN
    Call CARLIVEN
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
    If FORMA$ = "PRINSET.PRF" Then
        MsgBox "Seleccionando Fuentes de Impresión Disponibles ..."
    End If
    FORVENTA.Show 1
    Screen.MousePointer = SMP
    Call FRESHALL
    VENTABU% = VTB%
    BSUW& = 0: BIZW = 0: ATRIBU$ = ""
    '
End Function
'

Sub CARDATVEN()
     '

5800 GRABACT% = 0
     BSUW& = 0: BIZW& = 0
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%: N15% = 0
     vuelta% = 0
5801 N15% = FILEOPEN%(FORMA$, 0, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%: N15% = 0
         Kill FORMA$
         If vuelta% = 0 Then
             vuelta% = 1
             Call EXTRACT(FORMA$)
             GoTo 5801
         End If
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Close: End
     End If
     '
     ARCHI$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(ARCHI$)
     ARCHILOC$ = "!" + ARCHITRA$
     PROCE$ = TRIM$(Mid$(YYY$, 65, 8))
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
5803 S1$ = Left$(YYY$, 48)
     SALI$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APPI% = Asc(Mid$(COD$, 6))
        FORVENTA.BOTAGRE.Enabled = True
        If APPI% < 1 Then
            FORVENTA.BOTAGRE.Enabled = False
        End If
     IDL% = Asc(Mid$(COD$, 7))
        FORVENTA.BOTINSE.Enabled = True
        FORVENTA.BOTERASE.Enabled = True
        If IDL% < 1 Then
            FORVENTA.BOTINSE.Enabled = False
            FORVENTA.BOTERASE.Enabled = False
        End If
     FORVENTA.BOTEDIT.Enabled = False
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     If IDL% > 0 Or APPI% > 0 Then GRABACT% = 1
     PRX% = Asc(Mid$(COD$, 10))
     NCL% = 0
     NFT% = 0
     NCLF% = 0
     TCL% = 0
     HAYTOTA% = 0
     RG% = 1
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
         If RG% = 2 Then TVL%(TCL%) = 3: LDL%(TCL%) = 0
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     While PR%(TCL%) < PR%(TCL% - 1)
        PR%(TCL%) = PR%(TCL%) + 256
     Wend
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then
           PR%(TCL%) = 0
           LF%(TCL%) = 2
         End If
5809 TT%(TCL%) = Asc(Mid$(YYY$, 88))
     If TT%(TCL%) > 0 Then
        HAYTOTA% = 1
        If TT%(TCL%) <> TCL% Then
            RTT%(TT%(TCL%) Mod 100) = 1
        End If
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5816 ' SUPRIME COLUMNAS LONGITUD 0
     '
5812 NCL% = NCL% + 1
     CLT%(TCL%) = NCL%
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = REPLACEN$(TRIM$(Mid$(YYY$, 37, 32)))
     If Left$(ECL$(NCL%), 1) = "@" Then
          NCLF% = NCLF% + 1
     End If
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     PE%(NCL%) = Asc(Mid$(YYY$, 73))
          If RG% = 2 Then PE%(NCL%) = 2
     If PE%(NCL%) = 1 Then
        GRABACT% = 1
        FORVENTA.BOTEDIT.Enabled = True
     End If
5815 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
5816 GoTo 5805
5817 '
5820 ATOT% = NCL% - 1 - NCLF%
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
        LC%(KCL%) = ACL%(KCL%)
        If Left$(ECL$(KCL%), 1) <> "@" Then
            LC%(KCL%) = ACL%(KCL%)
            ATOT% = ATOT% + ACL%(KCL%)
            CL%(KCL%) = CLII%
            CLII% = CLII% + ACL%(KCL%) + 1
        End If
     Next KCL%
     '
     If ATOT% > 92 Then
        Call MENSERR(24, "Error: Ancho de Pantalla Tabulada " + FORMA$ + " =" + Str$(ATOT% + 2) + " Pos.")
        Close: End   ' Call FINAL("")
     End If
     '
5821 For KCL% = 1 To NCL%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NCL%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3
               GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
     '
5825 If BLP% > 0 Then
         If OFS% = 0 Then
             OFS% = (78 - ATOT%) / 2
             For KCL% = 1 To NCL%
                 CL%(KCL%) = CL%(KCL%) + OFS%
             Next KCL%
         End If
     End If
     '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
     RGFI% = 25: RGFS% = 1
     '
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 83)) < 1 Then GoTo 5847 ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     RGFL&(NFT%) = Asc(Mid$(YYY$, 84))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
5843 If RGFL&(NFT%) - 1 < RG0% Then
         RG0% = RGFL&(NFT%) - 1
     End If
     If RGFL&(NFT%) < RGFI% Then RGFI% = RGFL&(NFT%)
     If RGFL&(NFT%) > RGFS% Then RGFS% = RGFL&(NFT%)
     '
5844 ECLF$(NFT%) = RTrim$(Mid$(YYY$, 37, 32))
     If Len(ECLF$(NFT%)) > 0 Then
         ECLF$(NFT%) = ECLF$(NFT%) + ":"
     End If
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70))
     LEF%(NFT%) = Asc(Mid$(YYY$, 86))
     CEF%(NFT%) = Asc(Mid$(YYY$, 87))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8)
5847 GoTo 5835
 '
5850 Close #N15%: N15% = 0
     '
5700 PX% = InStr(UCase$(ATRIBU$), "TITUPAN=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 8)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If TRIM$(XX$) <> "" Then
             S1$ = XX$
             FORVENTA.Caption = S1$
         End If
     End If
     '
5705 PX% = InStr(UCase$(ATRIBU$), "ARCHIVO=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 8)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If TRIM$(XX$) <> "" Then ARCHI$ = XX$
     End If
     '
5710 PX% = InStr(UCase$(ATRIBU$), "BORDESUP=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Left$(XX$, 1) = "W" Then
              BSUW& = Val(Mid$(XX$, 2))
            Else
              RGI1& = Val(XX$)
              If RGI1& > 0 Then If RGI1& <= 22 Then RGI& = RGI1&
         End If
     End If
     '
5720 PX% = InStr(UCase$(ATRIBU$), "BORDEINF=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         RGF1& = Val(XX$)
         If RGF1& > RGI& Then If RGF1& <= 22 Then RGF& = RGF1&
     End If
     If (RGF& - RGI&) < 12 Then RGF& = RGI& + 12
     '
5730 PX% = InStr(UCase$(ATRIBU$), "BORDEIZQ=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9)
         PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
         XX$ = Left$(XX$, PY% - 1)
         If Left$(XX$, 1) = "W" Then
               BIZW& = Val(Mid$(XX$, 2))
             Else
               OFS1% = Val(XX$)
               If OFS1% > 0 Then If OFS1% <= 76 Then OFS% = OFS1%
         End If
         '
         CLII% = OFS% + 2
         For KCL% = 1 To NCL%
            LC%(KCL%) = ACL%(KCL%)
            If Left$(ECL$(KCL%), 1) <> "@" Then
               LC%(KCL%) = ACL%(KCL%)
               ATOT% = ATOT% + ACL%(KCL%)
               CL%(KCL%) = CLII%
               CLII% = CLII% + ACL%(KCL%) + 1
            End If
         Next KCL%
         '
     End If
     '
5740 PX% = InStr(UCase$(ATRIBU$), "COLORPAN=")
     If PX% > 0 Then
         XX$ = Mid$(ATRIBU$, PX% + 9, 3)
         If Mid$(XX$, 2, 1) = "-" Then COLPAN$ = XX$
     End If
     '
5750 NOCIE% = 0
     If InStr(UCase$(ATRIBU$), "/NC") > 0 Then NOCIE% = 1
     '
5760 PREGU% = 0
     If InStr(UCase$(ATRIBU$), "/PS") > 0 Then PREGU% = 1
     '
5770 PX% = InStr(UCase$(ATRIBU$), "MODO=")
     MODOS$ = ""
     If PX% > 0 Then
         MODOS$ = Mid$(ATRIBU$, PX% + 5, 3)
     End If
     '
5780 NOFORM% = 0
     If InStr(UCase$(ATRIBU$), "/NOFORM") > 0 Then NOFORM% = 1
     '
5790 For UU% = 1 To NCL%
        ARGATRI$ = "ARCHECO(" + TRIM$(Str$(UU%)) + ")="
        PX% = InStr(UCase$(ATRIBU$), ARGATRI$)
        If PX% > 0 Then
           XX$ = Mid$(ATRIBU$, PX% + Len(ARGATRI$))
           PY% = InStr(XX$, "/"): If PY% < 1 Then PY% = 1 + Len(XX$)
           XX$ = Left$(XX$, PY% - 1)
           If TRIM$(XX$) <> "" Then AE$(UU%) = XX$
        End If
     Next UU%
     '
End Sub
'

Sub DIBUVEN()     ' Rutina de dibujo pantalla tabulada
'
HREN% = FORVENTA.ScaleHeight / 25
FORVENTA.Caption = S1$
MAIZQ = 0
MAINT = 100
MADER = 200
'
ENCA$ = " "
'
For KCL% = 1 To NCL%
      '
      If Left$(ECL$(KCL%), 1) <> "@" Then
          If Left$(ECL$(KCL%), 1) = "<" Then
              ENCACO$ = AJUSTI$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Space$(MG%(KCL%)) + ENCACO$
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
            ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ENCACO$ = AJUSTD$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Mid$(ENCACO$, MG%(KCL%) + 1) + Space$(MG%(KCL%))
            Else
              ENCACO$ = TRIM$(ECL$(KCL%))
              While Len(ENCACO$) < LC%(KCL%)
                 ENCACO$ = " " + ENCACO$ + " "
              Wend
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
          End If
          ENCA$ = ENCA$ + ENCACO$
          ENCA$ = ENCA$ + " "
        Else
          ' CL%(KCL%) = 4 + Len(ECL$(KCL%))
      End If
      '
Next KCL%
'
ENCA$ = ENCA$ + "   "
FORVENTA.Top = Screen.Height / 25 * (RGI&)
FORVENTA.Height = Screen.Height / 25 * (RGF& - RGI& - 1)
FORVENTA.Left = Screen.Width / 80 * (OFS% - 1)
'
If BSUW& <> 0 Then FORVENTA.Top = BSUW&
If BIZW& <> 0 Then FORVENTA.Left = BIZW&
'
FORVENTA.TEPRUEBA.Caption = ENCA$
FORVENTA.LITAB.Width = FORVENTA.TEPRUEBA.Width
FORVENTA.FRAMENCA.Width = FORVENTA.TEPRUEBA.Width
FORVENTA.ENCACOL.Caption = ENCA$
FORVENTA.ENCACOL.Width = FORVENTA.TEPRUEBA.Width
FORVENTA.FRAMETOT.Width = FORVENTA.TEPRUEBA.Width
FORVENTA.TOTALI.Width = FORVENTA.TEPRUEBA.Width
FORVENTA.LITAB.Left = MAIZQ
FORVENTA.FRAMENCA.Left = MAIZQ
FORVENTA.ENCACOL.Left = 50
FORVENTA.FRAMETOT.Left = MAIZQ
FORVENTA.TOTALI.Left = 50
FORVENTA.Width = MAIZQ + FORVENTA.LITAB.Width + MAINT + FORVENTA.BOTEXIT.Width + MADER
COEFH = FORVENTA.LITAB.Width / Len(ENCA$)
'
FORVENTA.FRAMENCA.Top = 0
FORVENTA.FRAMENCA.Height = 1.6 * FORVENTA.ENCACOL.Height
FORVENTA.ENCACOL.Top = (FORVENTA.FRAMENCA.Height - FORVENTA.ENCACOL.Height) / 2
FORVENTA.ENCACOL.Caption = ENCA$
FORVENTA.LITAB.Top = FORVENTA.FRAMENCA.Top + FORVENTA.FRAMENCA.Height
'
FORVENTA.FRAMETOT.Height = 1.6 * FORVENTA.TOTALI.Height
FORVENTA.LITAB.Height = FORVENTA.ScaleHeight - FORVENTA.LITAB.Top - FORVENTA.FRAMETOT.Height
FORVENTA.FRAMETOT.Top = FORVENTA.LITAB.Top + FORVENTA.LITAB.Height
FORVENTA.TOTALI.Top = (FORVENTA.FRAMETOT.Height - FORVENTA.TOTALI.Height) / 2
FORVENTA.FRAMETOT.Visible = True
ALTUBA = FORVENTA.Height
FORVENTA.Height = FORVENTA.FRAMETOT.Top + FORVENTA.FRAMETOT.Height
FORVENTA.Top = FORVENTA.Top + (ALTUBA - FORVENTA.Height) / 2
'
FORVENTA.FRAMEBOT1.Top = FORVENTA.LITAB.Top
FORVENTA.FRAMEBOT1.Left = MAIZQ + FORVENTA.LITAB.Width + MAINT
FORVENTA.FRAMEBOT2.Top = FORVENTA.FRAMETOT.Top - 1560
FORVENTA.FRAMEBOT2.Left = MAIZQ + FORVENTA.LITAB.Width + MAINT
'
FORVENTA.LITAB.Clear
While FORVENTA.LITAB.TopIndex <= 0
   FORVENTA.LITAB.AddItem "   "
   'Call VERSCROLLVEN
   FORVENTA.LITAB.ListIndex = FORVENTA.LITAB.ListCount - 1
Wend
CARELI% = FORVENTA.LITAB.ListCount - 1
FORVENTA.LITAB.Clear
'Call VERSCROLLVEN
'
FORVENTA.VScroll1.Top = FORVENTA.LITAB.Top
FORVENTA.VScroll1.Height = FORVENTA.LITAB.Height
FORVENTA.VScroll1.Left = FORVENTA.LITAB.Left + FORVENTA.LITAB.Width - FORVENTA.VScroll1.Width - 20
FORVENTA.VScroll1.Visible = True
'
'FORVENTA.Refresh
'
End Sub
'

Sub CARLIVEN()
    '
    Screen.MousePointer = 11
    TINI = Timer
    FIN& = ULTREG&(ARCHILOC$)
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        '
701     XXX$ = REGLEIDO$(ARCHILOC$, U&)
        RELIB$ = Space$(ATOT% + 2)
        '
        For K2% = 1 To NCL%
           If CFT%(K2%) <> 1 Then
             If RTT%(CFT%(K2%)) < 1 Then
               If LF%(CFT%(K2%)) > 0 Then
                   XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
                   ACLX% = ACL%(K2%) - MG%(K2%)
                   TXT$ = CSTRING$(XXZ$, TVL%(CFT%(K2%)), ACLX%, LDL%(CFT%(K2%)), JT%(K2%))
                   '
                   If JT%(K2%) = 2 Then
                       TXT$ = TXT$ + Space$(MG%(K2%))
                     Else
                       TXT$ = Space$(MG%(K2%)) + TXT$
                   End If
                   If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%) - OFS%, LC%(K2%)) = TXT$
                   '
                   If TRIM$(AE$(K2%)) <> "" Then
                      XXE$ = ECOARCH$(AE$(K2%), XXZ$)
                      COLOG% = K2%
                      CLE% = CL%(CE%(COLOG%))
                      LCE% = LC%(CE%(COLOG%))
                      JTE% = JT%(CE%(COLOG%))
                      TXT$ = CSTRING$(XXE$, 6, LCE%, 0, JTE%)
                      If CLE% > 0 Then Mid$(RELIB$, CLE% - OFS%, LCE%) = TXT$
                   End If
                   '
                   TTT% = TT%(CFT%(K2%))
                   If TTT% > 0 Then
                      If (TTT% Mod 100) <= TCL% Then
                         TTX#(K2%) = TTX#(K2%) + Val(TXT$)
                         If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + Val(TXT$) Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - Val(TXT$)
                         If (TTT% Mod 100) <> CFT%(K2%) Then
                             COLOG% = CLT%(TTT% Mod 100)   ' CFT%(TTT% Mod 100)
                             CLE% = CL%(COLOG%)
                             LCE% = LC%(COLOG%)
                             JTE% = JT%(COLOG%)
                             TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                             If CLE% > 0 Then Mid$(RELIB$, CLE% - OFS%, LCE%) = TXT$
                         End If
                      End If
                   End If
                   '
                End If
             End If
           End If
        Next K2%
        '
        FORVENTA.LITAB.AddItem RELIB$ + Str$(U&)
        '
989 Next U&
   '
   Screen.MousePointer = 1
   If FORVENTA.LITAB.ListCount < 1 Then
        If APPI% = 1 Then
          MODEDI% = 1
        End If
      Else
        FORVENTA.LITAB.ListIndex = 0
        If FORVENTA.LITAB.ListCount > CARELI% Then
            FORVENTA.VScroll1.Visible = False
        End If
        ' Call VERSCROLLVEN
   End If
   '
End Sub
'



Sub ADDLISVEN()
    '
    FORVENTA.LITAB.AddItem Space$(ATOT%)
    Call VERSCROLLVEN
    FORVENTA.LITAB.ListIndex = FORVENTA.LITAB.ListCount - 1
    If FORVENTA.LITAB.TopIndex < FORVENTA.LITAB.ListIndex - CARELI% Then
        FORVENTA.LITAB.TopIndex = FORVENTA.LITAB.ListCount - CARELI% - 3
    End If
    MODEDI% = 1
    NUREGIS& = FORVENTA.LITAB.ListCount + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRAVEN
    'FORVENTA.Refresh
    '
End Sub
'

Sub INSLISVEN()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = FORVENTA.LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > FORVENTA.LITAB.ListCount - 1 Then
        ILISTA& = 0
    End If
    '
    FORVENTA.LITAB.AddItem Space$(ATOT%), ILISTA&
    Call VERSCROLLVEN
    FORVENTA.LITAB.ListIndex = ILISTA&
    If FORVENTA.LITAB.TopIndex < FORVENTA.LITAB.ListIndex - CARELI% Then
        FORVENTA.LITAB.TopIndex = FORVENTA.LITAB.ListIndex - CARELI% / 2
    End If
    '
    For U& = ULTREG&(ARCHILOC$) To ILISTA& + 1 Step -1
       REACTIVO$ = REGLEIDO$(ARCHILOC$, U&)
       Call GRAREG(ARCHILOC$, REACTIVO$, U& + 1)
    Next U&
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call GRAREG(ARCHILOC$, REACTIVO$, ILISTA& + 1)
    '
    MODEDI% = 2
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRAVEN
    '
End Sub
'
Sub BORLISVEN()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = FORVENTA.LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > FORVENTA.LITAB.ListCount - 1 Then
        Exit Sub
    End If
    '
    FORVENTA.LITAB.RemoveItem ILISTA&
    Call VERSCROLLVEN
    For U& = ILISTA& + 2 To ULTREG&(ARCHILOC$)
       REACTIVO$ = REGLEIDO$(ARCHILOC$, U&)
       Call GRAREG(ARCHILOC$, REACTIVO$, U& - 1)
    Next U&
    REACTIVO$ = String$(RECLEN%(ARCHILOC$), 0)
    Call GRAREG(ARCHILOC$, REACTIVO$, ULTREG&(ARCHILOC$))
    Call CIERRARCH(ARCHILOC$)
    '
    If ILISTA& > FORVENTA.LITAB.ListCount - 1 Then
       ILISTA& = FORVENTA.LITAB.ListCount - 1
    End If
    FORVENTA.LITAB.ListIndex = ILISTA&
    '
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
End Sub
'

