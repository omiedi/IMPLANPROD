Attribute VB_Name = "IMPORTAR"
Sub IMPORCLI()
   N1% = FILEOPEN%("F:\CLIENTES\ELROBLE\TRAIDOS\MAYCLI1.TXT", 1, 0)
   REGX& = 0: REGY& = 0
   ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
   ICLI$ = ""
   '
   'Line Input #N1%, ATT$
   '
   While Not EOF(N1%)
      Line Input #N1%, ATT$
      NC% = Val(Left$(ATT$, 8))
      For UI% = 1 To Len(ICLI$) Step 2
        If CVI(Mid$(ICLI$, UI%, 2)) = NC% Then GoTo 19
      Next UI%
      ICLI$ = ICLI$ + MKI$(NC%)
      '
      CLI$ = Mid$(ATT$, 23, 30)
      FANT$ = Mid$(ATT$, 23, 10)
      DOMI$ = Mid$(ATT$, 53, 30)
      CPOST% = Val(Mid$(ATT$, 129, 10))
      LOCA$ = Mid$(ATT$, 83, 30)
      TELE$ = Mid$(ATT$, 113, 16)
      CUIT$ = Mid$(ATT$, 260, 13)
      '
      PXIVA$ = UCase$(Mid$(ATT$, 171, 1))
      PIVA% = 1
        If PXIVA$ = "2" Then PIVA% = 0
        If PXIVA$ = "4" Then PIVA% = 6
      CATIB% = 1
      LIPRE% = 1
      COPAGI% = 1
      VENDE% = Val(Mid$(ATT$, 193, 1))
      CUECO% = 0
      PROV$ = ""
      If Val(Mid$(ATT$, 151, 2)) >= 10 Then
         PROV$ = Chr$(55 + Val(Mid$(ATT$, 151, 2)))
         DEPROVIX$ = TRIM$(Mid$(ATT$, 153, 17))
         If InStr(IPROVI$, PROV$) < 1 Then
           IPROVI$ = IPROVI$ + PROV$
           RBB$ = REGBLAN$("PROVINC")
           Call REPLA(RBB$, PROV$, 1, 1)
           Call REPLA(RBB$, DEPROVIX$, 2, 20)
           Call REGAPP("PROVINC", RBB$)
           Call CIERRARCH("PROVINC")
         End If
      End If
      '
      DESTA$ = TRIM$(Str$(Int(Val(Mid$(ATT$, 194, 8)))))
      If Val(DESTA$) < 1 Then DESTA$ = ""
      DESTA2$ = TRIM$(Str$(Int(Val(Mid$(ATT$, 215, 8)))))
      If Val(DESTA2$) < 1 Then DESTA2$ = ""
   If DESTA2$ <> "" Then
   AAA = BBB
   End If
      If DESTA$ <> "" And DESTA2$ <> "" Then DESTA$ = DESTA$ + "+" + DESTA2$
      '
      If DESTA$ <> "" Then
        RBB$ = REGBLAN$("DESCUCLI")
        Call REPLA(RBB$, MKI$(NC%), 1, 2)
        Call REPLA(RBB$, DESTA$, 3, 14)
        Call REGAPP("DESCUCLI", RBB$)
      End If
      '
      XX$ = REGBLAN$(ARX1$) + REGBLAN$(ARX2$)
      Call REPLA(XX$, MKI$(NC%), 1, 2)
      Call REPLA(XX$, TRIM$(CLI$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, Chr$(PIVA%), 136, 1)
      Call REPLA(XX$, PROV$, 149, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COPAGI%), 150, 1)
      Call REPLA(XX$, Chr$(VENDE%), 152, 1)
      Call REPLA(XX$, MKI$(CUECO%), 157, 2)
      
      REGX& = REGX& + 1
      Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
      Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
      '
      XX$ = REGBLAN$("BADACLIE")
      Call REPLA(XX$, MKI$(NC%), 1, 2)
      Call REPLA(XX$, TRIM$(CLI$), 3, 62)
      Call REPLA(XX$, FANT$, 65, 32)
      Call REPLA(XX$, TRIM$(DOMI$), 97, 64)
      Call REPLA(XX$, TRIM$(Str$(CPOST%)), 161, 16)
      Call REPLA(XX$, TRIM$(LOCA$), 177, 48)
      Call REPLA(XX$, TRIM$(TELE$), 257, 24)
      
      REGY& = REGY& + 1
      Call GRAREG("BADACLIE", XX$, REGY&)
      
19 Wend
   Call SETULTREG(ARX1$, REGX&)
   Call CIERRARCH(ARX1$)
   Call SETULTREG(ARX2$, REGX&)
   Call CIERRARCH(ARX2$)
   Call SETULTREG("BADACLIE", REGY&)
   Call CIERRARCH("BADACLIE")
   '
End Sub
'
Sub IMPORPRO()
   N1% = FILEOPEN%("F:\CLIENTES\ELROBLE\TRAIDOS\MAYPRO1.TXT", 1, 0)
   REGX& = 0: REGY& = 0
   ARX1$ = "PROVED1": ARX2$ = "PROVED2"
   ICLI$ = ""
   '
   'Line Input #N1%, ATT$
   '
   While Not EOF(N1%)
      Line Input #N1%, ATT$
      NC% = Val(Left$(ATT$, 8))
      For UI% = 1 To Len(ICLI$) Step 2
        If CVI(Mid$(ICLI$, UI%, 2)) = NC% Then GoTo 19
      Next UI%
      ICLI$ = ICLI$ + MKI$(NC%)
      '
      CLI$ = Mid$(ATT$, 23, 30)
      FANT$ = Mid$(ATT$, 23, 10)
      DOMI$ = Mid$(ATT$, 53, 30)
      CPOST% = Val(Mid$(ATT$, 129, 10))
      LOCA$ = Mid$(ATT$, 83, 30)
      TELE$ = Mid$(ATT$, 113, 16)
      CUIT$ = Mid$(ATT$, 260, 13)
      '
      'PXIVA$ = UCase$(Mid$(ATT$, 171, 1))
      PIVA% = 1
       ' If PXIVA$ = "2" Then PIVA% = 0
       ' If PXIVA$ = "4" Then PIVA% = 6
      CATIB% = 1
      LIPRE% = 0
      COPAGI% = 1
      VENDE% = 0 'Val(Mid$(ATT$, 193, 1))
      CUECO% = 0
      PROV$ = ""
      If Val(Mid$(ATT$, 151, 2)) >= 10 Then
         PROV$ = Chr$(55 + Val(Mid$(ATT$, 151, 2)))
         'DEPROVIX$ = TRIM$(Mid$(ATT$, 153, 17))
         'If InStr(IPROVI$, PROV$) < 1 Then
         '  IPROVI$ = IPROVI$ + PROV$
         '  RBB$ = REGBLAN$("PROVINC")
         '  Call REPLA(RBB$, PROV$, 1, 1)
         '  Call REPLA(RBB$, DEPROVIX$, 2, 20)
         '  Call REGAPP("PROVINC", RBB$)
         '  Call CIERRARCH("PROVINC")
         'End If
      End If
      '
      'DESTA$ = TRIM$(Str$(Int(Val(Mid$(ATT$, 194, 8)))))
      'If Val(DESTA$) < 1 Then DESTA$ = ""
      'DESTA2$ = TRIM$(Str$(Int(Val(Mid$(ATT$, 215, 8)))))
      'If Val(DESTA2$) < 1 Then DESTA2$ = ""
   'If DESTA2$ <> "" Then
   'AAA = BBB
   'End If
      'If DESTA$ <> "" And DESTA2$ <> "" Then DESTA$ = DESTA$ + "+" + DESTA2$
      '
      'If DESTA$ <> "" Then
      '  RBB$ = REGBLAN$("DESCUCLI")
      '  Call REPLA(RBB$, MKI$(NC%), 1, 2)
      '  Call REPLA(RBB$, DESTA$, 3, 14)
      '  Call REGAPP("DESCUCLI", RBB$)
      'End If
      '
      XX$ = REGBLAN$(ARX1$) + REGBLAN$(ARX2$)
      Call REPLA(XX$, MKI$(NC%), 1, 2)
      Call REPLA(XX$, TRIM$(CLI$), 3, 30)
      Call REPLA(XX$, TRIM$(DOMI$), 33, 25)
      Call REPLA(XX$, MKI$(CPOST%), 60, 2)
      Call REPLA(XX$, TRIM$(LOCA$), 62, 25)
      Call REPLA(XX$, TRIM$(TELE$), 87, 10)
      Call REPLA(XX$, TRIM$(FANT$), 97, 10)
      Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
      Call REPLA(XX$, Chr$(PIVA%), 152, 1)
      Call REPLA(XX$, PROV$, 149, 1)
      Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
      Call REPLA(XX$, Chr$(COPAGI%), 150, 1)
      Call REPLA(XX$, MKI$(CUECO%), 157, 2)
      
      REGX& = REGX& + 1
      Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
      Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
      '
      XX$ = REGBLAN$("BADAPROV")
      Call REPLA(XX$, MKI$(NC%), 1, 2)
      Call REPLA(XX$, TRIM$(CLI$), 3, 62)
      Call REPLA(XX$, FANT$, 65, 32)
      Call REPLA(XX$, TRIM$(DOMI$), 97, 64)
      Call REPLA(XX$, TRIM$(Str$(CPOST%)), 161, 16)
      Call REPLA(XX$, TRIM$(LOCA$), 177, 48)
      Call REPLA(XX$, TRIM$(TELE$), 257, 24)
      
      REGY& = REGY& + 1
      Call GRAREG("BADAPROV", XX$, REGY&)
      
19 Wend
   Call SETULTREG(ARX1$, REGX&)
   Call CIERRARCH(ARX1$)
   Call SETULTREG(ARX2$, REGX&)
   Call CIERRARCH(ARX2$)
   Call SETULTREG("BADAPROV", REGY&)
   Call CIERRARCH("BADAPROV")
   '
End Sub
'
Sub IMPORFORMUL()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   If COMALTER%("Forma de Importación:\\Completa\Solo Fórmulas") = 2 Then GoTo 51
   '
   Label58.Visible = True
   IKK& = 0
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 15
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
10   If EOF(N1%) Then
       Close #N1%
       GoTo 15
     End If
     '
     Line Input #N1%, AA$
     IKK& = IKK& + 1: Label58.Caption = TRIM$(Str$(IKK&))
     Label58.Refresh
     COD$ = TRIM$(Mid$(AA$, 1, 15))
     DEX$ = TRIM$(Mid$(AA$, 21, 55))
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
11   COD$ = TRIM$(Mid$(AA$, 76, 15))
     DEX$ = TRIM$(Mid$(AA$, 92, 51))
     UNI$ = "U."
     COSUNIX$ = TRIM$(Mid$(AA$, 149, 10))
     Call REPLACAR(COSUNIX$, ",", ".")
     COU = Val(COSUNIX$)
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     If COD$ <> "" Then
       '
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 12
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = REGBLAN$("MASTER")
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
       Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
       Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
     End If
     '
12   REGMA& = KKI% + 1
     REMAS$ = REGLEIDO$("MASTER", REGMA&)
     If CVS(Mid$(REMAS$, 111, 4)) < 0.0001 Then
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
     End If
     '
     GoTo 10
     '
15 Next UI%
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   For UI% = 0 To 24
     ARIMPO$ = "A:\" + TRIM$(Str$(UI%)) + ".PRN"
     If EXISTE%(ARIMPO$) < 1 Then GoTo 65
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 15))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR$ = TRIM$(Mid$(AA$, 76, 15))
     If TRIM$(CODPAR$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 143, 6))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 Next UI%
   '
   PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub
'
Sub IMPORMAST()
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   LU$ = LUDAT$
   '
   ARIMPO$ = LUDAT$ + "CAUTOTO.TXT"
   ARIMP2$ = LUDAT$ + "CAUTO1.TXT"
   Dim BUFFLE As String * 128
   '
   N1% = FILEOPEN%(ARIMPO$, 1, 128)
   N2% = FILEOPEN%(ARIMP2$, 0, 128)
   CACOD% = 0
   JJ& = 0
   '
GoTo 55
   '
10 If EOF(N1%) Then GoTo 45
   Line Input #N1%, A$
   B$ = TRIM$(A$)
   If B$ = "" Then GoTo 10
   If Left$(A$, 10) = "----------" Then GoTo 10
   If Left$(UCase$(B$), 6) = "CODIGO" Then GoTo 10
   If Left$(UCase$(A$), 8) = "PRODUCCI" Then GoTo 10
   If Left$(UCase$(A$), 8) = "FECHA DE" Then GoTo 10
   If Left$(UCase$(B$), 8) = "COMPOSIC" Then GoTo 10
   If Left$(UCase$(B$), 6) = "DESDE " Then GoTo 10
   If Left$(UCase$(B$), 8) = "- ARMADO" Then GoTo 10
   If Left$(UCase$(B$), 5) = "NIVEL" Then GoTo 10
   If Left$(UCase$(B$), 4) = "ALA " Then GoTo 10
   If Left$(UCase$(B$), 4) = "ARM " Then GoTo 10
   If Left$(UCase$(B$), 4) = "COM " Then GoTo 10
   If Left$(UCase$(B$), 4) = "INY " Then GoTo 10
   If Left$(UCase$(B$), 4) = "TOR " Then GoTo 10
   
   If Left$(UCase$(A$), 8) = "ARTICULO" Then A$ = TRIM$(Mid$(A$, 11))
   
   BUFLE$ = A$
   JJ& = JJ& + 1
   Put #N2%, JJ&, BUFLE$
   GoTo 10
   '
45 Close #N1%
   Close #N2%
   '
50 CACOD% = 0
   LU$ = LUDAT$
   '
   AGREMAS% = 0
   For U& = 2 To ULTREG&("MASTER")
     X$ = REGLEIDO$("MASTER", U&)
     CACOD% = CACOD% + 1
     CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   Next U&
   '
   N2% = FILEOPEN%(ARIMP2$, 0, 128)
   '
   FECO% = HOY(HOS$)
52 For KK& = 1 To JJ&
     Get #N2%, KK&, BUFLE$
     AA$ = BUFLE$
     BB$ = TRIM$(AA$)
     COD$ = TRIM$(Left$(BB$, 15))
     DEX$ = TRIM$(Mid$(BB$, 17, 34))
     CAN = Val(Mid$(BB$, 51, 6))
       If CAN = 0 Then CAN = 1
     UNID$ = TRIM$(Mid$(BB$, 59, 2))
       If UNID$ = "" Then UNID$ = "U."
     '
     If DEX$ <> "" Then
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 54
       Next KKI%
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = String$(128, 0)
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, Chr$(1), 95, 1)
       Call REPLA(REMAS$, UNID$, 109, 2)
       Call REPLA(REMAS$, MKS$(0), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
      End If
54 Next KK&
   Close #N2%
   '
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
55 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
      YACAR1%(KKI%) = 0
   Next KKI%
   '
   N2% = FILEOPEN%(ARIMP2$, 0, 128)
   If JJ& < 1 Then JJ& = LOF(N2%) / 128
   NIVBAS% = 0
   '
57 For KK& = 1 To JJ&
     Get #N2%, KK&, BUFLE$
     AA$ = BUFLE$
     If NIVELE%(AA$) = NIVBAS% Then
       BB$ = TRIM$(AA$)
       CODCON$ = TRIM$(Left$(BB$, 15))
         CI1% = CODINT%(CODCON$)
         If YACAR1%(CI1%) > 0 Then GoTo 65
       CANBAS = Val(Mid$(BB$, 51, 6))
         If CANBAS = 0 Then CANBAS = 1
         '
       For KK1& = KK& + 1 To JJ&
         Get #N2%, KK1&, BUFLE$
         AA$ = BUFLE$
         NVXX% = NIVELE%(AA$)
         If NVXX% <= NIVBAS% Then GoTo 65
         If NVXX% = NIVBAS% + 1 Then
           BB$ = TRIM$(AA$)
           CODPAR$ = TRIM$(Left$(BB$, 15))
           USO = Val(Mid$(BB$, 51, 6)) / CANBAS
           '
           CI1% = CODINT%(CODCON$)
           CI2% = CODINT%(CODPAR$)
           If CI1% > 0 Then
             If CI2% > 0 Then
               If USO > 0 Then
                 CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
                 YACAR1%(CI1%) = 1
               End If
             End If
           End If
         End If
       Next KK1&
     End If
     '
65 Next KK&
   '
   If NIVBAS% < 2 Then
     NIVBAS% = NIVBAS% + 1
     GoTo 57
   End If
   '
   PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")

End Sub

Sub ICUECORR()
   N1% = FILEOPEN%("F:\CLIENTES\ELROBLE\TRAIDOS\FICLI.TXT", 1, 0)
   REGX& = 0: REGY& = 0
   ICLI$ = ""
   '
   'Line Input #N1%, ATT$
   '
   While Not EOF(N1%)
      Line Input #N1%, ATT$
      TCX$ = Mid$(ATT$, 2, 1)
      TIX% = 2 ' FACTURA
      If TCX$ = "3" Then TIX% = 3 ' NOTA DE DEBITO
      If TCX$ = "2" Then TIX% = 6 ' NOTA DE CREDITO
      If TCX$ = "4" Then TIX% = 4 ' COBRANZA
      '
      NC% = Val(Mid$(ATT$, 3, 8))
      '
      FEMOX$ = Mid$(ATT$, 26, 10) + " "
      PPXX% = InStr(FEMOX$, " ")
      FEMOX$ = Left$(FEMOX$, PPXX% - 1)
      PPXX% = InStr(FEMOX$, "/19")
      If PPXX% > 0 Then
        FEMOX$ = Left$(FEMOX$, PPXX% - 1) + "/" + Mid$(FEMOX$, PPXX% + 3)
      End If
      PPXX% = InStr(FEMOX$, "/20")
      If PPXX% > 0 Then
        FEMOX$ = Left$(FEMOX$, PPXX% - 1) + "/" + Mid$(FEMOX$, PPXX% + 3)
      End If
      FEXI% = FECHANUM(FEMOX$)
      '
      VAX% = 1
      If Mid$(ATT$, 45, 1) = "B" Then VAX% = 2
      '
      NRO& = Val(Mid$(ATT$, 55, 4))
      SUCUI% = Val(Mid$(ATT$, 46, 4))
      '
      IMPOX$ = TRIM$(Mid$(ATT$, 59, 12))
      IMPOY$ = REPLACAR$(IMPOX$, ",", ".")
      IMPO# = Val(IMPOY$)
      INET# = (Int(100 * IMPO# / 1.21 + 0.5)) / 100
      IIVA# = IMPO# - INET#
      If TIX% <> 4 Then
        RBB$ = REGBLAN$("FACTUR")
        Call REPLA(RBB$, MKI$(NC%), 1, 2)
        Call REPLA(RBB$, MKS$(NRO&), 3, 4)
        Call REPLA(RBB$, MKI$(FEXI%), 7, 2)
        Call REPLA(RBB$, MKI$(FEXI%), 9, 2)
        Call REPLA(RBB$, Chr$(1), 11, 1)
        Call REPLA(RBB$, Chr$(1), 12, 1)
        Call REPLA(RBB$, Chr$(TIX%), 13, 1)
        Call REPLA(RBB$, Chr$(VAX%), 14, 1)
        Call REPLA(RBB$, MKD$(INET#), 15, 8)
        Call REPLA(RBB$, MKD$(IIVA#), 23, 8)
        Call REPLA(RBB$, MKD$(IMPO#), 39, 8)
        Call REPLA(RBB$, MKI$(SUCUI%), 47, 2)
        Call REGAPP("FACTUR", RBB$)
If TIX% = 6 Then
AAA = BBB
End If
       Else
         RBB$ = REGBLAN$("COBRAN")
        Call REPLA(RBB$, MKI$(NC%), 1, 2)
        Call REPLA(RBB$, MKS$(NRO&), 3, 4)
        Call REPLA(RBB$, MKI$(FEXI%), 7, 2)
        Call REPLA(RBB$, Chr$(TIX%), 9, 1)
        Call REPLA(RBB$, Chr$(VAX%), 10, 1)
        Call REPLA(RBB$, MKD$(IMPO#), 15, 8)
        Call REPLA(RBB$, MKI$(SUCUI%), 23, 2)
        Call REGAPP("COBRAN", RBB$)
      End If
   
   Wend
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("COBRAN")
   

End Sub

Sub IMPORMOV()
   N1% = FILEOPEN%("F:\CLIENTES\ELROBLE\TRAIDOS\DETFAC.TXT", 1, 0)
   REGX& = 0: REGY& = 0
   ICLI$ = ""
   '
   'Line Input #N1%, ATT$
   '
   While Not EOF(N1%)
      Line Input #N1%, ATT$
      TCX$ = Mid$(ATT$, 27, 1)
      TIX$ = "FC" ' FACTURA
      If TCX$ = "3" Then TIX$ = "ND" ' NOTA DE DEBITO
      If TCX$ = "2" Then TIX$ = "NC" ' NOTA DE CREDITO
      '
      NC% = Val(Mid$(ATT$, 29, 8))
      '
      FEMOX$ = Mid$(ATT$, 87, 10) + " "
      PPXX% = InStr(FEMOX$, " ")
      FEMOX$ = Left$(FEMOX$, PPXX% - 1)
      PPXX% = InStr(FEMOX$, "/19")
      If PPXX% > 0 Then
        FEMOX$ = Left$(FEMOX$, PPXX% - 1) + "/" + Mid$(FEMOX$, PPXX% + 3)
      End If
      PPXX% = InStr(FEMOX$, "/20")
      If PPXX% > 0 Then
        FEMOX$ = Left$(FEMOX$, PPXX% - 1) + "/" + Mid$(FEMOX$, PPXX% + 3)
      End If
      FEXI% = FECHANUM(FEMOX$)
      '
      'VAX% = 1
      'If Mid$(ATT$, 45, 1) = "B" Then VAX% = 2
      '
      NRO& = Val(Mid$(ATT$, 57, 8))
      'SUCUI% = Val(Mid$(ATT$, 46, 4))
      DOPRI$ = TIX$ + "." + TRIM$(Str$(NRO&))
      DOSEC$ = DOPRI$
      CODD$ = TRIM$(Str$(Val(Mid$(ATT$, 106, 10))))
      CII% = CODINT%(CODD$)
      '
      CANX$ = TRIM$(Mid$(ATT$, 158, 12))
      CANY$ = REPLACAR$(CANX$, ",", ".")
      CANTU = (-1) * Val(CANY$)
      If TIX$ = "NC" Then CANTU = (-1) * CANTU
      '
      IMPOX$ = TRIM$(Mid$(ATT$, 202, 12))
      IMPOY$ = REPLACAR$(IMPOX$, ",", ".")
      VUNI# = Val(IMPOY$)
      '
      'Call MOVISTO(FEXI%, CII%, "", "RT", DOPRI$, DOSEC$, RASOCLI$(NC%), VUNI#, "$", NC%, 1, CANTU)
      KKL& = KKL& + 1
FEX = FEXI%
AMASTOCK.Label64 = Str$(KKL&) & " " & FECHATEX$(FEX)
AMASTOCK.Label64.Refresh
'      If KKL& Mod 1000 = 0 Then
'MsgBox KKL& & "   " & Timer - TINI
'AAA = BBB
'TINI = Timer
'      End If
      '
   
   Wend
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("*.*")
   

End Sub

Sub IMPORMASTER()
   '
   Dim TXX$(128)
   Dim CODCAR$(16384), YACAR1%(16384), YACAR2%(16384)
   '
   CACOD% = 0
   LU$ = LUDAT$
   '
   Call COPYSTRU("LIPREVEN", "LIPRE001")
   AGREMAS% = 0: JJ& = 0: KK& = 0
   'For U& = 2 To ULTREG&("MASTER")
   '  X$ = REGLEIDO$("MASTER", U&)
   '  CACOD% = CACOD% + 1
   '  CODCAR$(CACOD%) = Mid$(X$, 77, 15)
   'Next U&
   '
   Label58.Visible = True
   IKK& = 0
   '
     ARIMPO$ = "F:\CLIENTES\ELROBLE\TRAIDOS\LISTOCK1.TXT"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
10   If EOF(N1%) Then
       Close #N1%
       GoTo 19
     End If
     '
     Line Input #N1%, AA$
     COD$ = TRIM$(Str$(Val(Mid$(AA$, 1, 8))))
     DEX$ = TRIM$(Mid$(AA$, 49, 30))
     GRUCO% = Val(Mid$(AA$, 23, 2))
     UNI$ = "U."   'TRIM$(Mid$(AA$, 64, 4))
     COU = 0
     MONEI% = 1
     FECO% = HOY(HOS$)
     '
     COSUNIX$ = TRIM$(Mid$(AA$, 101, 10))
     COSUN$ = REPLACAR(COSUNIX$, ",", ".")
     PREUNI# = Val(COSUN$)
     '
     COMENTA$ = TRIM$(Mid$(AA$, 296, 64))
     '
     
     If COD$ <> "" Then
       '
      If Len(COD$) <= 5 Then
      If DEX$ <> "" Then
       For KKI% = 1 To CACOD%
         If TRIM$(CODCAR$(KKI%)) = COD$ Then GoTo 11
       Next KKI%
       '
       If ULTREG&("MASTER") < 1 Then
         X$ = REGLEIDO$("MASTER", 1&)
         Call REPLA(X$, MKI$(1), 1, 2)
         Call GRAREG("MASTER", X$, 1)
       End If
       '
       CACOD% = CACOD% + 1
       CODCAR$(CACOD%) = AJUSTI$(COD$, 15)
       KKI% = CACOD%
       REGMA& = CACOD% + 1
       REMAS$ = String$(128, 0)
       Call REPLA(REMAS$, MKI$(CACOD%), 1, 2)
       Call REPLA(REMAS$, DEX$, 5, 64)
       Call REPLA(REMAS$, COD$, 77, 16)
       Call REPLA(REMAS$, Chr$(MONEI%), 95, 1)
       Call REPLA(REMAS$, UNI$, 109, 2)
       Call REPLA(REMAS$, MKS$(COU), 111, 4)
       Call REPLA(REMAS$, MKI$(FECO%), 115, 2)
       Call REPLA(REMAS$, MKI$(GRUCO%), 125, 2)
       Call GRAREG("MASTER", REMAS$, REGMA&)
       AGREMAS% = 1
       '
       If PREUNI# >= 0.005 Then
         LPP$ = REGBLAN$("LIPREVEN")
         Call REPLA(LPP$, MKI$(CACOD%), 1, 2)
         Call REPLA(LPP$, DEX$, 3, 42)
         Call REPLA(LPP$, MKD$(PREUNI#), 62, 8)
         JJ& = JJ& + 1
         Call GRAREG("LIPRE001", LPP$, JJ&)
       End If
       '
       If TRIM$(COMENTA$) <> "" Then
         LPP$ = REGBLAN$("ESPECIFI")
         Call REPLA(LPP$, MKI$(CACOD%), 1, 2)
         Call REPLA(LPP$, COMENTA$, 3, 62)
         KK& = KK& + 1
         Call GRAREG("ESPECIFI", LPP$, KK&)
       End If
      End If
     End If
     End If
     '
11 GoTo 10
   '
19 X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(CACOD%), 1, 2)
   Call GRAREG("MASTER", X$, 1&)
   '
   Call SETULTREG("MASTER", REGMA&)
   Call SETULTREG("LIPRE001", JJ&)
   Call SETULTREG("ESPECIFI", KK&)
   Call CIERRARCH("*.*")
   Close #N1%
   Call BAJALDISCO
Stop
   If AGREMAS% > 0 Then
      Call COMUNI("FLEXOFT debe Re-organizar la Base de Datos.\Una Vez Terminado, Repita la Importación")
      Call REVIMASTER
      Call FINAL("")
   End If
   '
End
   '
   '
51 For KKI% = 1 To 16384
      CODCAR$(KKI%) = ""
   Next KKI%
   '
   ARIMPO$ = "A:\FORMULAS.PRN"
     N1% = FILEOPEN(ARIMPO$, 1, 0)
     Line Input #N1%, AA$
     '
60   If EOF(N1%) Then
       Close #N1%
       GoTo 65
     End If
     '
     Line Input #N1%, AA$
     If TRIM$(Left$(AA$, 15)) <> "" Then
       CODCON$ = TRIM$(Left$(AA$, 13))
       Label58.Caption = CODCON$
       Label58.Refresh
     End If
     CODPAR$ = TRIM$(Mid$(AA$, 14, 15))
     If TRIM$(CODPAR$) <> "" Then
         COSUNIX$ = TRIM$(Mid$(AA$, 34, 10))
         USOUNIX$ = TRIM$(REPLACAR$(COSUNIX$, ",", "."))
         PPXX% = InStr(USOUNIX$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(USOUNIX$)
         ENTE$ = Left$(USOUNIX$, PPXX% - 1)
         DECI$ = Mid$(USOUNIX$, PPXX% + 1)
         While Len(DECI$) < 5: DECI$ = DECI$ + "0": Wend
         DECI$ = Left$(DECI$, 5)
         USO = Val(ENTE$) + (Val(DECI$)) / 100000
         '
         CI1% = CODINT%(CODCON$)
         CI2% = CODINT%(CODPAR$)
         If CI1% > 0 Then
           If CI2% > 0 Then
             If USO > 0 Then
               CODCAR$(CI1%) = CODCAR$(CI1%) + MKI$(CI2%) + MKS$(USO) + String$(2, 0)
               YACAR1%(CI1%) = 1
             End If
           End If
         End If
     End If
     GoTo 60
     '
65 PUNTCI$ = ""
   FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       X$ = REGLEIDO$("ESTRUNUE", I&)
       CI% = CVI(Left$(X$, 2) + String$(2, 0))
     PUNTCI$ = PUNTCI$ + MKI$(CI%)
   Next I&
   '
   '
   J& = 0
   For KKI% = 1 To 16384
     If CODCAR$(KKI%) <> "" Then
       ENCU% = 0: PX1% = 1
500    PX% = InStr(PX1%, PUNTCI$, MKI$(CIBAS%))
       If PX% < PX1% Then GoTo 70
       If (PX% Mod 2) <> 1 Then PX1% = PX% + 1: GoTo 500
       ENCU% = (PX% + 1) / 2
       GoTo 80
       '
70     XX1$ = MKI$(KKI%) + Left$(CODCAR$(KKI%), 120) + String$(128, 0)
       XX2$ = Mid$(CODCAR$(KKI%), 121, 256) + String$(256, 0)
       XX3$ = Mid$(CODCAR$(KKI%), 377)
       URE1& = ULTREG&("ESTRUNUE")
       URE2& = ULTREG&("ESTRUDOS")
       For KKK& = 1 + URE2& To 1 + URE1&
         Call GRAREG("ESTRUDOS", String$(255, 0) + Chr$(1), KKK&)
       Next KKK&
       Call GRAREG("ESTRUNUE", XX1$, URE1& + 1)
       Call GRAREG("ESTRUDOS", XX2$, URE1& + 1)
       PUNTCI$ = PUNTCI$ + MKI$(KKI%)
       '
       If XX3$ <> "" Then
         XX3$ = MKI$(KKI%) + XX3$ + String$(512, 0)
         RFF$ = RECFILT$("ESTRULAR", 1, MKI$(KKI%))
         RELAR& = 1 + ULTREG&("ESTRULAR")
         If Len(RFF$) >= 4 Then
           RELA& = CVS(RFF$)
           If RELA& > 0 Then
             If RELA& < RELAR& Then
               RELAR& = RELA&
             End If
           End If
         End If
         Call GRAREG("ESTRULAR", XX3$, RELAR&)
         Call CIERRARCH("ESTRULAR")
       End If
     End If
80 Next KKI%
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub


