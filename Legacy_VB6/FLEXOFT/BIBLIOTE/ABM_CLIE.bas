Attribute VB_Name = "ABM_CLIE"
Sub CARDATCLI(NC%)
    '
    Screen.MousePointer = 11
    '
    MODICLI.Text2.TEXT = TEL2CLI$(NC%)
    MODICLI.Text3.TEXT = RASOCLI$(NC%)
    MODICLI.Text4.TEXT = DOMICLI$(NC%)
    MODICLI.Text5.TEXT = CPOSCLI$(NC%)
    MODICLI.Text6.TEXT = LOCACLI$(NC%)
    MODICLI.Text8.TEXT = CSTRING$(MKS$(LIMICRE(NC%)), 3, 12, 2, 2)
    MODICLI.Text9.TEXT = PAISCLI$(NC%)
    MODICLI.Text10.TEXT = TEL1CLI$(NC%)
    MODICLI.Text11.TEXT = FAXCLI$(NC%)
    MODICLI.TEXT12.TEXT = EMAILCLI$(NC%)
    MODICLI.Text13.TEXT = CONTACTCLI$(NC%)
    MODICLI.Text15.TEXT = CUITCLI$(NC%)
    MODICLI.Text18.TEXT = NIBRCLI$(NC%)
    '
    MODICLI.COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    MODICLI.COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    MODICLI.COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    MODICLI.COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    MODICLI.COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    MODICLI.COTEXT(5).TEXT = TRIM$(Str$(VENDCLI%(NC%)))
    MODICLI.COTEXT(6).TEXT = TRIM$(CUEMADRE$(NC%))
    MODICLI.COTEXT(7).TEXT = TRIM$(Str$(GRUCOCLI%(NC%)))
    MODICLI.Text7.TEXT = TRIM$(DESCUCLI$(NC%))
    '
    For II = 0 To 7
       ECO$ = ""
       If TRIM$(MODICLI.COTEXT(II).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(II), 1)
          LF11% = LENFIELD(ARCHE$(II), 1)
          ARGU$ = CBIN$(TRIM$(MODICLI.COTEXT(II).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(II), ARGU$))
       End If
       '
       If ECO$ <> "" Then
             MODICLI.DETEXT(II).TEXT = ECO$
          Else
             MODICLI.DETEXT(II).TEXT = "No Definido"
       End If
    Next II
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub GRADATCLI()
    '
    If NC% < 0 Then
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(MODICLI.Text15.TEXT)
    If CUIT$ <> "" Then
      If VALICUIT%(CUIT$) <> 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Número de C.U.I.T. no Válido")
    '    If MODO% = 2 Then MODICLI.Text15.Text = CUITCLI$(NC%)
        MODICLI.Text15.SetFocus
        Exit Sub
      End If
    End If
    '
    If NC% = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
    If NC% < 0 Then ARX1$ = "PROVED1": ARX2$ = "PROVED2"
    '
    If MODO% = 1 Then
           REGX& = 1 + ULTREG&(ARX1$)
           XX$ = REGBLAN$(ARX1$) + REGBLAN$(ARX2$)
           Call REPLA(XX$, MKI$(NC%), 1, 2)
           If NC% = 0 Or REGICLI&(NC%) > 0 Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Número de Cuenta Ocupado.")
               Exit Sub
           End If
        Else
           REGX& = REGICLI&(NC%)
           If REGX& < 1 Or REGX& > ULTREG&(ARX1$) Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
               Exit Sub
           End If
           '
           XX$ = REGLEIDO$(ARX1$, REGX&) + REGLEIDO$(ARX2$, REGX&)
           If CVI(Left$(XX$, 2)) <> Abs(NC%) Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
               Exit Sub
           End If
    End If
    '
    CARASO% = 0
    If TRIM$(MODICLI.Text3.TEXT) <> TRIM$(Mid$(XX$, 3, 30)) Then CARASO% = 1
    '
    Call REPLA(XX$, TRIM$(MODICLI.Text3.TEXT), 3, 30)
    Call REPLA(XX$, TRIM$(MODICLI.Text4.TEXT), 33, 25)
    CPOST% = 0
    If Val(MODICLI.Text5.TEXT) > 0 Then
        If Val(MODICLI.Text5.TEXT) < 32768 Then
           CPOST% = Val(MODICLI.Text5.TEXT)
        End If
    End If
    Call REPLA(XX$, MKI$(CPOST%), 60, 2)
    Call REPLA(XX$, TRIM$(MODICLI.Text6.TEXT), 62, 25)
    Call REPLA(XX$, TRIM$(MODICLI.Text10.TEXT), 87, 10)
    Call REPLA(XX$, TRIM$(MODICLI.Text2.TEXT), 97, 10)
    Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
    Call REPLA(XX$, TRIM$(MODICLI.Text18.TEXT), 122, 15)
    Call REPLA(XX$, TRIM$(MODICLI.COTEXT(0).TEXT), 149, 1)
    If Abs(CVS(Mid$(XX$, 153, 4)) - Val(MODICLI.Text8.TEXT)) > 0.05 Then
      If DERACCE%("MODLICRE", "MODIFICACION LIMITE DE CREDITO") >= 1 Then
        Call REPLA(XX$, MKS$(Val(MODICLI.Text8.TEXT)), 153, 4)
      End If
    End If
    '
    PIVA% = VALIMI(MODICLI.COTEXT(1).TEXT, 0, 7)
    Call REPLA(XX$, Chr$(PIVA%), 136, 1)
    '
    CATIB% = VALIMI(MODICLI.COTEXT(2).TEXT, 0, 2)
    '
    LIPRE% = VALIMI(MODICLI.COTEXT(3).TEXT, 0, 255)
    If LIPRE% > 0 Then
       If ECOARCH("LISTAS", Chr$(LIPRE%)) = "" Then
          Call MENSERR(24, "Lista de Precios no Válida")
          LIPRE% = 0
       End If
    End If
    Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
    '
    COPAGI% = VALIMI(MODICLI.COTEXT(4).TEXT, 0, 255)
    If COPAGI% > 0 Then
       If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
       End If
    End If
    Call REPLA(XX$, Chr$(COPAGI%), 150, 1)
    '
    VENDE% = VALIMI(MODICLI.COTEXT(5).TEXT, 0, 255)
    If VENDE% > 0 Then
       If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
          Call MENSERR(24, "Vendedor no Válido")
          VENDE% = 0
       End If
    End If
    Call REPLA(XX$, Chr$(VENDE%), 152, 1)
    '
    CUECO% = VALIMI(MODICLI.COTEXT(7).TEXT, 0, 32767)
    If ECOARCH$("GRUCOCLI", MKI$(CUECO%)) = "" Then
        CUECO% = 0
    End If
    Call REPLA(XX$, MKI$(CUECO%), 157, 2)
    '
    CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(MODICLI.COTEXT(6).TEXT, 12))
    Call REPLA(XX$, MKI$(CCX%), 159, 2)
    '
    If MODO% = 1 Then
        Call REGAPP(ARX1$, Left$(XX$, 32))
        REGX& = ULTREG&(ARX1$)
        RGBLX$ = REGBLAN$(ARX2$)
        While ULTREG&(ARX2$) < REGX&
           Call REGAPP(ARX2$, RGBLX$)
        Wend
    End If
    '
    Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
    Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
    Call CIERRARCH(ARX1$)
    Call CIERRARCH(ARX2$)
    '
    DCC$ = TRIM$(MODICLI.Text7.TEXT)
    For UJ& = 1 To ULTREG&("DESCUCLI")
      YYY$ = REGLEIDO$("DESCUCLI", UJ&)
      If CVI(Left$(YYY$, 2)) = NC% Then
         Call REPLA(YYY$, DCC$, 3, 14)
         Call GRAREG("DESCUCLI", YYY$, UJ&)
         GoTo 99
      End If
    Next UJ&
    YYY$ = REGBLAN$("DESCUCLI")
    Call REPLA(YYY$, MKI$(NC%), 1, 2)
    Call REPLA(YYY$, DCC$, 3, 14)
    Call REGAPP("DESCUCLI", YYY$)
    '
99  Call CIERRARCH("DESCUCLI")
    '
    If NC% > 0 Then
       RFF$ = RECFILT$("BADACLIE", 1, MKI$(NC%))
       REBACLI& = 0
       If Len(RFF$) >= 4 Then
          REBACLI& = CVS(Left$(RFF$, 4))
       End If
       If REBACLI& < 1 Or REBACLI& > ULTREG&("BADACLIE") Then
          REBACLI& = 1 + ULTREG&("BADACLIE")
       End If
       '
       XX$ = REGBLAN$("BADACLIE")
       Call REPLA(XX$, MKI$(NC%), 1, 2)
       Call REPLA(XX$, MODICLI.Text3.TEXT, 3, 62)
       Call REPLA(XX$, MODICLI.Text2.TEXT, 65, 32)
       Call REPLA(XX$, MODICLI.Text4.TEXT, 97, 64)
       Call REPLA(XX$, MODICLI.Text5.TEXT, 161, 16)
       Call REPLA(XX$, MODICLI.Text6.TEXT, 177, 48)
       Call REPLA(XX$, MODICLI.Text9.TEXT, 225, 32)
       Call REPLA(XX$, MODICLI.Text10.TEXT, 257, 24)
       Call REPLA(XX$, MODICLI.Text11.TEXT, 281, 24)
       Call REPLA(XX$, MODICLI.TEXT12.TEXT, 305, 48)
       Call REPLA(XX$, MODICLI.Text13.TEXT, 385, 48)
       Call GRAREG("BADACLIE", XX$, REBACLI&)
       Call CIERRARCH("BADACLIE")
    End If
    '
    If CARASO% = 1 Then
       Call FRESHECHO(ARX1$)
       'FORMUSEL.TABLA(ITX%).Clear
    End If
    '
    Call CLEARFORMCLI
    If MODO% = 1 Then
        MODICLI.Text1.TEXT = ""
        NC% = 0
    End If
    MODICLI.Text1.SetFocus
    Screen.MousePointer = 1
    '
End Sub
'
Sub CLEARFORMCLI()
    '
    MODICLI.Text2.TEXT = ""
    MODICLI.Text3.TEXT = ""
    MODICLI.Text4.TEXT = ""
    MODICLI.Text5.TEXT = ""
    MODICLI.Text6.TEXT = ""
    MODICLI.Text7.TEXT = ""
    MODICLI.Text8.TEXT = ""
    MODICLI.Text9.TEXT = ""
    MODICLI.Text10.TEXT = ""
    MODICLI.Text11.TEXT = ""
    MODICLI.TEXT12.TEXT = ""
    MODICLI.Text13.TEXT = ""
    MODICLI.Text15.TEXT = ""
    MODICLI.Text18.TEXT = ""
    '
    For II = 0 To 7
       MODICLI.COTEXT(II).TEXT = ""
       MODICLI.DETEXT(II).TEXT = ""
    Next II
    '
    MODICLI.Refresh
    NC% = 0
    '
End Sub
'

