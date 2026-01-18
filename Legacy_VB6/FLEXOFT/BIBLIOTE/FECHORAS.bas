Attribute VB_Name = "FECHORAS"


Dim DIAS%(12)
Dim IDS%(12)
Dim DSX$(6)
Dim DIASE As String
'
Public INIMES%(16)    ' VECTOR CON LA FECHA  AL INICIO DEL MES ACTUAL EN INIMES%(0)
Public DENOMES$(16)   ' Y A LOS SIGUIENTES 12 MESESCORRESPONDIENTE

Function FECHATEX$(FEX)
    '
    Dim FS$(3), FEXX(3)
    FEXU = FEX
    
    
    '
10  If FEX = 0 Then
        F0$ = "        "
        GoTo 435
    End If
    '
    FEXX(3) = Int(FEX / 13 / 32) + 80
    FEX = FEX - 32 * 13 * (FEXX(3) - 80)
    FEXX(2) = Int(FEX / 32)
    FEXX(1) = FEX - 32 * FEXX(2)
    '
    While FEXX(3) > 99
       FEXX(3) = FEXX(3) - 100                                ' rutina a¤o 2000
    Wend
    '
    F0$ = ""
    For II% = 1 To 3
        F$ = Str$(FEXX(II%))
        F$ = Right$(F$, Len(F$) - 1)
        FS$(II%) = Right$("00" + F$, 2)
        F0$ = F0$ + FS$(II%)
        If II% < 3 Then F0$ = F0$ + "/"
    Next II%
    '
    SGL$ = "20": If XVALO(Mid$(F0$, 7, 2)) >= 80 Then SGL$ = "19"
    On Error Resume Next
    CDATX = CDate(Left$(F0$, 6) + SGL$ + Mid$(F0$, 7))
    If Err Then
       On Error GoTo 0
       FEXU = FEXU - 1
       FEX = FEXU
       GoTo 10
    End If
    On Error GoTo 0
    '
435 FECHATEX$ = F0$
    FEX = FEXU
    '
End Function
'
Function FECHATEXLAR$(FEX)
    '
    F0$ = FECHATEX$(FEX)
    If F0$ = "        " Then
          FECHATEXLAR$ = "01/01/1980"
       Else
          SGL$ = "20": If XVALO(Mid$(F0$, 7, 2)) >= 80 Then SGL$ = "19"
          FECHATEXLAR$ = Left$(F0$, 6) + SGL$ + Mid$(F0$, 7)
    End If
    '
End Function

Function SafeDate(DTX) As Date
    SFD = DTX ' se modifico el Str por CStr por que daba un error 13 no coinciden los tipos
    If Left$(CStr(SFD), 8) = "12:00:00" Or TRIM$(CStr(SFD)) = "0:00:00" Or TRIM$(CStr(SFD)) = "00:00:00" Then
       SFD = CVDAT(1)   ' CORRESPONDE A 01/01/80
    End If
    SafeDate = SFD
End Function
'
Function FECHANUM(FETX$)
    '
    Dim FE1(3)
    AX% = 0
    FETY$ = REPLACAR$(FETX$, ".", "/")
    FETX$ = REPLACAR$(FETY$, "-", "/")
    If FETX$ <> "" Then
      If Mid$(FETX$, 3, 1) <> "/" Then FETX$ = Left$(FETX$, 2) + "/" + Mid$(FETX$, 3)
      If Mid$(FETX$, 6, 1) <> "/" Then FETX$ = Left$(FETX$, 5) + "/" + Mid$(FETX$, 6)
    End If
    '
    If ((Len(FETX$) <> 8) Or (Mid$(FETX$, 3, 1) <> "/") Or (Mid$(FETX$, 6, 1) <> "/")) Then
        AX% = 1: GoTo 610
    End If
    '
    For II% = 1 To 3
        FE1(II%) = XVALO(Mid$(FETX$, 3 * II% - 2, 2))
        If II% <> 3 Then If FE1(II%) < 1 Then GoTo 600
    Next II%
    '
    If FE1(3) < 60 Then FE1(3) = 100 + FE1(3)                 ' rutina a¤o 2000
    '
    If ((FE1(1) > 31) Or (FE1(2) > 12)) Then GoTo 600
    FE = FE1(1) + 32 * FE1(2) + 32 * 13 * (FE1(3) - 80)
    GoTo 610

600 AX% = 2
610 If AX% > 0 Then FE = 0
    If FE < (-32767) Then FE = (-32767)
    If FE > 32767 Then FE = 32767
    '
    FECHANUM = FE
    '
End Function
'

Function CVDAT(FFII%) As Date
   '
   FEX = FFII%
10 CVTEX$ = TRIM$(FECHATEX$(FEX))
   If CVTEX$ = "" Then CVTEX$ = "01/01/80"
   CVTEX1$ = CVTEX$
11 SGL$ = "20": If XVALO(Mid$(CVTEX1$, 7, 2)) >= 80 Then SGL$ = "19"
   On Error Resume Next
   CVDAT = CDate(Left$(CVTEX1$, 6) + SGL$ + Mid$(CVTEX1$, 7))
   If Err Then
     FEX = DIANTE(FEX)
     GoTo 10
   End If
   '
   If REPITE% = 0 Then
     If Format$(CVDAT, "dd/mm/yy") <> CVTEX$ Then
       CVTEX1$ = Mid$(CVTEX1$, 4, 2) + "/" + Left$(CVTEX1$, 2) + Mid$(CVTEX1$, 6)
       REPITE% = REPITE% + 1
       GoTo 11
     End If
   End If
   '
End Function

Function CVINT%(FECHAD)
   CVINT% = 0
   On Error Resume Next
   If CDbl(FECHAD) > 16000 Then
     TXTX$ = Format(FECHAD, "dd/mm/yy")
     CVINT% = FECHANUM(TXTX$)
     If Abs(CVINT%) <= 33 Then CVINT% = 0
   End If
   On Error GoTo 0
End Function

Function HOY(HO$)    ' revision del 23/09/03
    '
    Static YAVEFLIM%, DATIX$
    '
    If Date$ <> DATIX$ Then
       DATIX$ = Date$
       YAVEFLIM% = 0
    End If
    '
    YAVEFLIM% = 1
    '
    'HO$ = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 2)
    HO$ = Format$(Now, "dd/mm/yy")
    HOY = FECHANUM(HO$)
    '
    If YAVEFLIM% <> 1 Then
       If UCase$(App.EXEName) <> "FLEXOFT" Then
         YAVEFLIM% = 1
         EPL$ = UCase$(EMPRELI$)
         If Left$(EPL$, 6) = "AMBEST" Then Exit Function
         If InStr(EPL$, "BORENS") > 0 Then Exit Function
         If HOY < 8773 + 15 Then Exit Function ' Corresponde a 05/01/2001
         '
         LU$ = LUDAT$
         Dim MSX As String * 128
         NX% = FILEOPEN%(LU$ + "MASTER.DAT", 0, 128)
         If LOF(NX%) >= 128 Then
            Get #NX%, 1, MSX$
            FLIM% = CVI(Mid$(MSX$, 124, 2))
            '
            If FLIM% < HOY + 15 Then Close: End
            If CVI(Mid$(MSX$, 126, 2)) <> 9876 - FLIM% Then
               Mid$(MSX$, 126, 2) = MKI$(9876 - FLIM%)
               Put #NX%, 1, MSX$
            End If
            PPX% = 23
            For i% = 124 To 127
              PPX% = PPX% + (i% - 123) * Asc(Mid$(MSX$, i%, 1))
            Next i%
            PPX% = PPX% Mod 256
            If Asc(Mid$(MSX$, 128, 1)) <> PPX% Then
              Mid$(MSX$, 128, 1) = Chr$(PPX%)
              Put #NX%, 1, MSX$
            End If
         End If
         Close #NX%
       End If
    End If
    '
End Function
'

Function FETEXTO$(A%)
    '
    If A% < 1 Then
        FT$ = Space$(24)
        GoTo 90
    End If
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    '
    FEX = A%
    FE$ = FECHATEX$(FEX)
    MX% = XVALO(Mid$(FE$, 4, 2))
    If MX% < 1 Or MX% > 12 Then
        Call MENSERR(24, "Proceso Interrumpido: Mes Fuera de Rango en Funcion 'FETEXTO'")
        ' Call FUERASYS
    End If
    MXS$ = MES$(MX%)
    '
    FT$ = TRIM$(Str$(XVALO(Left$(FE$, 2)))) + " de " + MXS$ + " de "
    If XVALO(Mid$(FE$, 7, 2)) >= 80 Then
           FT$ = FT$ + "19"
        Else
           FT$ = FT$ + "20"
    End If
    '
    FT$ = FT$ + Mid$(FE$, 7, 2)
    '
90  FETEXTO$ = FT$
    '
End Function
'

Function HORATEX$(XLL)
    '
    Static MODORA%
    If MODORA% < 1 Then
       MODORA% = 1
       If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
          MODORA% = 2
       End If
    End If
    '
    If XLL = "" Then XLL = 0
    On Error Resume Next
    XL% = Int(XVALO(XLL))
    If Err Then
       On Error GoTo 0
       HORATEX$ = CStr(XLL)
       Exit Function
    End If
    On Error GoTo 0
    If MODORA% <> 2 Then
         HORA% = Int(XL% / 100)
         MINU% = XL% Mod 100
         SEGU% = 100 * (XVALO(XLL) - XL%)
15       If MINU% >= 60 Then
            HORA% = HORA% + 1
            MINU% = MINU% - 60
         End If
         If SEGU% > 59 Then
            MINU% = MINU% + 1
            SEGU% = SEGU% - 60
            GoTo 15
         End If
         If HORA% < 0 Or MINU% < 0 Or MINU% > 59 Then
             HORA% = 0: MINU% = 0: SEGU% = 0
         End If
         '
         HORAS$ = Mid$(Str$(HORA%), 2)
         While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
         MINUS$ = Mid$(Str$(MINU%), 2)
         While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
         SEGUS$ = Mid$(Str$(SEGU%), 2)
         While Len(SEGUS$) < 2: SEGUS$ = "0" + SEGUS$: Wend
         '
         HORATEX$ = HORAS$ + ":" + MINUS$ + ":" + SEGUS$
       Else
         HORA% = Int(XL% / 100)
         MINU% = 100 * (XL% Mod 100) / 60
         SEGU% = 10000 * (XLL - XL%) / 60
         If HORA% < 0 Or HORA% > 24 Or MINU% < 0 Or MINU% > 99 Or XL% > 2400 Then
             HORA% = 0: MINU% = 0
         End If
         '
         HORAS$ = Mid$(Str$(HORA%), 2)
         While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
         MINUS$ = Mid$(Str$(MINU%), 2)
         While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
         SEGUS$ = Mid$(Str$(SEGU%), 2)
         While Len(SEGUS$) < 2: SEGUS$ = "0" + SEGUS$: Wend
         '
         HORATEX$ = HORAS$ + "," + MINUS$ + "," + SEGUS$
    End If
    '
End Function
'

Function HORATEXNOR$(XL%)
    '
    HORA% = Int(XL% / 100)
    MINU% = XL% Mod 100
    If HORA% < 0 Or HORA% > 24 Or MINU% < 0 Or MINU% > 59 Or XL% > 2400 Then
        HORA% = 0: MINU% = 0
    End If
    '
    HORAS$ = Mid$(Str$(HORA%), 2)
    While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
    MINUS$ = Mid$(Str$(MINU%), 2)
    While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
    '
    HORATEXNOR$ = HORAS$ + ":" + MINUS$
    '
End Function
'
Function FECHAFORMATOINVERSO$(FE%)
'YYYY/MM/DD
    FETEX1$ = FETEXCOR$(FE%)
    ANO1$ = Year(FETEX1$)
    MES1$ = FORMATOCON0(Month(FETEX1$), 2)
    DI1$ = Left$(FETEX1$, 2)
    FECHAFORMATOINVERSO$ = ANO1$ & MES1$ & DI1$

End Function
'
Function HORANUM(X$, Optional MODO$)
    '
    Static MODORA%
    If MODORA% < 1 Then
       MODORA% = 1
       If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
          MODORA% = 2
       End If
    End If
    '
    XX$ = TRIM$(X$)
    Dim Y$(3)
    Y$(1) = "": Y$(2) = "": Y$(3) = "": J% = 1
    '
    For i% = 1 To Len(XX$)
        AI% = Asc(Mid$(XX$, i%, 1))
        If AI% < 48 Or AI% > 57 Then
                J% = J% + 1
                If J% > 3 Then GoTo 800
                GoTo 799
        End If
        Y$(J%) = Y$(J%) + Chr$(AI%)
799 Next i%
    '
800 HORA% = 0
    If XVALO(Y$(1)) <= 999 Then
           HORA% = XVALO(Y$(1))
        Else
           If UCase(MODO$) <> "NOMESS" Then
             Call MENSERR(24, "Argumento fuera de rango en 'HORANUM'")
           End If
    End If
    If MODORA% <> 2 Then
        MINU% = XVALO(Y$(2)) ' If XVALO(Y$(2)) < 60 Then MINU% = XVALO(Y$(2)) Else Call MENSERR(24, "Argumento fuera de rango en 'HORANUM'")
        SEGU% = XVALO(Y$(3)) ' IF XVALO(Y$(3))<60 THEN SEGU%=XVALO(Y$(3) ELSE           ' Call MENSERR(24, "Argumento fuera de rango en 'HORANUM'")
        If SEGU% >= 60 Then SEGU% = SEGU% - 60: MINU% = MINU% + 1
        If MINU% >= 60 Then MINU% = MINU% - 60: HORA% = HORA% + 1:
      Else
        MINU% = 60 * XVALO(Y$(2)) / 100
        SEGU% = 60 * XVALO(Y$(3)) / 100
    End If
    '
    HORANUM = 100 * HORA% + MINU% + SEGU% / 100
    '
End Function

'
Function HORASEGU(X$)
    '
    XX$ = TRIM$(X$)
    Dim Y$(3)
    Y$(1) = "": Y$(2) = "": Y$(3) = "": J% = 1
    '
    For i% = 1 To Len(XX$)
        AI% = Asc(Mid$(XX$, i%, 1))
        If AI% < 48 Or AI% > 57 Then
                J% = J% + 1
                If J% > 3 Then GoTo 800
                GoTo 799
        End If
        Y$(J%) = Y$(J%) + Chr$(AI%)
799 Next i%
    '
800 HORASEGU = 60 * 60 * XVALO(Y$(1))
    HORASEGU = HORASEGU + 60 * XVALO(Y$(2))
    HORASEGU = HORASEGU + XVALO(Y$(3))
    '
End Function
'
Function HORAMINU$(XL%)
    '
    HORA% = 0
    MINU% = XL%
    '
    While MINU% >= 60
        HORA% = HORA% + 1
        MINU% = MINU% - 60
    Wend
    '
    HORAS$ = TRIM$(Str$(HORA%))
    While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
    MINUS$ = Mid$(Str$(MINU%), 2)
    While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
    '
    HORAMINU$ = HORAS$ + ":" + MINUS$
    '
End Function
'
Function HORMINSE$(TIE)
    '
    Static MODORA%
    If MODORA% < 1 Then
       MODORA% = 1
       If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
          MODORA% = 2
       End If
    End If
    '
    If MODORA% <> 2 Then
         ' SE COMENTA PORQUE NO PARECE TENER SENTIDO EL LIMITAR A 9999 HORAS
         'If TIE > 9999# * 3600 + 59 * 60 + 59 Then
         '   HORMINSE$ = "9999:59:59"
         '   Exit Function
         'End If
         '
         HORAS& = 0
         MINUS& = 0
         SEGUS& = Int(TIE + 0.5)
         '
         While SEGUS& >= 60
             MINUS& = MINUS& + 1
             SEGUS& = SEGUS& - 60
         Wend
         '
         While MINUS& >= 60
             HORAS& = HORAS& + 1
             MINUS& = MINUS& - 60
         Wend
         '
         HORAST$ = TRIM$(Str$(HORAS&))
         While Len(HORAST$) < 3: HORAST$ = " " + HORAST$: Wend
         MINUST$ = TRIM$(Str$(MINUS&))
         While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
         SEGUST$ = TRIM$(Str$(SEGUS&))
         While Len(SEGUST$) < 2: SEGUST$ = "0" + SEGUST$: Wend
         '
         HORMINSE$ = HORAST$ + ":" + MINUST$ + ":" + SEGUST$
         ' SE COMENTA PORQUE NO PARECE TENER SENTIDO EL LIMITAR A 9999 HORAS
         'If HORAS > 9999 Then HORMINSE$ = "9999:59:59"
       Else
         ' SE COMENTA PORQUE NO PARECE TENER SENTIDO EL LIMITAR A 9999 HORAS
         'If TIE > 9999# * 3600 + 59 * 60 + 59 Then
         '   HORMINSE$ = "9999,99,99"
         '   Exit Function
         'End If
         '
         HORAS& = 0
         MINUS& = 0
         SEGUS& = Int(TIE + 0.5)
         '
         While SEGUS& >= 60
             MINUS& = MINUS& + 1
             SEGUS& = SEGUS& - 60
         Wend
         '
         While MINUS& >= 60
             HORAS& = HORAS& + 1
             MINUS& = MINUS& - 60
         Wend
         '
         MINUS& = MINUS& * 100 / 60
         SEGUS& = SEGUS& * 100 / 60
         '
         HORAST$ = TRIM$(Str$(HORAS&))
         While Len(HORAST$) < 3: HORAST$ = " " + HORAST$: Wend
         MINUST$ = TRIM$(Str$(MINUS&))
         While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
         SEGUST$ = TRIM$(Str$(SEGUS&))
         While Len(SEGUST$) < 2: SEGUST$ = "0" + SEGUST$: Wend
         '
         HORMINSE$ = HORAST$ + "," + MINUST$ + "," + SEGUST$
         ' SE COMENTA PORQUE NO PARECE TENER SENTIDO EL LIMITAR A 9999 HORAS
         'If HORAS > 9999 Then HORMINSE$ = "9999,99,99"
    End If
    '
End Function
'
Function HORMINSE_LAR$(TIE)
    '
    Static MODORA%
    If MODORA% < 1 Then
       MODORA% = 1
       If UCase$(TRIM$(CONTROL$("PADIAMOD", "CONFHORA"))) = "MINUCENT" Then
          MODORA% = 2
       End If
    End If
    '
    If MODORA% <> 2 Then
         If TIE > 999999# * 3600 + 59 * 60 + 59 Then
            HORMINSE_LAR$ = "999999:59:59"
            Exit Function
         End If
         '
         HORAS& = 0
         MINUS& = 0
         SEGUS& = Int(TIE + 0.5)
         '
         While SEGUS& >= 60
             MINUS& = MINUS& + 1
             SEGUS& = SEGUS& - 60
         Wend
         '
         While MINUS& >= 60
             HORAS& = HORAS& + 1
             MINUS& = MINUS& - 60
         Wend
         '
         HORAST$ = TRIM$(Str$(HORAS&))
         While Len(HORAST$) < 6: HORAST$ = " " + HORAST$: Wend
         MINUST$ = TRIM$(Str$(MINUS&))
         While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
         SEGUST$ = TRIM$(Str$(SEGUS&))
         While Len(SEGUST$) < 2: SEGUST$ = "0" + SEGUST$: Wend
         '
         HORMINSE_LAR$ = HORAST$ + ":" + MINUST$ + ":" + SEGUST$
         If HORAS > 999999 Then HORMINSE_LAR$ = "999999:59:59"
       Else
         If TIE > 999999# * 3600 + 59 * 60 + 59 Then
            HORMINSE_LAR$ = "999999,99,99"
            Exit Function
         End If
         '
         HORAS& = 0
         MINUS& = 0
         SEGUS& = Int(TIE + 0.5)
         '
         While SEGUS& >= 60
             MINUS& = MINUS& + 1
             SEGUS& = SEGUS& - 60
         Wend
         '
         While MINUS& >= 60
             HORAS& = HORAS& + 1
             MINUS& = MINUS& - 60
         Wend
         '
         MINUS& = MINUS& * 100 / 60
         SEGUS& = SEGUS& * 100 / 60
         '
         HORAST$ = TRIM$(Str$(HORAS&))
         While Len(HORAST$) < 6: HORAST$ = " " + HORAST$: Wend
         MINUST$ = TRIM$(Str$(MINUS&))
         While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
         SEGUST$ = TRIM$(Str$(SEGUS&))
         While Len(SEGUST$) < 2: SEGUST$ = "0" + SEGUST$: Wend
         '
         HORMINSE_LAR$ = HORAST$ + "," + MINUST$ + "," + SEGUST$
         If HORAS > 999999 Then HORMINSE_LAR$ = "999999,99,99"
    End If
    '
End Function
'
Function HORMINSEMIL$(TIE)
    '
    If TIE > 999# * 3600 + 59 * 60 + 59 Then
       HORMINSEMIL$ = "999:59:59.999"
       Exit Function
    End If
    '
    HORAS& = 0
    MINUS& = 0
    SEGUS = TIE
    '
    While SEGUS >= 59.995
        MINUS& = MINUS& + 1
        SEGUS = SEGUS - 60
    Wend
    '
    While MINUS& >= 60
        HORAS& = HORAS& + 1
        MINUS& = MINUS& - 60
    Wend
    '
    HORAST$ = TRIM$(Str$(HORAS&))
    While Len(HORAST$) < 3: HORAST$ = " " + HORAST$: Wend
    MINUST$ = TRIM$(Str$(MINUS&))
    While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
    SEGUST$ = TRIM$(FORMATNUM$(SEGUS, "F6.3"))
    If TRIM$(SEGUST$) = "" Then SEGUST$ = "00.000"
    While Len(SEGUST$) < 6: SEGUST$ = "0" + SEGUST$: Wend
    '
    HORMINSEMIL$ = HORAST$ + ":" + MINUST$ + ":" + SEGUST$
    If HORAS > 999 Then HORMINSEMIL$ = "999:59:59.999"

End Function
'
Function HORMINSEDIEM$(TIE)
    '
    If TIE > 999# * 3600 + 59 * 60 + 59 Then
       HORMINSEDIEM$ = "999:59:59.9999"
       Exit Function
    End If
    '
    HORAS& = 0
    MINUS& = 0
    SEGUS = TIE
    '
    While SEGUS >= 59.9995
        MINUS& = MINUS& + 1
        SEGUS = SEGUS - 60
    Wend
    '
    While MINUS& >= 60
        HORAS& = HORAS& + 1
        MINUS& = MINUS& - 60
    Wend
    '
    HORAST$ = TRIM$(Str$(HORAS&))
    While Len(HORAST$) < 3: HORAST$ = " " + HORAST$: Wend
    MINUST$ = TRIM$(Str$(MINUS&))
    While Len(MINUST$) < 2: MINUST$ = "0" + MINUST$: Wend
    SEGUST$ = TRIM$(FORMATNUM$(SEGUS, "F7.4"))
    If TRIM$(SEGUST$) = "" Then SEGUST$ = "00.0000"
    While Len(SEGUST$) < 7: SEGUST$ = "0" + SEGUST$: Wend
    '
    HORMINSEDIEM$ = HORAST$ + ":" + MINUST$ + ":" + SEGUST$
    If HORAS > 999 Then HORMINSEDIEM$ = "999:59:59.9999"

End Function
'
Function TIMECENT$()
   '
   TINI = Timer
   CENTSEG = (Int(100 * (TINI - Int(TINI)) + 0.5)) Mod 100
   CECE$ = TRIM$(Str$(CENTSEG))
   While Len(CECE$) < 2
      CECE$ = "0" + CECE$
   Wend
   TIMECENT$ = Time$ + "." + CECE$
   '
End Function
'

Function DIFHORA(A, B)
    '
    If A > B Then
        HA = 0: HB = 0
        ' Call MENSERR(24, "Argumento negativo en rutina 'DIFHORA'")
        GoTo 99
    End If
    '
    HA = Int(A / 100) + (A Mod 100) / 60 + 100 * (A - Int(A)) / 3600
    HB = Int(B / 100) + (B Mod 100) / 60 + 100 * (B - Int(B)) / 3600
    '
99  DIFHORA = HB - HA
    '
End Function
'

Sub SELMESA(Des%, Has%, CORRESP$)
  '
  CORRESP$ = ""
  Des% = 0: Has% = 0
  '
  Call CARSELMESA
  Call SELECHO("SELMESA")
  MESASEL$ = VALACT1$("SELMESA")
  If TRIM$(MESASEL$) <> "" Then
    CORRESP$ = MESANO$(MESASEL$, Des%, Has%)
  End If
  '
End Sub
'
Function SELMESB$(Des%, Has%)
   '
   Call SELMESA$(Des%, Has%, CORRESP$)
   SELMESB$ = CORRESP$
   '
End Function
'
Sub DESHASFECHA(Des%, Has%, PERIO$, Optional MODO%)
    '
10  If Des% < 1 Or Has% < Des% Then
       FEX = HOY(HO$)
       Has% = FEX
       HASS$ = FECHATEX$(FEX)
       DESS$ = "01" + Mid$(HASS$, 3)
       Des% = FECHANUM(DESS$)
    End If
    If MODO% = 1 Then Des% = FECHANUM("01/01/80")
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        Des% = 0: Has% = 0: PERIO$ = ""
        Exit Sub
    End If
    '
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Des% > Has% Then
        GoTo 10
    End If
    FEX = Des%: PERIO$ = FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)

End Sub

Sub DESHAS(LEYENDA$, RENGLO%, Des%, Has%)
    '
MsgBox "FALTA REESCRIBIR DESHAS EN MODULO FECHORAS"
'    LEYE$ = Left$(trim$(LEYENDA$), 64)
'    While Len(LEYE$) < 37
'       LEYE$ = " " + LEYE$ + " "
'    Wend
'    '
'    ABA% = RENGLO% - 1: If ABA% > 22 Or ABA% < 6 Then ABA% = 20
'    ARR% = ABA% - 4
'    DER% = 76
'    IZQ% = DER% - Len(LEYE$) - 1
'    Call SETCOLOR("R-B")
'    Call VENTANA(ARR%, IZQ%, ABA%, DER%, "w", "")
'    LOCATE ARR% + 1, IZQ% + 1: COLOR 15: Print LEYE$
'    Call Label(ARR% + 3, IZQ% + 9, "Desde")
'    Call Label(ARR% + 3, DER% - 10, "Hasta")
    '
'650 DESS$ = ACCEPT$(ARR% + 3, IZQ% + 9, "D8", "", "")
'    DES% = FECHANUM(DESS$)
'    If DES% < 1 Then Exit Sub
'    '
'    HASS$ = ACCEPT$(ARR% + 3, DER% - 10, "D8", DESS$, "")
'    If Left$(HASS$, 1) = "<" Then GoTo 650
'    HAS% = FECHANUM(HASS$)
'    If HAS% < DES% Then
'       DES% = 0
'       HAS% = 0
'       Exit Sub
'    End If
'    '
End Sub
'
Function DIASEM%(FECHA)
  
30701 IDS%(1) = 0
      IDS%(2) = 3
      IDS%(3) = 3
      IDS%(4) = 6
      IDS%(5) = 1
      IDS%(6) = 4
      IDS%(7) = 6
      IDS%(8) = 2
      IDS%(9) = 5
      IDS%(10) = 0
      IDS%(11) = 3
      IDS%(12) = 5
      '
30702 DSX$(0) = "DOMINGO"
      DSX$(1) = "LUNES"
      DSX$(2) = "MARTES"
      DSX$(3) = "MIERCOLES"
      DSX$(4) = "JUEVES"
      DSX$(5) = "VIERNES"
      DSX$(6) = "SABADO"
      '
      FEX = FECHA
      FE$ = FECHATEX$(FEX)

30703 DIA% = XVALO(Left$(FE$, 2))
      MMX% = XVALO(Mid$(FE$, 4, 2))
      ANIO% = 1900 + XVALO(Mid$(FE$, 7))
      If ANIO% < 1980 Then ANIO% = ANIO% + 100
30704 BISI% = 0
      If (ANIO% Mod 4) <> 0 Then GoTo 30705
      If (ANIO% Mod 100) = 0 And (ANIO% Mod 400) <> 0 Then GoTo 30705
      BISI% = 1
30705 CABIS% = Int(ANIO% / 4) - Int(ANIO% / 100) + Int(ANIO% / 400)
      If BISI% = 1 Then CABIS% = CABIS% - 1
30706 OFMES% = IDS%(MMX%)
30707 OFDIA% = ANIO% + CABIS% + DIA% - 1 + OFMES%
      If BISI% = 1 Then
          If MMX% > 2 Then
              OFDIA% = OFDIA% + 1
          End If
      End If
30708 DS% = OFDIA% Mod 7
      DIASE$ = DSX$(DS%)
      '
      DIASEM% = DS%
      '
End Function
'
Function DSEM$(FECHA)
    '
    DS1% = DIASEM%(FECHA)
    DSEM$ = DIASE$
    '
End Function

Function DSEMIN$(FECHA)
    '
    DS1% = DIASEM%(FECHA)
    DSEMIN$ = Left$(DIASE$, 1) + LCase$(Mid$(DIASE$, 2))
    '
End Function

Function DIENTRE%(A, B)
    '
    a1 = A: B1 = B
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    If a1 > B1 Then
        DIENTRE% = 0
        Exit Function
    End If

    F1$ = FECHATEX$(a1)
    F2$ = FECHATEX$(B1)

    X% = XVALO(Left$(F2$, 2)) - XVALO(Left$(F1$, 2))
    Y2% = XVALO(Mid$(F2$, 4, 2))
    Y3% = XVALO(Mid$(F2$, 7, 2))
        If Y3% < 50 Then Y3% = 100 + Y3%
    FIN& = 12 * (Y3% - 80) + Y2%
    '
    X2% = XVALO(Mid$(F1$, 4, 2))
    X3% = XVALO(Mid$(F1$, 7, 2))
        If X3% < 50 Then X3% = 100 + X3%

    While 12 * (X3% - 80) + X2% < FIN&
        DIASI% = DIAS%(X2%)
        If (X3% Mod 4) = 0 Then
          If X3% <> 0 Then
            If X2% = 2 Then
                DIASI% = DIASI% + 1
            End If
          End If
        End If

        X% = X% + DIASI%
        X2% = X2% + 1
        If X2% > 12 Then
            X2% = 1
            X3% = X3% + 1
        End If
    Wend
    '
    DIENTRE% = X%
    '
End Function
'
Function DIASPOST(FEXU, CADIAS%)
    FEX = FEXU
    For KKI% = 1 To CADIAS%
      FEX = DIPOST(FEX)
    Next KKI%
    DIASPOST = FEX
End Function
Function CANT_DIASPOSTHABILES(FEXU, CADIAS%)
       'SE LE PASA UNA FECHA (NUMERICA) Y LA CANTIDAD DE DIAS POSTERIORES
       'LA FUNCION CALCULA LA CANTIDAD DE DIAS HABILES
       'EJ. SI LE PASO UN DIA LUNES Y LA CANTIDAD DE DIAS ES 2 DEVUELVE CANTIDAD DE DIAS 2 POR QUE NO HAY SABADOS Y DOMINGOS
       'SI SE LE PASA UN DIA FECHA QUE CAE JUEVES Y CANTIDAD DE DIAS 2 DEVULEVE LA CANT. DIAS 4 POR QUE LE SUMA UN SAB Y UN DOMINGO.
       
       SaboDom& = 0
       Dim FEXII  As Date
       J& = 0
       FE% = FEXU
       FEXII = CVDAT(FE%)
       CAD% = CADIAS%
        While J& <= CADIAS%
'            If Weekday(FEXII, vbMonday) = 7 Or Weekday(FEXII, vbMonday) = 6 Then  'Pregunto si es Sabado O Domingo
            If DIASEM%(FE%) = 0 Or DIASEM%(FE%) = 6 Then
              SaboDom& = SaboDom& + 1
            Else
              J& = J& + 1
            End If
            'FEXII = DateAdd("d", 1, FEXII)
            FE% = DIASPOST(FE%, 1)
       Wend
       CANT_DIASPOSTHABILES = CAD% + SaboDom&

End Function
Function DIPOST(DI)
    '
    Static DIXA$
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    DIX = DI
    DIBASE = DI
    DIX0 = DI
    DIXA$ = FECHATEX$(DIX0)
    '
20  DIX = DIX + 1
    DIXS$ = FECHATEX$(DIX)
    If DIXS$ = DIXA$ Then
       DIBASE = DIBASE + 1
       DIX = DIBASE
       GoTo 20
    End If
    DIXA$ = DIXS$
    '
    If XVALO(Mid$(DIXS$, 7, 2)) Mod 4 = 0 Then DIAS%(2) = 29
    FEX1 = XVALO(Left$(DIXS$, 2))
    FEX2 = XVALO(Mid$(DIXS$, 4, 2))
    FEX3 = XVALO(Mid$(DIXS$, 7, 2))
    '
    If FEX2 < 1 Or FEX2 > 12 Then GoTo 20
    If FEX1 < 1 Or FEX1 > DIAS%(FEX2) Then GoTo 20
    '
    DIPOST = DIX
    '
End Function
'
  
Function DIANTE(DI, Optional CANDIAS%)
    '
    'Static DIXA$
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    DIX = DI
    DIBASE = DI
    '
20  DIX = DIX - 1
    DIXS$ = FECHATEX$(DIX)
    If DIXS$ = DIXA$ Then
       DIBASE = DIBASE - 1
       DIX = DIBASE
       GoTo 20
    End If
    DIXA$ = DIXS$
    '
    If XVALO(Mid$(DIXS$, 7, 2)) Mod 4 = 0 Then DIAS%(2) = 29
    FEX1 = XVALO(Left$(DIXS$, 2))
    FEX2 = XVALO(Mid$(DIXS$, 4, 2))
    FEX3 = XVALO(Mid$(DIXS$, 7, 2))
    '
    If FEX2 < 1 Or FEX2 > 12 Then GoTo 20
    If FEX1 < 1 Or FEX1 > DIAS%(FEX2) Then GoTo 20
    '
    DIANTE = DIX
    '
    If CANDIAS% > 1 Then
      DIAN1 = DIANTE
      DIBAS1 = DIBASE
      If DIENTRE%(DIAN1, DIBAS1) < CANDIAS% Then
        GoTo 20
      End If
    End If
    '
End Function
'
Function DIASANTES(FEXU, CADIAS%)
    FEX = FEXU
    For KKI% = 1 To CADIAS%
      FEX = DIANTE(FEX)
    Next KKI%
    DIASANTES = FEX
End Function
'
Function FECHACRE%(FEDEPO%, HORACRE%)
    DIASACRE% = Int((HORACRE% + 12) / 24)
    FeAcredX = FEDEPO%
    While DIASACRE% > 0
      FEX = FeAcredX
      FeAcredX = DIPOST(FEX)
      FEX = FeAcredX
      If DIASEM%(FEX) > 0 Then
        If DIASEM%(FEX) < 6 Then
          DIASACRE% = DIASACRE% - 1
        End If
      End If
    Wend
    FECHACRE% = FeAcredX
End Function
'
Function COMESACT%()
   Static CMA%
   If CMA% = 0 Then
      FEX = HOY(HOS$)
      MACT% = XVALO(Mid$(HOS$, 4, 2))
10    FEY = DIANTE(FEX)
      If XVALO(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
         FEX = FEY
         GoTo 10
      End If
      CMA% = FEX
   End If
   '
COMESACT% = CMA%
End Function
Function COMESACTUAL%(FEX1)

      'DEVUELVE EL COMIENZO DEL MES PASANDOLE UNA FECHA
      FECTEX$ = FECHATEX$(FEX1)
      MACT% = XVALO(Mid$(FECTEX$, 4, 2))
10    FEY = DIANTE(FEX1)
      If XVALO(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
         FEX = FEY
         GoTo 10
      End If
      CMA% = FEX
      '
      COMESACTUAL% = CMA%
End Function
'
Function COMIENFE%(FECHAI)
   FEX = FECHAI
   FEXS$ = "01" + Mid$(FECHATEX$(FEX), 3)
   COMIENFE% = FECHANUM(FEXS$)
End Function

Function FINMESACT%()
   Static FMA%
   If FMA% = 0 Then
      FEX = HOY(HOS$)
      MACT% = XVALO(Mid$(HOS$, 4, 2))
10    FEY = DIPOST(FEX)
      If XVALO(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
         FEX = FEY
         GoTo 10
      End If
      FMA% = FEX
   End If
   '
FINMESACT% = FMA%
End Function
Function FINMESACTUAL%(FE)
      FEX1 = FE
      FETEX1$ = FECHATEX$(FEX1)
      MACT% = XVALO(Mid$(FETEX1$, 4, 2))
10    FEY = DIPOST(FEX1)
      If XVALO(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
         FEX1 = FEY
         GoTo 10
      End If
      FMA% = FEX1
      '
      FINMESACTUAL% = FMA%
End Function
'
Function COMESANT%()
   '
   Static CMA%
   If CMA% = 0 Then
      FEX = HOY(HOS$)
      MACT% = XVALO(Mid$(HOS$, 4, 2)) - 1
      If MACT% < 1 Then MACT% = 12
10    FEY = DIANTE(FEX)
      If XVALO(Left$(FECHATEX$(FEY), 2)) <= 1 Then
         If XVALO(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
            GoTo 20
         End If
      End If
      FEX = FEY
      GoTo 10
      '
20    CMA% = FEY
   End If
   '
COMESANT% = CMA%
End Function
'

Function SEMANU%(FECHA, DELX, ALX)
    '
    FFX = FECHA: FFS$ = FECHATEX$(FFX)
    COMAN$ = "01/01" + Mid$(FFS$, 6)
    COMANUX = FECHANUM(COMAN$)
    AAA = COMANUX
    DICOMAN% = DIASEM%(AAA)
    OFS% = 0
    If PRISEMIN% <> 1 Then
      If DICOMAN% = 0 Or DICOMAN% > 2 Then OFS% = 1
    End If
    '
    AAA = COMANUX
    B = FECHA
    CASEMA% = Int(DIENTRE%(AAA, B) / 7)
    B = FECHA
    If DIASEM%(B) < DICOMAN% Then
        CASEMA% = CASEMA% + 1
    End If
    '
    FFX = FECHA
    FF2X = FFX
    FF1% = DIASEM%(FF2X)
    While FF1% <> 1
23      FFX = FFX - 1
        FF2X = FFX
        FF1% = DIASEM%(FF2X)
        FF2X = FFX: FF2$ = FECHATEX$(FF2X)
               If XVALO(Left$(FF2$, 2)) < 1 Then GoTo 23
               If XVALO(Mid$(FF2$, 4, 2)) < 1 Then GoTo 23
               If XVALO(Mid$(FF2$, 4, 2)) > 12 Then GoTo 23
    Wend
    On Error Resume Next
    DELX = FFX
    On Error GoTo 0

    FFX = FECHA
    FF2X = FFX
    FF1 = DIASEM%(FF2X)
    While FF1 <> 0
33      FFX = FFX + 1
        FF2X = FFX
        FF1 = DIASEM%(FF2X)
        FF2X = FFX
        FF2$ = FECHATEX$(FF2X)
               If XVALO(Left$(FF2$, 2)) > 31 Then GoTo 33
               If XVALO(Mid$(FF2$, 4, 2)) < 1 Then GoTo 33
               If XVALO(Mid$(FF2$, 4, 2)) > 12 Then GoTo 33
    Wend
    ALX = FFX

    SEMANU% = CASEMA% + 1 - OFS%
    While SEMANU% < 1: SEMANU% = SEMANU% + 53: Wend
    '
    If PRISEMIN% = 1 Then
      If SEMANU% = 53 Then
        FEX = DELX: DDX1$ = FECHATEX$(FEX)
        FEX = ALX: DDX2$ = FECHATEX$(FEX)
        If Mid$(DDX2$, 7, 2) <> Mid$(DDX1$, 7, 2) Then SEMANU% = SEMANU% - 52
      End If
    End If
    '
End Function
Function UBICAR_DIA_SEMANA%(NUMERO_ORDEN1%, DIAX1$, MM$, YYYY$)
   'PASANDOLE POR EJEMPLO: MARTES, 3 (3ER  MARTES DEL MES DE MARZO)
   'DEVULVE LA FECHA EJEMPLO 17/03/15
   
   DIAX$ = TRIM$(UCase$(DIAX1$))
   DIAX$ = REPLACAR$(DIAX$, "Á", "A")
   DIAX$ = REPLACAR$(DIAX$, "É", "E")
   DIAX$ = REPLACAR$(DIAX$, "Í", "I")
   DIAX$ = REPLACAR$(DIAX$, "Ó", "O")
   DIAX$ = REPLACAR$(DIAX$, "Ú", "U")
   
   ORDENX% = 0
   UBICAR_DIA_SEMANA% = 0
   '
   If NUMERO_ORDEN1% < 1 Then Exit Function
   '
   
   ANOX$ = Right(YYYY$, 2)
  
   fech$ = "01/" & MM$ & "/" & ANOX$ ' FECHA PARAMETRO
   'ACA ERROR 'ESTA INVIERTIENDO LA FECHA EN VEZ DE 01/03/2017 ESTA DEVOLVIENDO EL NUMERO QUE EQUIVALE A 03/01/17

   FE% = FECHANUM(fech$)  ' TOMA EL NUMERO CORRESPONDIENTE DE LA FECHA
'   AAA = FECHATEX$(FE%)
   NUMERO_FIN% = XVALO(Mid$(FECHATEX$(FINMESACTUAL%(FE%)), 1, 2)) 'NUMERO DE FIN DEL MES EJEMPLO 30 O 31

   For i% = 0 To NUMERO_FIN% - 1              'REALIZA UN FOR DESDE 1 A FIN DE MES
      FE1% = DIASPOST(FE%, i%)                 'SUMA DE A 1 DIA CON LA FUNCION DIAPOST (POR LAS DUDAS POR SI NO ES CORRELATIVO)
      If DSEM$(FE1%) = DIAX$ Then             'SI ES IGUAL AL DIATEXTO INGRESA EN EL IF
        ORDENX% = ORDENX% + 1                 'ACUMULA EL ORDEN
        UBICAR_DIA_SEMANA% = FE1%             'SI ES EL DIA LO ASIGNA SIEMPRE POR SI ES SEMANA 5 Y NO EXISTE VA A DEVOLVER LA 4ta.
        If ORDENX% = NUMERO_ORDEN1% Then      'VALIDA SI COINCIDE CON EL NUMERO DE ORDEN (ENTONCES SI ES EL MISMO DIA Y EL MISMO NUMERO DE ORDEN EJ. 3er Miercoles
           UBICAR_DIA_SEMANA% = FE1%
           Exit For
        End If
      End If
   Next i%
   
   
End Function

'
Function FECHSEMAN(SEMANO$)
    'HAY QUE PASARLE EL NUMERO DE SEMANA Y EL AÑO EJ. 0821 (SEMANA 8 DE AÑO 21)
    SMNX$ = TRIM$(SEMANO$)
    While Len(SMNX$) < 4: SMNX$ = "0" + SMNX$: Wend
    SMNX$ = Right$(SMNX$, 4)
    '
    SMN% = XVALO(Left$(SMNX$, 2))
    If SMN% < 1 Or SMN% > 53 Then
      FECHSEMAN = 0
      Exit Function
    End If
    ANX$ = Mid$(SMNX$, 3)
    '
    FEXX$ = "01/01/" + ANX$
    FEXU = FECHANUM(FEXX$)
15  AA1 = SEMANU(FEXU, DELX, ALX)
    If AA1 > 52 Then AA1 = 0
    If AA1 < SMN% Then
      FEXU = DIPOST(ALX)
      GoTo 15
    End If
    FECHSEMAN = FEXU
    FFX$ = FECHATEX(FECHSEMAN)
End Function

Sub DATARGEN(Archi$, FEGE%, HOGE%, USGE%, TEGE%, CORRE$)
     '
     ARX$ = UCase$(TRIM$(Archi$))
     If ARX$ = "" Then Exit Sub
     '
     FEGE% = 0: HOGE% = 0: USGE% = 0: TEGE% = 0: CORRE$ = ""
     '
     FIN& = ULTREG&("DATARGEN")
     For U& = 1 To FIN&
        X$ = REGLEIDO$("DATARGEN", U&)
        If TRIM$(Left$(X$, 16)) = ARX$ Then
            FEGE% = CVI(Mid$(X$, 17, 2))
            HOGE% = CVI(Mid$(X$, 19, 2))
            USGE% = Asc(Mid$(X$, 21, 1))
            TEGE% = Asc(Mid$(X$, 22, 1))
            CORRE$ = Mid$(X$, 33, 96)
            GoTo 99
        End If
     Next U&
     '
99 End Sub
'
Sub GRARGEN(Archi$, TXT$)
     '
     ARX$ = UCase$(TRIM$(Archi$))
     If ARX$ = "" Then Exit Sub
     '
     FIN& = ULTREG&("DATARGEN")
     For U& = 1 To FIN&
        X$ = REGLEIDO$("DATARGEN", U&)
        If TRIM$(Left$(X$, 16)) = ARX$ Then GoTo 100
     Next U&
     U& = FIN& + 1
     X$ = REGBLAN$("DATARGEN")
     X$ = REPLACE$(X$, ARX$, 1, 16)
     '
100  Call REPLA(X$, MKI$(HOY(HO$)), 17, 2)
     Call REPLA(X$, MKI$(Int(HORANUM(Time$))), 19, 2)
     Call REPLA(X$, MKI$(USNBR% Mod 100), 21, 2)
     Call REPLA(X$, TXT$, 33, 96)
     Call GRAREG("DATARGEN", X$, U&)
     '
End Sub
'

Function FEGEAR%(Archi$)
     '
     Call DATARGEN(Archi$, FEGE%, HOGE%, USGE%, TEGE%, CORRE$)
     FEGEAR% = FEGE%
     '
End Function
'

Function HOGEAR%(Archi$)
     '
     Call DATARGEN(Archi$, FEGE%, HOGE%, USGE%, TEGE%, CORRE$)
     HOGEAR% = HOGE%
     '
End Function
'

Function USGEAR%(Archi$)
     '
     Call DATARGEN(Archi$, FEGE%, HOHE%, USGE%, TEGE%, CORRE$)
     USGEAR% = USGE%
     '
End Function
'
Function DiferenciaDiasEntreFechas(Desde%, Hasta%)

   DiferenciaDiasEntreFechas = DateDiff("d", FECHATEX$(Desde%), FECHATEX$(Hasta%))
End Function
Function CORRAR$(Archi$)
     '
     Call DATARGEN(Archi$, FEGE%, HOHE%, USGE%, TEGE%, CORRE$)
     CORRAR$ = CORRE$
     '
End Function

Function MESANO$(TXT$, Des%, Has%)
    '
    Dim MES$(12), DIAS%(12)
    MES$(1) = "ENERO": DIAS%(1) = 31
    MES$(2) = "FEBRERO": DIAS%(2) = 29
    MES$(3) = "MARZO": DIAS%(3) = 31
    MES$(4) = "ABRIL": DIAS%(4) = 30
    MES$(5) = "MAYO": DIAS%(5) = 31
    MES$(6) = "JUNIO": DIAS%(6) = 30
    MES$(7) = "JULIO": DIAS%(7) = 31
    MES$(8) = "AGOSTO": DIAS%(8) = 31
    MES$(9) = "SETIEMBRE": DIAS%(9) = 30
    MES$(10) = "OCTUBRE": DIAS%(10) = 31
    MES$(11) = "NOVIEMBRE": DIAS%(11) = 30
    MES$(12) = "DICIEMBRE": DIAS%(12) = 31
    '
    R$ = TRIM$(UCase$(TXT$))
    R1$ = "": R2$ = ""
    MMXX$ = "": DDXX% = 0: HHXX% = 0
    '
    PPXX% = InStr(R$, " ")
    If PPXX% > 3 Then
        R1$ = Left$(R$, 3)
        R2$ = TRIM$(Mid$(R$, PPXX% + 1))
        For JI% = 1 To 12
            If Left$(MES$(JI%), 3) = R1$ Then
                R1$ = MES$(JI%)
                GoTo 910
            End If
        Next JI%
        GoTo 999
        '
910     If XVALO(R2$) < 0 Or XVALO(R2$) > 2080 Then GoTo 999
        ANXX% = XVALO(R2$)
        While ANXX% > 99: ANXX% = ANXX% - 100: Wend
        R2$ = TRIM$(Str$(ANXX%))
        While Len(R2$) < 2: R2$ = "0" + R2$: Wend
        '
        ANXX% = ANXX% - 80
        If ANXX% < 0 Then ANXX% = ANXX% + 100
        DDXX% = 32 * 13 * ANXX% + 32 * JI% + 1
        HHXX% = DDXX% + DIAS%(JI%) - 1
        MMXX$ = R1$ + " " + R2$
        '
    End If
    '
999 MESANO$ = MMXX$
    Des% = DDXX%
    FFXX$ = FECHATEX$(HHXX%)
    Has% = FECHANUM(FFXX$)
    '
End Function
'

Function ABREMESAN$(CODIPERIO%, MODO%)
    '
    Dim MES$(12), DIAS%(12)
    MES$(1) = "Enero": DIAS%(1) = 31
    MES$(2) = "Febrero": DIAS%(2) = 29
    MES$(3) = "Marzo": DIAS%(3) = 31
    MES$(4) = "Abril": DIAS%(4) = 30
    MES$(5) = "Mayo": DIAS%(5) = 31
    MES$(6) = "Junio": DIAS%(6) = 30
    MES$(7) = "Julio": DIAS%(7) = 31
    MES$(8) = "Agosto": DIAS%(8) = 31
    MES$(9) = "Setiembre": DIAS%(9) = 30
    MES$(10) = "Octubre": DIAS%(10) = 31
    MES$(11) = "Noviembre": DIAS%(11) = 30
    MES$(12) = "Diciembre": DIAS%(12) = 31
    '
    MESIO% = CODIPERIO% Mod 12
    ANIO% = Int(CODIPERIO% / 12)
    If MESIO% = 0 Then
      MESIO% = 12
      ANIO% = ANIO% - 1
    End If
    If ANIO% > 99 Then ANIO% = ANIO% - 100
    ANIOX$ = TRIM$(Str$(ANIO%))
    While Len(ANIOX$) < 2: ANIOX$ = "0" + ANIOX$: Wend
    '
    LOMA% = 3: If MODO% = 1 Then LOMA% = 16
    MMXX$ = TRIM$(Left$(MES$(MESIO%), LOMA%)) + " " + ANIOX$
    ABREMESAN$ = MMXX$
    '
End Function
'

Sub CARSELMESA(Optional MASMESES%)
    '
    CANT_MESES_POST% = 2
    If MASMESES% > 0 Then CANT_MESES_POST% = MASMESES%: GoTo 5

    Static YACAR%
    If YACAR% = 0 Then
        '
5       Screen.MousePointer = 11
        U& = 0
        Dim MES$(12)
        '
        MES$(1) = "ENERO"
        MES$(2) = "FEBRERO"
        MES$(3) = "MARZO"
        MES$(4) = "ABRIL"
        MES$(5) = "MAYO"
        MES$(6) = "JUNIO"
        MES$(7) = "JULIO"
        MES$(8) = "AGOSTO"
        MES$(9) = "SETIEMBRE"
        MES$(10) = "OCTUBRE"
        MES$(11) = "NOVIEMBRE"
        MES$(12) = "DICIEMBRE"
        '
        HOO% = HOY(HOS$)
        MMXX% = XVALO(Mid$(HOS$, 4, 2)) + CANT_MESES_POST%
        AAXX% = XVALO(Mid$(HOS$, 7, 2))
        If MMXX% > 12 Then
          MMXX% = MMXX% - 12
          AAXX% = AAXX% + 1
        End If
        '
10      R2$ = TRIM$(Str$(AAXX%)): While Len(R2$) < 2: R2$ = "0" + R2$: Wend
        R1$ = MES$(MMXX%)
        '
        REMESA$ = R1$ + " " + R2$
        U& = U& + 1
        Call GRAREG("SELMESA", REMESA$ + Space$(12), U&)
        '
        If U& < 48 Then
            MMXX% = MMXX% - 1
            If MMXX% < 1 Then
                MMXX% = 12
                AAXX% = AAXX% - 1
                If AAXX% < 0 Then
                    AAXX% = AAXX% + 100
                End If
            End If
            GoTo 10
        End If
        '
        Call CIERRARCH("SELMESA")
        YACAR% = 1
        Screen.MousePointer = 1
        '
    End If
    '
End Sub

Sub CARSELMESA2(CANTIMESESMAS%)
    '
        ' SE JPASA PARAMAETRO PARA QUE MUESTRE MAS MESES HACIA ADELANTE
        If CANTIMESESMAS% < 1 Then CANTIMESESMAS% = 2
    
        '
        Screen.MousePointer = 11
        U& = 0
        Dim MES$(12)
        '
        MES$(1) = "ENERO"
        MES$(2) = "FEBRERO"
        MES$(3) = "MARZO"
        MES$(4) = "ABRIL"
        MES$(5) = "MAYO"
        MES$(6) = "JUNIO"
        MES$(7) = "JULIO"
        MES$(8) = "AGOSTO"
        MES$(9) = "SETIEMBRE"
        MES$(10) = "OCTUBRE"
        MES$(11) = "NOVIEMBRE"
        MES$(12) = "DICIEMBRE"
        '
        HOO% = HOY(HOS$)
        MMXX% = XVALO(Mid$(HOS$, 4, 2)) + CANTIMESESMAS%
        AAXX% = XVALO(Mid$(HOS$, 7, 2))
        If MMXX% > 12 Then
          MMXX% = MMXX% - 12
          AAXX% = AAXX% + 1
        End If
        '
10      R2$ = TRIM$(Str$(AAXX%)): While Len(R2$) < 2: R2$ = "0" + R2$: Wend
        R1$ = MES$(MMXX%)
        '
        REMESA$ = R1$ + " " + R2$
        U& = U& + 1
        Call GRAREG("SELMESA", REMESA$ + Space$(12), U&)
        '
        If U& < 48 Then
            MMXX% = MMXX% - 1
            If MMXX% < 1 Then
                MMXX% = 12
                AAXX% = AAXX% - 1
                If AAXX% < 0 Then
                    AAXX% = AAXX% + 100
                End If
            End If
            GoTo 10
        End If
        '
        Call CIERRARCH("SELMESA")
        Screen.MousePointer = 1
        '
    
    '
End Sub


Function METEXTO$(FECHA%, Optional ANIOCORTO%)
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    '
    FEX = FECHA%
    FETEY$ = FECHATEX$(FEX)
    METEX$ = MES$(XVALO(Mid$(FETEY$, 4, 2)))
    If ANIOCORTO% = 0 Then
        ANTEX$ = "19": If XVALO(Mid$(FETEY$, 7, 2)) < 70 Then ANTEX$ = "20"
    End If
    ANTEX$ = ANTEX$ + Mid$(FETEY$, 7, 2)
    '
    METEXTO$ = METEX$ + " " + ANTEX$
    '
End Function
'

Function FETEXCOR$(FECHA%)
    '
    Static CTX%, MES$(12)
    '
    If CTX% < 1 Then
      For KKI% = 1 To 12
        KKJJ$ = TRIM$(Str$(KKI%))
        While Len(KKJJ$) < 2: KKJJ$ = "0" + KKJJ$: Wend
        KKKK = CDate(KKJJ$ + "/" + KKJJ$ + "/" + KKJJ$)
        'KKKK% = FECHANUM(KKJJ$ + "/" + KKJJ$ + "/" + KKJJ$)
        MES$(KKI%) = Mid$(Format(KKKK, "DD-MMM-YYYY"), 4, 3)
        'MES$(KKI%) = Mid$(METEXTO$(KKKK%), 1, 3)
      Next KKI%
      CTX% = 1
    End If
    '

    FEX = FECHA%
    FETEY$ = FECHATEX$(FEX)
    METEX$ = MES$(XVALO(Mid$(FETEY$, 4, 2)))
    ANTEX$ = "19": If XVALO(Mid$(FETEY$, 7, 2)) < 70 Then ANTEX$ = "20"
    ANTEX$ = ANTEX$ + Mid$(FETEY$, 7, 2)
    '
    
    FETEXCOR$ = Left$(FETEY$, 2) + "/" + METEX$ + "/" + ANTEX$
    '
End Function
'
Function FETEXCOR1$(FECHA%, Optional MASHORA$)
    ' SI SE LE PASA COMO SEGUNDO PARAMETRO LA PALABRA "HASTA"
    'SE LE CONCATENA A LA FECHA " 23:59:59"
    Static LANGUA$
    If FECHA% < 1 Then
       FETEXCOR1$ = "01/01/1980"
       Exit Function
    End If
    '
    If LANGUA$ = "" Then
         LANGUA$ = "ESP"
         VUELTA% = 0: FORFECHA$ = "02/ENE/2008"
         Call APERBASDAT("STOCKS")
10       strSQL = "select DoPriLar from movisto WHERE FechMov = '" + FORFECHA$ + "'"
         On Error Resume Next
         FLEXCONN.Execute strSQL
         If Err Then
            On Error GoTo 0
            VUELTA% = VUELTA% + 1
            LANGUA$ = "ING"
            FORFECHA$ = "02/JAN/2008"
            If VUELTA% > 1 Then
               Call MENSERR(24, "Error Fatal - Imposible Aplicar Formato de Fecha\en Servidor SQL-SERVER. Consulte.")
               Call FINAL("")
            End If
            On Error GoTo 0
            GoTo 10
         End If
         On Error GoTo 0
    End If
    '
    ' TOMA LA FECHA EN TEXTO CORTA Y LA PASA A INGLES
    FTC$ = UCase$(FETEXCOR$(FECHA))
    If LANGUA$ = "ING" Then
       If Mid$(FTC$, 4, 3) = "ENE" Then Mid$(FTC$, 4, 3) = "JAN"
       If Mid$(FTC$, 4, 3) = "ABR" Then Mid$(FTC$, 4, 3) = "APR"
       If Mid$(FTC$, 4, 3) = "AGO" Then Mid$(FTC$, 4, 3) = "AUG"
       If Mid$(FTC$, 4, 3) = "SET" Then Mid$(FTC$, 4, 3) = "SEP"
       If Mid$(FTC$, 4, 3) = "DIC" Then Mid$(FTC$, 4, 3) = "DEC"
    End If
    '
    ' SI LA CONFIGURACION REGIONAL DE LA PC ES INGLES Y EL SQL SERVER EN ESPAÑOL,
    ' CAMBIO EL FORMATO DE LA FECHA QUE VIENE DEL FETEXCOR (CONF REGIONAL EN INGLES) A ESPAÑOL
    MesConfiguraRegional$ = Mid$(FTC$, 4, 3)
    If InStr(1, "JANAPRAUGDEC", MesConfiguraRegional$) > 0 And LANGUA$ = "ESP" Then
       If Mid$(FTC$, 4, 3) = "JAN" Then Mid$(FTC$, 4, 3) = "JAN"
       If Mid$(FTC$, 4, 3) = "APR" Then Mid$(FTC$, 4, 3) = "ABR"
       If Mid$(FTC$, 4, 3) = "AUG" Then Mid$(FTC$, 4, 3) = "AGO"
       If Mid$(FTC$, 4, 3) = "DEC" Then Mid$(FTC$, 4, 3) = "DIC"
    End If
    '
    If TRIM$(UCase$(MASHORA$)) = "HASTA" Then ' VALIDA SI TIENE EL PARAMETRO "HASTA"
       FTC$ = FTC$ & " 23:59:00" ' LE AGREGA LA HORA
    End If
    '
    FETEXCOR1$ = FTC$
End Function
'
Function FETEXCOR3$(FECHA%, Optional MASHORA$)
    ' SI SE LE PASA COMO SEGUNDO PARAMETRO LA PALABRA "HASTA"
    'SE LE CONCATENA A LA FECHA " 23:59:59"
    Static LANGUA$
    If FECHA% < 1 Then
       FETEXCOR3$ = "01/01/1980"
       Exit Function
    End If
    '
    If LANGUA$ = "" Then
         LANGUA$ = "ESP"
         VUELTA% = 0: FORFECHA$ = "02/ENE/2008"
         Call APERBASDAT("STOCKS")
10       strSQL = "select DoPriLar from movisto WHERE FechMov = '" + FORFECHA$ + "'"
         On Error Resume Next
         FLEXCONN.Execute strSQL
         If Err Then
            On Error GoTo 0
            VUELTA% = VUELTA% + 1
            LANGUA$ = "ING"
            FORFECHA$ = "02/JAN/2008"
            If VUELTA% > 1 Then
               Call MENSERR(24, "Error Fatal - Imposible Aplicar Formato de Fecha\en Servidor SQL-SERVER. Consulte.")
               Call FINAL("")
            End If
            On Error GoTo 0
            GoTo 10
         End If
         On Error GoTo 0
    End If
    '
    ' TOMA LA FECHA EN TEXTO CORTA Y LA PASA A INGLES
    FTC$ = UCase$(FETEXCOR$(FECHA))
    If LANGUA$ = "ING" Then
       If Mid$(FTC$, 4, 3) = "ENE" Then Mid$(FTC$, 4, 3) = "JAN"
       If Mid$(FTC$, 4, 3) = "ABR" Then Mid$(FTC$, 4, 3) = "APR"
       If Mid$(FTC$, 4, 3) = "AGO" Then Mid$(FTC$, 4, 3) = "AUG"
       If Mid$(FTC$, 4, 3) = "SET" Then Mid$(FTC$, 4, 3) = "SEP"
       If Mid$(FTC$, 4, 3) = "DIC" Then Mid$(FTC$, 4, 3) = "DEC"
    End If
    '
    If TRIM$(UCase$(MASHORA$)) = "HASTA" Then ' VALIDA SI TIENE EL PARAMETRO "HASTA"
       FTC$ = FTC$ & " 23:59:59" ' LE AGREGA LA HORA
       'EN ESTA FUNCION SE AGREGA QUE SI VARIABLE MASHORA$ TIENE CONTENIDO (02/SEP/2016 04:08:50 P.M.) QUE NO SEA 'HASTA' CONVIERTE LA HORA
       'CON EL SIGUIENTE CRITERIO:
       'CONVIERTE '02/SEP/2016 04:08:50 P.M.' EN '02/SEP/2016 16:08:50' POR QUE DE LA OTRA MANERA DA ERROR EN LA CONSULTA SQL
    ElseIf TRIM$(MASHORA$) <> "" Then  'EJ. DE CONTENIDO DE MASHORA$='02/SEP/2016 04:08:50 P.M.'
     MASHORA$ = TRIM$(UCase$(MASHORA$))
     Dim strArray() As String
     On Error Resume Next
     strArray = Split(MASHORA$, " ")
     fech = FECHATEX(CVINT(strArray(0)))
     HORAX$ = strArray(1)
     HORAX$ = HORAX$ + " " + strArray(2)
     On Error GoTo 0
     If Len(HORAX$) > 10 Then
       AMOPM$ = Right(MASHORA$, 4)
       If AMOPM$ = "P.M." Then
         If XVALO(Mid(HORAX$, 1, 2)) >= 12 Then
'            If XVALO(Mid(HORAX$, 3, 2)) > 0 Then
              TARDE% = 0: GoTo 50 'PARA QUE POR EJ. SI LA HORA ORIGINAL ES 12:21:00 P.M. NO DEVUELVA 24:21:00 P.M.
'             GoTo 99
'            End If
         End If
         TARDE% = 12
50    HORAX$ = SAFETEXT$(XVALO(Mid(HORAX$, 1, 2)) + TARDE%)
      MASHORA$ = " " & HORAX$ & Mid(strArray(1), 3, 6)
      FTC$ = FTC$ & MASHORA$

     End If
     ElseIf Len(MASHORA$) >= 8 And Len(MASHORA$) < 10 Then
     'para el caso si se le pasa en este formato  11:21:00 no entra en la opcion anterior
     MASHORA$ = " " & MASHORA$
     FTC$ = FTC$ & MASHORA$
     End If

    End If
    '
    FETEXCOR3$ = FTC$
     
'     If Len(MASHORA$) > 10 Then
'      'SI ES PM LE SUMA 12
'      'VALIDO QUE SI ES PM QUE NO SEA >  QUE 12 POR QUE DA ERROR
'       AMOPM$ = Right(MASHORA$, 4)
'       If AMOPM$ = "P.M." Then
'         If XVALO(Mid(MASHORA$, 12, 2)) >= 12 Then
'            If XVALO(Mid(MASHORA$, 15, 2)) > 0 Then
'              TARDE% = 0: GoTo 50 'PARA QUE POR EJ. SI LA HORA ORIGINAL ES 12:21:00 P.M. NO DEVUELVA 24:21:00 P.M.
''             GoTo 99
'            End If
'         End If
'         TARDE% = 12
'50    HORAX$ = SAFETEXT$(XVALO(Mid(MASHORA$, 10, 2)) + TARDE%)
'      MASHORA$ = " " & HORAX$ & Mid(MASHORA$, 12, 6)
'      FTC$ = FTC$ & MASHORA$
'
'     End If
'     ElseIf Len(MASHORA$) >= 8 And Len(MASHORA$) < 10 Then
'     'para el caso si se le pasa en este formato  11:21:00 no entra en la opcion anterior
'     MASHORA$ = " " & MASHORA$
'     FTC$ = FTC$ & MASHORA$
'     End If
'
'    End If
'    '
'    FETEXCOR3$ = FTC$
End Function
'


Function FETEXCOR2$(FECHA%, Optional MASHORA$)
    
    FTC$ = UCase$(FETEXCOR$(FECHA))
    If Mid$(FTC$, 4, 3) = "ENE" Then Mid$(FTC$, 4, 3) = "JAN"
    If Mid$(FTC$, 4, 3) = "ABR" Then Mid$(FTC$, 4, 3) = "APR"
    If Mid$(FTC$, 4, 3) = "AGO" Then Mid$(FTC$, 4, 3) = "AUG"
    If Mid$(FTC$, 4, 3) = "SET" Then Mid$(FTC$, 4, 3) = "SEP"
    If Mid$(FTC$, 4, 3) = "DIC" Then Mid$(FTC$, 4, 3) = "DEC"
    '
    FETEXCOR2$ = FTC$
    
End Function

Function NUNCA$()
    NUNCA$ = FETEXCOR1$(FECHANUM("01/01/80"))
End Function
'
Function AHORA$()
    AHORA$ = FETEXCOR1$(HOY(HOS$)) + " " + Format$(Time, "hh:mm:ss")
    'OJO, EN ARCON HUBO QUE CAMBIAR EL AHORA$ POR NOW POR QUE AL GRABAR DABA UN ERROR DE NO COINCIDEN LOS TIPOS.
    'SOLO EN LA PC DE ALEJANDRO NO PASABA ASI EN EL RESTO.
End Function

Function RANFECHA$(AI1%, BI1%)
   FEX = AI1%
   RFEX$ = FECHATEX$(FEX)
   FEX = BI1%
   RANFECHA$ = RFEX$ + " - " + FECHATEX$(FEX)
End Function

Function FEVENCIM%(FEMISI%, CPAGO%)

End Function

Sub GENVECMES()
    '
    HOII% = HOY(HOS$)
    INIMES%(0) = 0
    DENOMES$(0) = "VENCIDOS"
    INIMES%(1) = FECHANUM("01" + Mid$(FECHATEX$(HOII%), 3))
    DENOMES$(1) = UCase$(REPLACAR$(Mid$(FETEXCOR$(INIMES%(1)), 4), "/", " "))
    '
    IMES% = 1
    FINI = HOII%
    CMES% = XVALO(Mid$(FECHATEX$(FINI), 4, 2))
    Do While IMES% < 13
      FINI = DIPOST(FINI)
      If XVALO(Mid$(FECHATEX$(FINI), 4, 2)) <> CMES% Then
         IMES% = IMES% + 1
         CMES% = XVALO(Mid$(FECHATEX$(FINI), 4, 2))
         INIMES%(IMES%) = FINI
         DENOMES$(IMES%) = UCase$(REPLACAR$(Mid$(FETEXCOR$(INIMES%(IMES%)), 4), "/", " "))
      End If
    Loop
    '
End Sub

Function FINMESANT%()
   Static FMA%
   If FMA% = 0 Then
     FMA% = DIANTE(COMESACT%)
   End If
   '
   FINMESANT% = FMA%
End Function
Function FOR_DDMMYY%(FECHAF)
    '
    'CONVIERTE FECHA EN FORMATO MM/DD/AA A DD/MM/AA
    
    FOR_DDMMYY% = 0
    
    FETX$ = FECHAF
    FETY$ = REPLACAR$(FETX$, ".", "/")
    FETX$ = TRIM$(REPLACAR$(FETY$, "-", "/"))
    DIAX$ = Mid$(FETX$, 4, 2)
    Mesx$ = Mid$(FETX$, 1, 2)
    ANOX$ = Right(FETX$, 2)
    '
    FORMATODMY$ = DIAX$ + "/" + Mesx$ + "/" + ANOX$
    FOR_DDMMYY% = FECHANUM(FORMATODMY$)
    
End Function

Function FOR_DDMMYYYY%(FECHAF$)
    '
    'CONVIERTE FECHA EN FORMATO MM/DD/AA A DD/MM/AA
    
    If TRIM$(FECHAF$) = "" Then FOR_DDMMYYYY% = 33: Exit Function
    FETX$ = FECHAF
    FETY$ = REPLACAR$(FETX$, ".", "/")
    FETX$ = TRIM$(REPLACAR$(FETY$, "-", "/"))
    POS1% = InStr(FETX$, "/")
    If POS1% < 3 Then
       DIAX$ = XVALO(Mid$(FETX$, 1, POS1% - 1))
       DIAX$ = FORMATOCON0$(XVALO(DIAX$), 2)
    Else
       DIAX$ = Left(FETX$, POS1% - 1)
    End If
    '
    saldoFETX$ = Mid$(FETX$, POS1% + 1)
    POS2% = InStr(saldoFETX$, "/")
    If POS2% < 3 Then
       Mesx$ = XVALO(Mid$(saldoFETX$, 1, POS2% - 1))
       Mesx$ = FORMATOCON0$(XVALO(Mesx$), 2)
     Else
       Mesx$ = Left(saldoFETX$, POS2% - 1)
    End If
    
    
    saldoFETX$ = Mid$(saldoFETX$, POS2% + 1)
    ANOX$ = Right(saldoFETX$, 2)
    

    '
    FORMATODMY$ = DIAX$ + "/" + Mesx$ + "/" + ANOX$
    FOR_DDMMYYYY% = FECHANUM(FORMATODMY$)
    If FOR_DDMMYYYY% < 33 Then FOR_DDMMYYYY% = 33
End Function

Function MESNUM$(MESANO1$)
   'DEVUELVE SI SE LE PASA EL MES EL NUMERO DE MES EN FORMATO MM
   'SE HACE EL INSTRING POR SI SE LE PASA EN FORMATO ABRIL 15
   POS1% = InStr(TRIM$(MESANO1$), " ")
   Mesx$ = TRIM$(Mid$(MESANO1$, 1, POS1% - 1))
   If Mesx$ = "" Then Exit Function
   '
   Select Case UCase$(Mesx$)
    Case "ENERO"
      MESNUM$ = "01"
    Case "FEBRERO"
      MESNUM$ = "02"
    Case "MARZO"
      MESNUM$ = "03"
    Case "ABRIL"
      MESNUM$ = "04"
    Case "MAYO"
      MESNUM$ = "05"
    Case "JUNIO"
      MESNUM$ = "06"
    Case "JULIO"
      MESNUM$ = "07"
    Case "AGOSTO"
      MESNUM$ = "08"
    Case "SETIEMBRE"
      MESNUM$ = "09"
    Case "OCTUBRE"
      MESNUM$ = "10"
    Case "NOVIEMBRE"
      MESNUM$ = "11"
    Case "DICIEMBRE"
      MESNUM$ = "12"
   End Select
   
   
End Function

Function NUMMES%(MESENTEXTO$)

   MESPARAM$ = TRIM$(UCase(MESENTEXTO$))
   If MESPARAM$ = "" Then Exit Function
   POS1% = InStr(MESPARAM$, " ")
   If POS1% > 0 Then
      MESPARAM$ = Left$(MESPARAM$, POS1% - 1)
   End If
   Select Case MESPARAM$
        Case "ENERO": NUMMES% = 1
        Case "FEBRERO": NUMMES% = 2
        Case "MARZO": NUMMES% = 3
        Case "ABRIL": NUMMES% = 4
        Case "MAYO": NUMMES% = 5
        Case "JUNIO": NUMMES% = 6
        Case "JULIO": NUMMES% = 7
        Case "AGOSTO": NUMMES% = 8
        Case "SETIEMBRE": NUMMES% = 9
        Case "SEPTIEMBRE": NUMMES% = 9
        Case "OCTUBRE": NUMMES% = 10
        Case "NOVIEMBRE": NUMMES% = 11
        Case "DICIEMBRE": NUMMES% = 12
   End Select
End Function
Function NUMMES_TRESCARACTERES%(MESENTEXTO$)
   'FUNCIONA IGUAL QUE FUNCION NUMES PERO COMO PARAMETRO SE LE PASA EL MES SOLO LOS 3 PRIMEROS CARACTERES
   MESPARAM$ = TRIM$(UCase(MESENTEXTO$))
   If MESPARAM$ = "" Then Exit Function
   Select Case MESPARAM$
        Case "ENE": NUMMES_TRESCARACTERES% = 1
        Case "FEB": NUMMES_TRESCARACTERES% = 2
        Case "MAR": NUMMES_TRESCARACTERES% = 3
        Case "ABR": NUMMES_TRESCARACTERES% = 4
        Case "MAY": NUMMES_TRESCARACTERES% = 5
        Case "JUN": NUMMES_TRESCARACTERES% = 6
        Case "JUL": NUMMES_TRESCARACTERES% = 7
        Case "AGO": NUMMES_TRESCARACTERES% = 8
        Case "SET": NUMMES_TRESCARACTERES% = 9
        Case "SEP": NUMMES_TRESCARACTERES% = 9
        Case "OCT": NUMMES_TRESCARACTERES% = 10
        Case "NOV": NUMMES_TRESCARACTERES% = 11
        Case "DIC": NUMMES_TRESCARACTERES% = 12
   End Select
End Function


Function ENROQUE_DE_FECHAS(FECHA)
   
   'SI LA FECHA NO ES VALIDA DEVUELVE 0
   ' SI FECHA ES VALIDA DEVUELVE LA FECHA CON FORMATO YYYY,MM,DD
   
   fech = TRIM$(FECHA)
   '01/01/2009
    ANO$ = Mid$(fech, 7, 4)
    If Len(TRIM$(ANO$)) = 2 Then
      ANO1$ = Format(fech, dd, MM, YYYY)
      ANO2$ = Mid$(ANO1$, 7, 2)
      ANO$ = ANO2$ & ANO$
    End If
    MES$ = Mid$(fech, 4, 2)
    DIA$ = Mid$(fech, 1, 2)
    fech = ANO$ & "/" & MES$ & "/" & DIA$
    If IsDate(fech) Then
      ENROQUE_DE_FECHAS = fech
    Else
      MsgBox "Fecha No Válida"
      ENROQUE_DE_FECHAS = 0

    End If
End Function
Function FechaHoraNum(fech$)
   'DEVUELVE EN FORMATO NUMERICO TANTO LA FECHA COMO LA HORA CONSIDERANDO SI ES AM O PM UTILIZADO PARA ORDENAMIENTO DE GRILLA BAP
   fech$ = TRIM$(fech$)
   'SI NO TIENE HORA HACE CVINT Y SE VA
   If Len(fech$) <= 11 Then
      FechaHoraNum = CVINT(fech$)
      Exit Function
   End If
   
   SOLOFECHA$ = Mid$(fech$, 1, 10)
   SOLOFECHANUM = CVINT(SOLOFECHA$)
   HORA$ = Mid$(fech$, 12, 2)
   If UCase$(Right(fech$, 4)) = "P.M." Then HORA$ = TRIM$(Str$(XVALO(HORA$) + 12))
   HORAMINSEG$ = HORA$ & Mid$(fech$, 15, 6)
   HORAMINSEG$ = REPLACAR$(HORAMINSEG$, ":", "")
   SoloHoraMinSegNum = XVALO(HORAMINSEG$)
   FechaHoraNum = TRIM$(Str$(SOLOFECHANUM)) & TRIM$(Str$(SoloHoraMinSegNum))
   
   '02/SEP/2016 04:08:50 P.M.'
   
End Function
Function DIFERENCIAHORASCENTESIMAL(DESDEX, HASTAX)
  'RECIBE LA HORA DESDE Y HASTA EN FORMATO NUMERICO
  'DEVUELVE LA DIFERENCIA ENTRE DESDEX Y HASTAX QUE ES UN RANGO DE HORAS
  'LUEGO LO PASA A FORMATO CENTESIMAL
  'EJEMPLO LE PASO 10:00 A 15:30 DEVUELVE 5.50
'  DESDEX1 = DESDEX / 100
'  HASTAX1 = HASTAX / 100
'  DIFEREN = HASTAX - DESDEX
  DESDEX1 = CHoraSentecimal(DESDEX)
  HASTAX1 = CHoraSentecimal(HASTAX)
  DDIFEREN = HASTAX1 - DESDEX1
  DIFERENCIAHORASCENTESIMAL = DDIFEREN
'  DIFERENCIAHORASCENTESIMAL = CHoraSentecimal(DIFEREN)
   '
End Function

Function CHoraSentecimal(HORNUM)
  'RECIBE UNA HORA EN FORMATO NUMERICO EJ: SI EN FORMATO TEXTO ES 10:30 RECIBE 1030
  'y devuelve 10,50 en formato centesimal.

  HORNUM1 = HORNUM / 100
  SOLOHORAS = Int(HORNUM1 + 0.05)
  SALDO = HORNUM1 - SOLOHORAS
  MINUTCENTECIMAL = (100 / 60) * SALDO
  CHoraSentecimal = SOLOHORAS + MINUTCENTECIMAL

End Function

Function HAY_SOLAPAMIENTO_HORARIO%(DES1%, HAS1%, CONTRA_DES%, CONTRA_HAS%)
    'SE CREO PARA ENVAPLAST DONDE LOS  USUARIOS CARGAN DATOS DE HORARIOS DESDE HASTA EN DIFERNTES SECTORES
    HAY_SOLAPAMIENTO_HORARIO% = 0
    'DES1% Y HAS1% SON LOS DATOS QUE SE VAN A COMPARAR CONTRA LOS YA SUPUESTAMENTE GUARDADOS (CONTRA_DES%, CONTRA_HAS%)
    'EN EL CASO QUE DES1% Y HAS1% SE SOLAPEN CON LOS HORARIOS GUARDADOS LA FUNCION DEVUELVE 1 CASO CONTRARIO = 0
    If DES1% < CONTRA_HAS% Then
      If HAS1% > CONTRA_DES% Then
        HAY_SOLAPAMIENTO_HORARIO% = 1
      End If
    End If
      
    
End Function

Function HOYSERVER%()
    SQLX$ = "SELECT GETDATE() AS FECHA"
    Set RS = READSET(SQLX$, , 1)
    fff = RS!FECHA
    HOYSER$ = Format$(fff, "dd/mm/yy")
    HOYSERVER% = FECHANUM(HOYSER$)
End Function
Function NOWSERVER()
    SQLX$ = "SELECT GETDATE() AS FECHA"
    Set RS = READSET(SQLX$, , 1)
    fff = RS!FECHA
    NOWSERVER = fff
    RS.Close
    Set RS = Nothing
End Function

Function DEFECHASSQL()
'        Microsoft SQL Server ofrece algunas funciones para trabajar con fechas y horas. Estas son algunas:
'
'        - getdate(): retorna la fecha y hora actuales. Ejemplo:
'
'         select getdate();
'        - datepart(partedefecha,fecha): retorna la parte específica de una fecha, el año, trimestre, día, hora, etc.
'
'        Los valores para "partedefecha" pueden ser: year (año), quarter (cuarto), month (mes), day (dia), week (semana), hour (hora), minute (minuto), second (segundo) y millisecond (milisegundo). Ejemplos:
'
'         select datepart(month,getdate());
'        retorna el número de mes actual;
'
'         select datepart(day,getdate());
'        retorna el día actual;
'
'         select datepart(hour,getdate());
'        retorna la hora actual;
'
'        - datename(partedefecha,fecha): retorna el nombre de una parte específica de una fecha. Los valores para "partedefecha" pueden ser los mismos que se explicaron anteriormente. Ejemplos:
'
'         select datename(month,getdate());
'        retorna el nombre del mes actual;
'
'         select datename(day,getdate());
'        - dateadd(partedelafecha,numero,fecha): agrega un intervalo a la fecha especificada, es decir, retorna una fecha adicionando a la fecha enviada como tercer argumento, el intervalo de tiempo indicado por el primer parámetro, tantas veces como lo indica el segundo parámetro. Los valores para el primer argumento pueden ser: year (año), quarter (cuarto), month (mes), day (dia), week (semana), hour (hora), minute (minuto), second (segundo) y millisecond (milisegundo). Ejemplos:
'
'         select dateadd(day,3,'1980/11/02');
'        retorna "1980/11/05", agrega 3 días.
'
'         select dateadd(month,3,'1980/11/02');
'        retorna "1981/02/02", agrega 3 meses.
'
'         select dateadd(hour,2,'1980/11/02');
'        retorna "1980/02/02 2:00:00", agrega 2 horas.
'
'         select dateadd(minute,16,'1980/11/02');
'        retorna "1980/02/02 00:16:00", agrega 16 minutos.
'
'        - datediff(partedelafecha,fecha1,fecha2): calcula el intervalo de tiempo (según el primer argumento) entre las 2 fechas. El resultado es un valor entero que corresponde a fecha2-fecha1. Los valores de "partedelafecha) pueden ser los mismos que se especificaron anteriormente. Ejemplos:
'
'         select datediff (day,'2005/10/28','2006/10/28');
'        retorna 365 (días).
'
'         select datediff(month,'2005/10/28','2006/11/29');
'        retorna 13 (meses).
'
'        - day(fecha): retorna el día de la fecha especificada. Ejemplo:
'
'         select day(getdate());
'        - month(fecha): retorna el mes de la fecha especificada. Ejemplo:
'
'         select month(getdate());
'        - year(fecha): retorna el año de la fecha especificada. Ejemplo:
'
'         select year(getdate());
'        Se pueden emplear estas funciones enviando como argumento el nombre de un campo de tipo datetime o smalldatetime.

End Function
Public Function DiasDelMes(Optional ByVal FECHA As Variant) As Integer
    'DEBE RECIBIR LA FECHA DD/MM/YYYY


    Dim MES As Integer, Y  As Integer


    If IsMissing(FECHA) Then FECHA = Date


    If IsDate(FECHA) Then
        Y = Year(FECHA)
        MES = Month(FECHA)
    ElseIf IsNumeric(FECHA) Then
        Y = Year(Date)
        MES = IIf(FECHA > 0 And FECHA < 13, CInt(FECHA), 0)
    ElseIf VarType(FECHA) = vbString Then
        Y = Year(Date)
        Select Case UCase(Left$(FECHA, 3))
            Case "FEB":                                             MES = 2
            Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   MES = 1
            Case "APR", "JUN", "SEP", "NOV":                        MES = 4
        End Select
    End If


    Select Case MES
        Case 2:                     DiasDelMes = IIf(saltarYear(FECHA), 29, 28)
        Case 1, 3, 5, 7, 8, 10, 12: DiasDelMes = 31
        Case 4, 6, 9, 11:           DiasDelMes = 30
    End Select


End Function

Public Function saltarYear(ByVal valor As Variant) As Boolean


    On Error GoTo LocalError


    Dim iYear As Integer

    If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)


    If TypeName(iYear) = "Integer" Then
        saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
    End If
Exit Function
LocalError:

End Function
Function NumeroFinDelMes%(FE%, Optional FECHA As Date)

    C4 = FETEXCOR1$(FE%)
    NumeroFinDelMes% = Day(DateSerial(Year([C4]), Month([C4]) + 1, 1) - 1)
'    Dim Primer As Date
'    Dim Ultimo As Date
'    If fe% > 0 Then
'       fecha = FETEXCOR1$(fe%)
'    End If
'
'    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
'    Primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
'    FinDelMes = DateSerial(Year(fecha), Month(fecha) + 1, 0)
'otra forma
'    fecha = FECHATEX$(fe%)
'    If IsDate(fecha) Then
'        FinDelMes = DateAdd("m", 1, fecha)
'        FinDelMes = DateSerial(Year(FinDelMes), Month(FinDelMes), 1)
'        FinDelMes = DateAdd("d", -1, FinDelMes)
'    End If


End Function

Function FinDelMes(FE%)

    Dim Primer As Date
    Dim Ultimo As Date
    If FE% > 0 Then
       FECHA = FETEXCOR1$(FE%)
    End If

    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    Primer = DateSerial(Year(FECHA), Month(FECHA) + 0, 1)
    FinDelMes = DateSerial(Year(FECHA), Month(FECHA) + 1, 0)
'otra forma
'    fecha = FECHATEX$(fe%)
'    If IsDate(fecha) Then
'        FinDelMes = DateAdd("m", 1, fecha)
'        FinDelMes = DateSerial(Year(FinDelMes), Month(FinDelMes), 1)
'        FinDelMes = DateAdd("d", -1, FinDelMes)
'    End If


End Function


'Function TRUCOS_PARA_CALCULAR_FECHAS()
''''Saber la cantidad de días de un año determinado
''''Calcular la cantidad de días de un mes específico
''''Averiguar si un día corresponde a un fin de semana
''''Averiguar el último día de un Mes determinado
''''Averiguar el último día de una semana
''''
''''
''''
''''Agregar para el ejemplo 5 CommandButton: Command1, Command2, Command3, Command4 y Command5
''''
''''
''''
''''Código en un form
''''
''''
''''Private Sub Command1_Click()
''''MsgBox DiasDelAño(2006) & " días"
''''End Sub
''''
''''
''''Private Sub Command2_Click()
''''MsgBox DiasDelMes("05/09/2006") & " días"
''''End Sub
''''
''''
''''Private Sub Command3_Click()
''''MsgBox esFinSemana(Date)
''''End Sub
''''
''''
''''Private Sub Command4_Click()
''''MsgBox FinDelMes(Date)
''''End Sub
''''
''''
''''Private Sub Command5_Click()
''''MsgBox FinDeSemana(Date)
''''End Sub
''''
''''
''''
''''Private Sub Form_Load()
''''Command1.Caption = " Cantidad de Dias de un Año "
''''Command2.Caption = " Cantidad de Días de un Mes "
''''Command3.Caption = " Fin de semana ? "
''''Command4.Caption = " Obtener último día de un Mes "
''''Command5.Caption = " Obtener último día de una semana "
''''End Sub
''''
''''
''''
''''
''''
'''''###################################
''''
''''
'''''Funciones
''''
''''
'''''###################################
''''
''''
'''''Calcula la cantidad de días de un año
''''Public Function DiasDelAño(ByVal valor As Variant) As Integer
''''    If IsDate(valor) Or IsNumeric(valor) Then DiasDelAño = IIf(saltarYear(valor), 366, 365)
''''End Function
''''
''''
''''Public Function saltarYear(ByVal valor As Variant) As Boolean
''''
''''
''''    On Error GoTo LocalError
''''
''''
''''    Dim iYear As Integer
''''
''''    If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)
''''
''''
''''    If TypeName(iYear) = "Integer" Then
''''        saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
''''    End If
''''Exit Function
''''
''''
''''LocalError:
''''End Function
''''
''''
''''
'''''Calcula la cantidad de días de un mes de una fecha determinada
'''''La fecha debe tener el formato mm/dd/yyyy
'''''Si no se pasa el parámetro se asume la fecha de hoy
''''
''''
''''Public Function DiasDelMes(Optional ByVal Fecha As Variant) As Integer
''''
''''
''''
''''    Dim mes As Integer, y  As Integer
''''
''''
''''    If IsMissing(Fecha) Then Fecha = Date
''''
''''
''''    If IsDate(Fecha) Then
''''        y = Year(Fecha)
''''        mes = Month(Fecha)
''''    ElseIf IsNumeric(Fecha) Then
''''        y = Year(Date)
''''        mes = IIf(Fecha >; 0 And Fecha <; 13, CInt(Fecha), 0)
''''    ElseIf VarType(Fecha) = vbString Then
''''        y = Year(Date)
''''        Select Case UCase(Left$(Fecha, 3))
''''            Case "FEB":                                             mes = 2
''''            Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   mes = 1
''''            Case "APR", "JUN", "SEP", "NOV":                        mes = 4
''''        End Select
''''    End If
''''
''''
''''    Select Case mes
''''        Case 2:                     DiasDelMes = IIf(saltarYear(Fecha), 29, 28)
''''        Case 1, 3, 5, 7, 8, 10, 12: DiasDelMes = 31
''''        Case 4, 6, 9, 11:           DiasDelMes = 30
''''    End Select
''''
''''
''''End Function
''''
''''
''''
'''''Devuelve si un determinado día corresponde a un fin de semana
Public Function esFinSemana(ByVal FE) As Boolean

    FECHA = FECHATEX$(FE)
    If IsDate(FECHA) Then
       If (Weekday(FECHA) = 1) Or (Weekday(FECHA) = 7) Then
          esFinSemana = True
       End If
    End If
End Function
''''
''''
'''''Devuelve el último días del Mes
''''Public Function FinDelMes(Fecha As Variant) As Date
''''
''''
''''    If IsDate(Fecha) Then
''''        FinDelMes = DateAdd("m", 1, Fecha)
''''        FinDelMes = DateSerial(Year(FinDelMes), Month(FinDelMes), 1)
''''        FinDelMes = DateAdd("d", -1, FinDelMes)
''''    End If
''''
''''
''''End Function
''''
''''
'''''Devuelve el último día de la semana
''''
''''
''''Function FinDeSemana(ByVal Fecha As Date) As Date
''''
''''
''''Saber la cantidad de días de un año determinado
''''Calcular la cantidad de días de un mes específico
''''Averiguar si un día corresponde a un fin de semana
''''Averiguar el último día de un Mes determinado
''''Averiguar el último día de una semana
''''
''''
''''
''''Agregar para el ejemplo 5 CommandButton: Command1, Command2, Command3, Command4 y Command5
''''
''''
''''
''''Código en un form
''''
''''
''''Private Sub Command1_Click()
''''MsgBox DiasDelAño(2006) & " días"
''''End Sub
''''
''''
''''Private Sub Command2_Click()
''''MsgBox DiasDelMes("05/09/2006") & " días"
''''End Sub
''''
''''
''''Private Sub Command3_Click()
''''MsgBox esFinSemana(Date)
''''End Sub
''''
''''
''''Private Sub Command4_Click()
''''MsgBox FinDelMes(Date)
''''End Sub
''''
''''
''''Private Sub Command5_Click()
''''MsgBox FinDeSemana(Date)
''''End Sub
''''
''''
''''
''''Private Sub Form_Load()
''''Command1.Caption = " Cantidad de Dias de un Año "
''''Command2.Caption = " Cantidad de Días de un Mes "
''''Command3.Caption = " Fin de semana ? "
''''Command4.Caption = " Obtener último día de un Mes "
''''Command5.Caption = " Obtener último día de una semana "
''''End Sub
''''
''''
''''
''''
''''
'''''###################################
''''
''''
'''''Funciones
''''
''''
'''''###################################
''''
''''
'''''Calcula la cantidad de días de un año
''''Public Function DiasDelAño(ByVal valor As Variant) As Integer
''''    If IsDate(valor) Or IsNumeric(valor) Then DiasDelAño = IIf(saltarYear(valor), 366, 365)
''''End Function
''''
''''
''''Public Function saltarYear(ByVal valor As Variant) As Boolean
''''
''''
''''    On Error GoTo LocalError
''''
''''
''''    Dim iYear As Integer
''''
''''    If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)
''''
''''
''''    If TypeName(iYear) = "Integer" Then
''''        saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
''''    End If
''''Exit Function
''''
''''
''''LocalError:
''''End Function
''''
''''
''''
'''''Calcula la cantidad de días de un mes de una fecha determinada
'''''La fecha debe tener el formato mm/dd/yyyy
'''''Si no se pasa el parámetro se asume la fecha de hoy
''''
''''
''''Public Function DiasDelMes(Optional ByVal Fecha As Variant) As Integer
''''
''''
''''
''''    Dim mes As Integer, y  As Integer
''''
''''
''''    If IsMissing(Fecha) Then Fecha = Date
''''
''''
''''    If IsDate(Fecha) Then
''''        y = Year(Fecha)
''''        mes = Month(Fecha)
''''    ElseIf IsNumeric(Fecha) Then
''''        y = Year(Date)
''''        mes = IIf(Fecha >; 0 And Fecha <; 13, CInt(Fecha), 0)
''''    ElseIf VarType(Fecha) = vbString Then
''''        y = Year(Date)
''''        Select Case UCase(Left$(Fecha, 3))
''''            Case "FEB":                                             mes = 2
''''            Case "JAN", "MAR", "MAY", "JUL", "AUG", "OCT", "DEC":   mes = 1
''''            Case "APR", "JUN", "SEP", "NOV":                        mes = 4
''''        End Select
''''    End If
''''
''''
''''    Select Case mes
''''        Case 2:                     DiasDelMes = IIf(saltarYear(Fecha), 29, 28)
''''        Case 1, 3, 5, 7, 8, 10, 12: DiasDelMes = 31
''''        Case 4, 6, 9, 11:           DiasDelMes = 30
''''    End Select
''''
''''
''''End Function
''''
''''
''''
'''''Devuelve si un determinado día corresponde a un fin de semana
''''Public Function esFinSemana(ByVal Fecha As Variant) As Boolean
''''
''''
''''    If IsDate(Fecha) Then
''''       If (Weekday(Fecha) = 1) Or (Weekday(Fecha) = 7) Then
''''          esFinSemana = True
''''       End If
''''    End If
''''End Function
''''
''''
'''''Devuelve el último días del Mes
''''Public Function FinDelMes(Fecha As Variant) As Date
''''
''''
''''    If IsDate(Fecha) Then
''''        FinDelMes = DateAdd("m", 1, Fecha)
''''        FinDelMes = DateSerial(Year(FinDelMes), Month(FinDelMes), 1)
''''        FinDelMes = DateAdd("d", -1, FinDelMes)
''''    End If
''''
''''
''''End Function
''''
''''
'''''Devuelve el último día de la semana
''''
''''
''''Function FinDeSemana(ByVal Fecha As Date) As Date
''''
''''
''''    If IsDate(Fecha) Then
''''        FinDeSemana = FormatDateTime(Fecha - Weekday(Fecha) + 7, vbGeneralDate)
''''    End If
''''
''''
'End Function

Function DiasLaborales(FE1%, FE2%)
Dim J, k As Integer


Dim FechaInicio As Date
Dim FECHAFIN As Date
Dim FechaAux As Date
Dim Diferencia As Integer

FechaInicio = CVDAT(FE1%)
FECHAFIN = CVDAT(FE2%)

'FechaInicio = "03/01/2002" 'FechaInicio = txtFechaInicio
'FechaFin = "04/01/2002" 'FechaFin = txtFechaFin

Diferencia = DateDiff("d", FechaInicio, FECHAFIN) + 1

k = 0
FechaAux = FechaInicio
For J = 1 To Diferencia
  If Weekday(FechaAux, vbMonday) = 7 Then 'Pregunto si es Domingo
  ElseIf Weekday(FechaAux, vbMonday) = 6 Then 'Pregunto si es Sabado
  'ElseIf Feriado(FechaAux) Then 'Pregunto si es Feriado
  Else
    k = k + 1
  End If
  FechaAux = DateAdd("d", 1, FechaAux)
Next J
DiasLaborales = k

End Function

Function RangoSemanaSegunFecha$(FechaInicio%, RET_ULTIMODIASEMANA, RET_NUMSEMANA, Optional INICIADIAHABIL%, Optional FECHAFINRANGO%)
   'PASANDO UNA FECHA EN FORMATO NUMERICO, DEVUELVE LA FECHA DE INICIO  Y DE FIN DE LA SEMANA EN QUE CAE
   'TAMBIEN EL NUMERO DE SEMANA,EN EL NOMBRE DE LA FUNCION EL TEXTO: 17/06 al 23/06 del lunes al domingo inclusive
   'obtento el n.semana a partir de la fecha

    'SI INICIA CON DIAHABIL > 0 ES PARA SI SE ESTA RECORRIENDO UNAS FECHAS Y POR EJ. SE PASA
    'UNA FECHA QUE CAE SABADO O DOMINGO ARRANCA CON LA SEMANA POSTERIOR A PARTIR DEL LUNES
    If INICIADIAHABIL% > 0 Then
      d = DIASEM%(FechaInicio%)
      If d = 0 Then
         FechaInicio% = DIASPOST(FechaInicio%, 1)
      ElseIf d = 6 Then
         FechaInicio% = DIASPOST(FechaInicio%, 2)
      End If
    End If
      RET_NUMSEMANA = SEMANU%(FechaInicio%, A, B)
      RET_ULTIMODIASEMANA = B
      If FECHAFINRANGO% > 0 Then 'es para que si la la semana se pasa de la fecha del fin de rango llega hasta ese dia y no contemple toda la semana
         If FECHAFINRANGO% < RET_ULTIMODIASEMANA Then RET_ULTIMODIASEMANA = FECHAFINRANGO%
      End If
      FEIN$ = FECHATEX$(FechaInicio%)
      FEFIN$ = FECHATEX$(RET_ULTIMODIASEMANA)
      DESX$ = Left$(FEIN$, 5)
      HASX$ = Left$(FEFIN$, 5)
      
      RangoSemanaSegunFecha$ = DESX$ & " al " & HASX$
      RangoSemanaSegunFecha$ = FEIN$ & "-" & FEFIN$
      'si se esta recorriendo un for con un rango de fechas en la primera vuelta hay que pasarle en la primera que INICIADIAHABIL% =1
      'si por ejemplo el rango a recorrer es del 01/01/19 al 31/01/19 cada vez que vuelve a pasar por el for se le deberia pasar
      'el ultimo dia (RET_ULTIMODIASEMANA) + 1 para que vuleva a inicializar el mismo deberia ser lunes.
End Function
'CONSULTA SQL DEVULEVE CON FORMATO DE FECHAS A DEMANDA
'Select top 10  CONVERT(varchar,fechemisi,103) as [DD/MM/YYYY] from cajabanc
'EL TROZO CONVERT(varchar,fechemisi,103) ES EL QUE VALE EL RESTO ES LO CLASICO
'EN DISEÑADOR DE FORMULARIO EN LA CAJA DE TEXTO DEL CAMPO DEBERIA IR SOLO: CONVERT(varchar,fechemisi,103)
'EL NRO 103 EQUIVALE AL FORMATO [DD/MM/YYYY]
',2 as [YY.MM.DD], EL 4) ES  [DD.MM.YY] ,5 [DD-MM-YY]
',13 [DD MMM YYYY HH:MM:SS:MMM], 100 [MMM DD YYYY HH:SS (AM/PM)], 106 [DD MMM YYYY](31 Mar 2019)
'+ INFO EN https://www.sqlshack.com/es/funciones-y-formatos-de-sql-convert-date/

Function ListaDeAnios() As Integer
    
    With FRMZELECHO
        Call CARGA_ENCABEZADO(.MSF2, "Año/A10", FRMZELECHO)
    
        ANIOACTUAL% = Year(Now)
        For i = 0 To 14
            .MSF2.Rows = .MSF2.Rows + 1
            .MSF2.TextMatrix(.MSF2.Rows - 1, 1) = ANIOACTUAL% - i
        Next i
    End With
    
    FRMZELECHO.Caption = "Lista de Años"
    FRMZELECHO.Width = 2500
    FRMZELECHO.Show 1
    
    ListaDeAnios = XVALO(RESP_ZELE_VECT(1))
    
End Function
Function FECHAFORMATOLARGO$(DIAX%)
  If DIAX% < 33 Then DIAX% = 33
  DIA$ = FECHATEX$(DIAX%)
  FECHAFORMATOLARGO$ = FormatDateTime(DIA$, vbLongDate)
End Function
