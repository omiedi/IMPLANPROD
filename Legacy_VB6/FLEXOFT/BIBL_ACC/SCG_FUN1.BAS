Attribute VB_Name = "SCG_FUN"
Public NX1%, NX3%, NX4%, NXK5%, NX6%, NX7%, NX8%, NX9%, NX10%, NX11%, NX12%
Public ACUEN%, ADEN%
'
Public D0 As String * 64
Public D10 As String * 64
Public DX7 As String * 128
Public D1$, D2$, D3$, D4$, D5$, D6$, D7$, D8$
Public D11$, D12$, D13$, D14$, D15$, D16$, D17$, D18$, D19$
'
Public E0 As String * 64
Public E10 As String * 32
Public E20 As String * 32
Public EX As String * 128
Public E1$, E2$, E3$, E4$, E5$, E6$, E7$, E8$
Public E11$, E12$, E13$, E14$, E15$, E16$, E17$, E18$, E19$, E00$
'
Public C19 As String * 112
Public C20 As String * 20
Public I0 As String * 16
Public C0 As String * 64
'
Public NVI%, NUCUEN%
'
Public TIASI%
Public CUENTAINT%
Public NROASIEN
Public NROASDEF%
Public FECHASIEN%
Public IMPOPAS#
Public DETAPAS$
Public REFASIEN$
Public POSEDI%
'
Dim PRIMODCAJA&
'
Public YAESTABAS%
'
Function TIPCUE$(CUEI%)
   '
   Call APERBASDAT("PLANCUEN")
   '
   Set PLCUETEMP = Contable.OpenRecordset("SELECT * FROM PlanCuen WHERE StatCue > 0", dbOpenDynaset, dbReadOnly)
   '
   TICUEX$ = ""
   On Error Resume Next '
   PLCUETEMP.MoveFirst
   If Err Then GoTo 50
   '
   PLCUETEMP.FindFirst "CodInCue =" & CUEI%
   If PLCUETEMP.NoMatch = False Then
     TICUEX$ = TRIM$(UCase$(Left$(PLCUETEMP!TipoCue, 1)))
   End If
   '
50 On Error GoTo 0
   If TICUEX$ = "H" Then
         TICUEX$ = "Haber"
      ElseIf TICUEX$ = "I" Then
         TICUEX$ = "Ingresos"
      ElseIf TICUEX$ = "G" Then
         TICUEX$ = "Gastos"
      Else
         TICUEX$ = "Debe"
   End If
   '
   TIPCUE$ = TICUEX$
   '
99 End Function
'
Function OBSERCUE$(CUEI%)
   '
   Call APERBASDAT("PLANCUEN")
   '
   OBSERCUE$ = ""
   Set PLCUETEMP = Contable.OpenRecordset("SELECT * FROM PlanCuen WHERE StatCue > 0", dbOpenDynaset, dbReadOnly)
   '
   On Error Resume Next '
   PLCUETEMP.MoveFirst
   PLCUETEMP.FindFirst "CodInCue =" & CUEI%
   If PLCUETEMP.NoMatch = False Then
     OBSERCUE$ = PLCUETEMP.Observa
   End If
   On Error GoTo 0
   '
End Function
'
Function CUEINT7%(CUE$)
   '
   Static CUEANT$, CUEIANT%
   If CUE$ = CUEANT$ Then
      CUEINT7% = CUEIANT%
      GoTo 99
   End If
   '
   Call APERBASDAT("PLANCUEN")
   CUEINT7% = 0
   Set PLCUETEMP = Contable.OpenRecordset("SELECT * FROM PlanCuen WHERE CodExCue = '" & CUE$ & "' AND StatCue > 0", dbOpenDynaset, dbReadOnly)
   '
   On Error Resume Next '
   PLCUETEMP.MoveFirst
   If Err Then GoTo 50
   '
'   PLCUETEMP.FindFirst "CodExCue = '" & CUE$ & "'"
'   If PLCUETEMP.NoMatch = False Then
     CUEINT7% = PLCUETEMP.CODINCUE
'   End If
   On Error GoTo 0
   GoTo 98
   '
50 On Error GoTo 0
   URECU& = ULTREG&("CUENTAS")
   For KU& = 1 To URECU&
       Z1$ = REGLEIDO$("CUENTAS", KU&)
       If TRIM$(Left$(Z1$, 12)) = TRIM$(CUE$) Then
         CUEINT7% = KU&
         GoTo 98
       End If
   Next KU&
   '
98 CUEANT$ = CUE$
   CUEIANT% = CUEINT7%
   '
99 End Function
'
Function SARRASCUE#(CUEIII%)
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   SUCUI% = 0
   SUARRA# = 0
   FEINI = CDbl(CVDAT(COMEJ%))
   FEFIN = CDbl(CVDAT(FINEJ%))
   '
   SQLX$ = "SELECT * FROM DetaAsien "
   SQLX$ = SQLX$ + "WHERE (CueCoInt=" & CUEIII%
   SQLX$ = SQLX$ + " AND SucuAsien = " & SUCUI%
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) >= " & FEINI
   SQLX$ = SQLX$ + " AND CDBL(Fechasien) <= " & FEFIN
   SQLX$ = SQLX$ + " AND statpase >0 ); "
   'SQLX$ = SQLX$ + "ORDER BY FECHASIEN, ORFEASIEN ASC;"
   Set PASEASIEN = Contable.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   On Error Resume Next
   PASEASIEN.MoveFirst
   If Err Then GoTo 99
   '
   With PASEASIEN
     Do While .EOF = False
       SUARRA# = SUARRA# + !IDebePase - !IHabePase
       .MoveNext
     Loop
   End With
   '
99 SARRASCUE# = SUARRA#
   Screen.MousePointer = SCMA
   '
End Function

Function REFGENAS(IDASI)
   '
   Static IDASAN, REFASAN
   '
   If IDASI = IDASAN Then
     REFGENAS = REFASAN
     Exit Function
   End If
   '
   SQLX$ = "SELECT * FROM EncaBEAsien "
   SQLX$ = SQLX$ + "WHERE NuIdAsien = '" & IDASI & "'"
   Set ENCASI = Contable.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   REFGENAS = ""
   With ENCASI
     On Error Resume Next
     .MoveLast
     If Err < 1 Then
       REFGENAS = TRIM$(!Refeasien)
     End If
     On Error GoTo 0
   End With
   '
   IDASAN = IDASI
   REFASAN = REFGENAS
   '
End Function
'
Sub ABRENCAJA()
   If NXK5% < 1 Then
      NXK5% = FILEOPEN%(LUDAT$ + "ENCAJA" + SE$ + "." + EXTE$, 0, 64)
   End If
End Sub
'
Sub ABRESUCAJA()
   If NX6% < 1 Then
      NX6% = FILEOPEN%(LUDAT$ + "SUCAJA" + SE$ + "." + EXTE$, 0, 64)
   End If
End Sub
'
Sub ABREICAJA()
   If NX7% < 1 Then
      NX7% = FILEOPEN%(LU$ + "ISUCAJA" + SE$ + "." + EXTE$, 0, 128)
   End If
End Sub
'
Sub ABRENCDIARIO()
   If NX9% < 1 Then
      NX9% = FILEOPEN%(LUDAT$ + "ENCDIAR" + SE$ + "." + EXTE$, 0, 64)
   End If
End Sub
'
Sub ABREDIARIO()
   If NX10% < 1 Then
      NX10% = FILEOPEN%(LUDAT$ + "DIARIO" + SE$ + "." + EXTE$, 0, 32)
   End If
   '
   If NX12% < 1 Then
      NX12% = FILEOPEN%(LUDAT$ + "DTDIAR" + SE$ + "." + EXTE$, 0, 32)
   End If
End Sub
'
Sub ABREIDIAR()
   If NX11% < 1 Then
      NX11% = FILEOPEN%(LU$ + "IDIARIO" + SE$ + "." + EXTE$, 0, 128)
   End If
End Sub
'
Sub ABRESALMEN()
   If NX8% < 1 Then
      NX8% = FILEOPEN%(LUDAT$ + "CIERRE" + SE$ + "." + EXTE$, 0, 112)
   End If
End Sub

Sub ABRESALCAJ()
   If NX4% < 1 Then
      NX4% = FILEOPEN%(LUDAT$ + "SALCAJA" + "." + EXTE$, 0, 80)
   End If
End Sub
'
Sub CIERRACONTA()
   '
   If NX1% > 0 Then Close #NX1%: NX1% = 0
   If NX3% > 0 Then Close #NX3%: NX3% = 0
   If NX4% > 0 Then Close #NX4%: NX4% = 0
   If NXK5% > 0 Then Close #NXK5%: NXK5% = 0
   If NX6% > 0 Then Close #NX6%: NX6% = 0
   If NX7% > 0 Then Close #NX7%: NX7% = 0
   If NX8% > 0 Then Close #NX8%: NX8% = 0
   If NX9% > 0 Then Close #NX9%: NX9% = 0
   If NX10% > 0 Then Close #NX10%: NX10% = 0
   If NX11% > 0 Then Close #NX11%: NX11% = 0
   If NX12% > 0 Then Close #NX12%: NX12% = 0
   If ACUEN% > 0 Then Close #ACUEN%: ACUEN% = 0
   If ADEN% > 0 Then Close #ADEN%: ADEN% = 0
   NUCUEN% = 0: NVI% = 0
   '
End Sub
'
Sub LEENCAJA(RGENC%)
   Call ABRENCAJA
   Get #NXK5%, RGENC%, D0$
   D1$ = Mid$(D0$, 1, 4)
   D2$ = Mid$(D0$, 5, 2)
   D3$ = Mid$(D0$, 7, 2)
   D4$ = Mid$(D0$, 9, 8)
   D5$ = Mid$(D0$, 17, 2)
   D6$ = Mid$(D0$, 19, 2)
   D7$ = Mid$(D0$, 21, 40)
   D8$ = Mid$(D0$, 61, 4)
End Sub
'
Sub LEESUCAJA(RGSUC&)
   Call ABRESUCAJA
   Get #NX6%, RGSUC&, D10$
   D11$ = Mid$(D10$, 1, 2)
   D12$ = Mid$(D10$, 3, 4)
   D13$ = Mid$(D10$, 7, 2)
   D14$ = Mid$(D10$, 9, 2)
   D15$ = Mid$(D10$, 11, 2)
   D16$ = Mid$(D10$, 13, 8)
   D17$ = Mid$(D10$, 21, 2)
   D18$ = Mid$(D10$, 23, 32)
   D19$ = Mid$(D10$, 55, 4)
End Sub
'
Sub LEENCDIARIO(RGENC%)
   Call ABRENCDIARIO
   Get #NX9%, RGENC%, E0$
   E1$ = Mid$(E0$, 1, 4)
   E2$ = Mid$(E0$, 5, 2)
   E3$ = Mid$(E0$, 7, 2)
   E4$ = Mid$(E0$, 9, 8)
   E5$ = Mid$(E0$, 17, 2)
   E6$ = Mid$(E0$, 19, 2)
   E7$ = Mid$(E0$, 21, 40)
   E8$ = Mid$(E0$, 61, 4)
End Sub
'
Sub LEEDIARIO(RGSUC%)
   Call ABREDIARIO
   Get #NX10%, RGSUC%, E10$
   E11$ = Mid$(E10$, 1, 2)
   E12$ = Mid$(E10$, 3, 4)
   E13$ = Mid$(E10$, 7, 2)
   E14$ = Mid$(E10$, 9, 2)
   E15$ = Mid$(E10$, 11, 2)
   E16$ = Mid$(E10$, 13, 8)
   E17$ = Mid$(E10$, 21, 2)
   E18$ = Mid$(E10$, 23, 1)
   E19$ = Mid$(E10$, 24, 9)
   E00$ = ""
   If RGSUC% <= LOF(NX12%) / 32 Then
      Get #NX12%, RGSUC%, E20$
      E00$ = E20$
   End If
End Sub
'
Sub LEEPASCAJ(RGSUC&, MODO%)
   '
   Call LEESUCAJA(RGSUC&)
   CUENTAINT% = CVI(D11$)
   NROASIEN = CVS(D12$)
   NROASDEF% = CVI(D13$)
   FECHASIEN% = CVI(D15$)
   IMPOPAS# = CVD(D16$)
   DETAPAS$ = D18$
   REFASIEN$ = ""
   '
   If MODO% = 2 Then
      LI% = 0: LS% = MAXENC% + 1
11    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then GoTo 29 ' error de encabezado
      L% = LI% + E%
      If L% < 1 Or L% > MAXENC% Then GoTo 29 ' error de encabezado
      Call LEENCAJA(L%)
      If CVS(D1$) < NROASIEN - 0.01 Then LI% = L%: GoTo 11
      If CVS(D1$) > NROASIEN + 0.01 Then LS% = L%: GoTo 11
      REGENCA% = L%
      '
      Call LEENCAJA(REGENCA%)
      REFASIEN$ = D7$
   End If
   '
   Exit Sub
   '
29 Call MENSERR(24, "Error de Cabecera de Asiento " + TRIM$(Str$(NROASIEN)))
   Call FINAL("")
   '
End Sub
'
Sub LEEPASDIA(RGSUC%, MODO%)
   '
   Call LEEDIARIO(RGSUC%)
   '
   CUENTAINT% = CVI(E11$)
   NROASIEN = CVS(E12$)
   NROASDEF% = CVI(E13$)
   FECHASIEN% = CVI(E15$)
   IMPOPAS# = CVD(E16$)
   DETAPAS$ = ""
   REFASIEN$ = ""
   '
   If RGSUC% <= LOF(NX12%) / 32 Then
      Get #NX12%, RGSUC%, E20$
      DETAPAS$ = E20$
   End If
   '
   If MODO% = 2 Then
      LI% = 0: LS% = MAXENCDIARIO% + 1
11    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then GoTo 29 ' error de encabezado
      L% = LI% + E%
      If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 29 ' error de encabezado
      Call LEENCDIARIO(L%)
      If CVS(E1$) < NROASIEN - 0.01 Then LI% = L%: GoTo 11
      If CVS(E1$) > NROASIEN + 0.01 Then LS% = L%: GoTo 11
      REGENCA% = L%
      '
      Call LEENCDIARIO(REGENCA%)
      REFASIEN$ = E7$
   End If
   '
   Exit Sub
   '
29 Call MENSERR(24, "Error de Cabecera de Asiento " + TRIM$(Str$(NROASIEN)))
   Call FINAL("")
   '
End Sub
'
Function MAXENCDIARIO%()
   Static UREMA1%, RUREMA1$
   Static UREMA2%, RUREMA2$
   '
   Call ABRENCDIARIO
   MAXENCDIARIO% = 0
   '
   If UREMA1% > 0 Then
      Call LEENCDIARIO(UREMA1%)
      If E0$ = RUREMA1$ Then
         If UREMA2% > 0 Then
            Call LEENCDIARIO(UREMA2%)
            If E0$ = RUREMA2$ Then
               MAXENCDIARIO% = UREMA1%
               Exit Function
            End If
         End If
      End If
   End If
   '
   For I% = LOF(NX9%) / 64 To 1 Step -1
      Call LEENCDIARIO(I%)
      If CVS(E1$) > 0 Then
         MAXENCDIARIO% = I%
         UREMA1% = I%: RUREMA1$ = E0$
         UREMA2% = UREMA1% + 1
         Call LEENCDIARIO(UREMA2%)
         RUREMA2$ = E0$
         Exit Function
      End If
   Next I%
End Function
'
Function MAXDIARIO%()
   Call ABREDIARIO
   MAXDIARIO% = 0
   For I% = LOF(NX10%) / 32 To 1 Step -1
      Call LEEDIARIO(I%)
      If CVI(E11$) > 0 Then
         MAXDIARIO% = I%
         Exit Function
      End If
   Next I%
End Function
'
Function MAXENC%()
   Call ABRENCAJA
   MAXENC% = 0
   For I% = LOF(NXK5%) / 64 To 1 Step -1
      Call LEENCAJA(I%)
      If CVS(D1$) > 0.01 Then
         MAXENC% = I%
         Exit Function
      End If
   Next I%
End Function
'
Function MAXDIA&()
'OK
   Call ABRESUCAJA
   MAXDIA& = 0
   For I& = LOF(NX6%) / 64 To 1 Step -1
      Call LEESUCAJA(I&)
      If CVI(D11$) > 0 Then
         MAXDIA& = I&
         Exit Function
      End If
   Next I&
   
End Function
'
Function ULTAS()
   ULTAS = 0
   If MAXENC% > 0 Then
      Call LEENCAJA(MAXENC%)
      ULTAS = CVS(D1$)
   End If
End Function
'
Function ULTASDIA()
   ULTASDIA = 0
   If MAXENCDIARIO% > 0 Then
      Call LEENCDIARIO(MAXENCDIARIO%)
      ULTASDIA = CVS(E1$)
   End If
End Function
'
Function FULTAS%()
   FULTAS% = COMEJ%
   If MAXENC% > 0 Then
      Call LEENCAJA(MAXENC%)
      FULTAS% = CVI(D3$)
   End If
End Function
'
Function FULTASDIA%()
   '
   SUCUI% = 0
   LIMINF = CDbl(CVDAT(COMEJ%))
   LIMSUP = CDbl(CVDAT(FINEJ%))
   SQLX$ = "SELECT * FROM EncabeAsien "
   SQLX$ = SQLX$ + "WHERE ("
   SQLX$ = SQLX$ + "StatAsien > 0 "
   SQLX$ = SQLX$ + "AND SucuAsien = " & SUCUI% & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) >= " & LIMINF & " "
   SQLX$ = SQLX$ + "AND CDBL(FECHASIEN) <= " & LIMSUP & " "
   SQLX$ = SQLX$ + ") "
   SQLX$ = SQLX$ + "ORDER by FECHASIEN DESC;"
   '
   SUCUI% = 0
   Set SELECASIEN = Contable.OpenRecordset(SQLX$, dbOpenDynaset, dbReadOnly)
   On Error Resume Next
   SELECASIEN.MoveFirst
   If Err Then
       FULTASDIA% = COMEJ%
     Else
       FULAS$ = Format(SELECASIEN!FECHASIEN, "DD/MM/YY")
       FULTASDIA% = FECHANUM(FULAS$)
   End If
   On Error GoTo 0
   If FULTASDIA% < COMEJ% Then FULTASDIA% = COMEJ%
   If FULTASDIA% > FINEJ% Then FULTASDIA% = FINEJ%
   '
End Function
'
Function FEPROAS%()
   FEPROAS% = FULTAS%
   If FEPROAS% > FINEJ% Then
      MERRO$ = "Imposible Continuar Ejecución.\  \"
      MERRO$ = MERRO$ + "Aparecen Asientos con Fecha POSTERIOR\"
      MERRO$ = MERRO$ + "al Cierre del Ejercicio Activo.\  \"
      MERRO$ = MERRO$ + "Consulte Urgente al Soporte de Sistemas FLEXOFT."
      Call MENSERR(24, MERRO$)
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
End Function

Function FEPROASDIA%()
   FEPROASDIA% = FULTASDIA%
Exit Function
   If FEPROASDIA% > FINEJ% Then
      MERRO$ = "Imposible Continuar Ejecución.\  \"
      MERRO$ = MERRO$ + "Aparecen Asientos con Fecha POSTERIOR\"
      MERRO$ = MERRO$ + "al Cierre del Ejercicio Activo.\  \"
      MERRO$ = MERRO$ + "Consulte Urgente al Soporte de Sistemas FLEXOFT."
      Call MENSERR(24, MERRO$)
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
End Function

Function COLMES%(FF%)
   '
   FEX = COMEJ%
   CMJ$ = FECHATEX$(FEX)
   '
   FEX = FF%
   FFT$ = FECHATEX$(FEX)
   COLMES% = XVALO(Mid$(FFT$, 4, 2)) - XVALO(Mid$(CMJ$, 4, 2)) + 2
   If COLMES% < 2 Then COLMES% = COLMES% + 12
   '
End Function
'
Function PUNMOCAJA$(CUE%)
   '
   SCA = Screen.MousePointer
   Screen.MousePointer = 11
   CUEBIN$ = MKI$(CUE%)
   PMC$ = ""
   Call ABRESUCAJA
   Call ABREICAJA
   '
   FINL& = LOF(NX7%) / 128
   For KX% = 1 To LOF(NX7%) / 128
      KXL& = KX%
      Call TRACE(20, KXL&, FINL&)
      Get #NX7%, KX%, DX7$
      REG$ = DX7$: OF% = 0
10    POSI% = InStr(REG$, CUEBIN$)
      If POSI% = 0 Then GoTo 19
      If (POSI% Mod 2) = 0 Then
         REG$ = Mid$(REG$, POSI% + 1)
         OF% = OF% + POSI% / 2
         GoTo 10
      End If
      REGI& = 64 * (KXL& - 1) + OF% + (POSI% + 1) / 2
      If REGI& > MAXDIA& Then
         Call MENSERR(24, "Error de Indexación Caja y Bancos")
         Call FINAL("")
      End If
      '
15    Call LEESUCAJA(REGI&)
      If CVI(D11$) <> CUE% Then
         Call INDEXCAJA(REGI&)
         GoTo 15
      End If
      '
      PMC$ = PMC$ + MKS$(REGI&)
      '
      REG$ = Mid$(REG$, POSI% + 2)
      OF% = OF% + (POSI% + 1) / 2
      GoTo 10
      '
19 Next KX%
   '
   PUNMOCAJA$ = PMC$
   Screen.MousePointer = SCA
   '
End Function
'
Function PUNMODIAR$(CUE%)
   '
   CUEBIN$ = MKI$(CUE%)
   PMC$ = ""
   Call ABREIDIAR
   '
   For KX% = 1 To LOF(NX11%) / 128
      Get #NX11%, KX%, EX$
      REG$ = EX$: OF% = 0
10    POSI% = InStr(REG$, CUEBIN$)
      If POSI% = 0 Then GoTo 19
      If (POSI% Mod 2) = 0 Then
         REG$ = Mid$(REG$, POSI% + 1)
         OF% = OF% + POSI% / 2
         GoTo 10
      End If
      REGI% = 64 * (KX% - 1) + OF% + (POSI% + 1) / 2
      If REGI% > MAXDIARIO% Then
         Call MENSERR(24, "Error de Indexación Contabilidad General")
         Call INDEXDIARIO(1)
         Call FINAL("")
      End If
      '
      Call LEEDIARIO(REGI%)
      If CVI(E11$) <> CUE% Then
         Call MENSERR(24, "Error de Indexación Contabilidad General")
         Call INDEXDIARIO(1)
         Call FINAL("")
      End If
      '
      PMC$ = PMC$ + MKI$(REGI%)
      '
      REG$ = Mid$(REG$, POSI% + 2)
      OF% = OF% + (POSI% + 1) / 2
      GoTo 10
      '
19 Next KX%
   '
   PUNMODIAR$ = PMC$
   '
End Function
'
Sub INDEXCAJA(PRIREG&)
   '
   SMPA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   If NX7% > 0 Then
      Close #NX7%: NX7% = 0
   End If
   '
   If PRIREG& < 1 Then PRIREG& = 1
   '
   NX7% = FILEOPEN%(LU$ + "ISUCAJA" + SE$ + "." + EXTE$, 0, 2)
   Dim DXX As String * 2
   '
   PRXX& = PRIREG&
10 If PRXX& > 1 Then
      Get #NX7%, PRXX& - 1, DXX$
      Call LEESUCAJA(PRXX& - 1)
      If DXX$ <> D11$ Then
         PRXX& = PRXX& - 1
         GoTo 10
      End If
   End If
   '
   UIXX& = MAXDIA& + 64
   If LOF(NX7%) / 2 > UIXX& Then UIXX& = LOF(NX7%) / 2
   For I& = PRXX& To UIXX&
      Call LEESUCAJA(I&)
      DY$ = D11$
      LSet DXX$ = DY$
      Put #NX7%, I&, DXX$
   Next I&
   Close #NX7%: NX7% = 0
   Call ABREICAJA
   '
   Screen.MousePointer = SMPA
   '
End Sub
'
Sub INDEXDIARIO(PRIREG%)
   '
   SMPA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   If NX11% > 0 Then
      Close #NX11%: NX11% = 0
   End If
   '
   If PRIREG% < 1 Then PRIREG% = 1
   '
   NX11% = FILEOPEN%(LU$ + "IDIARIO" + SE$ + "." + EXTE$, 0, 2)
   Dim EXX As String * 2
   '
   PRXX% = PRIREG%
10 If PRXX% > 1 Then
      Get #NX11%, PRXX% - 1, EXX$
      Call LEEDIARIO(PRXX% - 1)
      If EXX$ <> E11$ Then
         PRXX% = PRXX% - 1
         GoTo 10
      End If
   End If
   '
   UIXX% = MAXDIARIO% + 64
   If LOF(NX11%) / 2 > UIXX% Then UIXX% = LOF(NX11%) / 2
   For I% = PRXX% To UIXX%
      Call LEEDIARIO(I%)
      EY$ = E11$
      LSet EXX$ = EY$
      Put #NX11%, I%, EXX$
   Next I%
   Close #NX11%: NX11% = 0
   Call ABREIDIAR
   '
   Screen.MousePointer = SMPA
   '
End Sub
'
Sub GRASICAJA(NROAS, FEAS%, REFE$, CONTA%, CTA%(), IMPA#(), DETA$())
   '
   Screen.MousePointer = 11
   '
   Call ABRENCAJA
   Call ABRESUCAJA
   Call ABRESALMEN
   Call ABRESALCAJ
   '
   If NROAS < 1 Or NROAS > 9999.9 Then
      Call MENSERR(24, "Imposible Grabar.\Número de Asiento Fuera de Rango.")
      Call FINAL("")
   End If
   '
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Then
      Call MENSERR(24, "Imposible Grabar.\Fecha de Asiento Fuera de Ejercicio Actual")
      Call FINAL("")
   End If
   '
   REGENCA% = 0: REGENC% = 0
   REGINIA& = 0: REGINI& = 0
   CONTANT% = 0
   FEASA% = 0
   '
   If NROAS <= ULTAS + 0.05 Then
      LI% = 0: LS% = MAXENC% + 1
81    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then GoTo 819 '    INTERCA% = 1: Exit Function
      L% = LI% + E%
      If L% < 1 Or L% > MAXENC% Then GoTo 819 ': INTERCA% = 1: Exit Function
      Call LEENCAJA(L%)
      If CVS(D1$) < NROAS - 0.01 Then LI% = L%: GoTo 81
      If CVS(D1$) > NROAS + 0.01 Then LS% = L%: GoTo 81
      REGENCA% = L%
      '
      Call LEENCAJA(REGENCA%)
      FEASA% = CVI(D3$)
      '
819   LIL& = 0: LSL& = MAXDIA& + 1
82    EL& = Int((LSL& - LIL&) / 2)
      If EL& < 1 Then GoTo 84
      LL& = LIL& + EL&
      If LL& < 1 Or LL& > MAXDIA& Then GoTo 84
      Call LEESUCAJA(LL&)
      If CVS(D12$) < NROAS - 0.01 Then LIL& = LL&: GoTo 82
      If CVS(D12$) > NROAS + 0.01 Then LSL& = LL&: GoTo 82
      '
      For K1& = LL& To 1 Step -1
         Call LEESUCAJA(K1&)
         If Abs(CVS(D12$) - NROAS) > 0.01 Then GoTo 83
         REGINIA& = K1&
      Next K1&
      '
83    For K1& = 1 To MAXDIA& - REGINIA& + 1
         Call LEESUCAJA(REGINIA& + K1& - 1)
         If Abs(CVS(D12$) - NROAS) > 0.01 Then GoTo 84
         CONTANT% = K1&
      Next K1&
   End If
   '
84 If CONTA% > 0 Then
      If REGENCA% > 0 Then
         REGENC% = REGENCA%
         GoTo 85
       ElseIf REGENCA% = 0 Then
         For RGX% = MAXENC% To 1 Step -1
            Call LEENCAJA(RGX%)
            If CVS(D1$) < NROAS - 0.01 Then
               REGENC% = RGX% + 1
               GoTo 85
            End If
            Put #NXK5%, RGX% + 1, D0$
         Next RGX%
         REGENC% = 1
      End If
    ElseIf CONTA% = 0 Then
      If REGENCA% > 0 Then
         For RGX% = REGENCA% + 1 To MAXENC% + 1
            Call LEENCAJA(RGX%)
            Put #NXK5%, RGX% - 1, D0$
         Next RGX%
      End If
      REGENC% = 0
   End If
   '
85 REGINI& = REGINIA&
   If REGINI& < 1 Then
      For RGE& = MAXDIA& To 1 Step -1
         Call LEESUCAJA(RGE&)
         If CVS(D12$) < NROAS - 0.01 Then
            REGINI& = RGE& + 1
            GoTo 86
         End If
      Next RGE&
      REGINI& = 1
   End If
   '
86 If REGENC% > 1 Then
      Call LEENCAJA(REGENC% - 1)
      If CVI(D3$) > FEAS% Then
         FEAS% = CVI(D3$)
         FEX = FEAS%
         Call MENSERR(24, "Imposible Grabar este Asiento\en la Fecha Elegida.\  \Se Grabará con Fecha " + FECHATEX$(FEX))
      End If
   End If
   '
87 If REGENC% > 0 Then
      If REGENC% < MAXENC% Then
         Call LEENCAJA(REGENC% + 1)
         If CVI(D3$) < FEAS% Then
            FEAS% = CVI(D3$)
            FEX = FEAS%
            Call MENSERR(24, "Imposible Grabar este Asiento\en la Fecha Elegida.\  \Se Grabará con Fecha " + FECHATEX$(FEX))
         End If
      End If
      '
      Get #NXK5%, REGENC%, D0$
      Call REPLA(D0$, MKS$(NROAS), 1, 4)
      Call REPLA(D0$, MKI$(0), 5, 2)
      Call REPLA(D0$, MKI$(FEAS%), 7, 2)
      Call REPLA(D0$, MKD$(SUMD#), 9, 8)
      Call REPLA(D0$, MKI$(0), 17, 2)
      Call REPLA(D0$, MKI$(CONTA%), 19, 2)
      Call REPLA(D0$, REFE$, 21, 40)
      Call REPLA(D0$, Chr$(TIASI%) + String$(3, 0), 61, 4)
      Put #NXK5%, REGENC%, D0$
   End If
   '
90 For I& = REGINIA& To REGINIA& + CONTANT% - 1
      Call LEESUCAJA(I&)
      REGCUE% = CVI(D11$)
      FFI% = CVI(D15$)
      IMPO# = (-1) * CVD(D16$)
      GoSub 189            ' AJUSTAR SALDOS MENSUALES
      If I& < PRIMODCAJA& Then PRIMODCAJA& = I&
   Next I&
   '
   If CONTA% < CONTANT% Then
      '
      DIFE% = CONTANT% - CONTA%
      For I& = REGINI& + CONTANT% To MAXDIA&
         Get #NX6%, I&, D10$
         Put #NX6%, I& - DIFE%, D10$
         If PRIMODCAJA& < 1 Or (I& - DIFE%) < PRIMODCAJA& Then
            PRIMODCAJA& = I& - DIFE%
         End If
      Next I&
      '
      For I& = MAXDIA& - DIFE% + 1 To LOF(NX6%) / 64
         LSet D10$ = String$(64, 0)
         Put #NX6%, I&, D10$
         If PRIMODCAJA& < 1 Or I& < PRIMODCAJA& Then
            PRIMODCAJA& = I&
         End If
      Next I&
      '
   ElseIf CONTA% > CONTANT% Then
      '
      DIFE% = CONTA% - CONTANT%
      For I& = MAXDIA& To REGINI& + CONTANT% Step -1
         Get #NX6%, I&, D10$
         Put #NX6%, I& + DIFE%, D10$
         If PRIMODCAJA& < 1 Or (I& + DIFE%) < PRIMODCAJA& Then
            PRIMODCAJA& = I& + DIFE%
         End If
      Next I&
      '
      For I& = MAXDIA& + DIFE% + 1 To LOF(NX6%) / 64
         LSet D10$ = String$(64, 0)
         Put #NX6%, I&, D10$
         If PRIMODCAJA& < 1 Or I& < PRIMODCAJA& Then
            PRIMODCAJA& = I&
         End If
         If DTDIA% = 1 Then
            'LSET D20$=SPACE$(32)
            'PUT #N10%,I%
         End If
      Next I&
      '
   End If
   '
   FFI% = FEAS%
   For IC% = 1 To CONTA%
      Get #NX6%, REGINI& + IC% - 1, D10$
      Call REPLA(D10$, MKI$(CTA%(IC%)), 1, 2)
      Call REPLA(D10$, MKS$(NROAS), 3, 4)
      Call REPLA(D10$, MKI$(0), 7, 2)
      Call REPLA(D10$, MKI$(0), 9, 2)
      Call REPLA(D10$, MKI$(FEAS%), 11, 2)
      Call REPLA(D10$, MKD$(IMPA#(IC%)), 13, 8)
      Call REPLA(D10$, MKI$(REGENC%), 21, 2)
      Call REPLA(D10$, DETA$(IC%), 23, 32)
      Call REPLA(D10$, String$(10, 0), 55, 10)
      Put #NX6%, REGINI& + IC% - 1, D10$
      If REGINI& + IC% - 1 < PRIMODCAJA& Then
         PRIMODCAJA& = REGINI& + IC% - 1
      End If
      '
      REGCUE% = CTA%(IC%)
      IMPO# = IMPA#(IC%)
      GoSub 189          ' ACTUALIZAR SALDOS MENSUALES
   Next IC%
   '
   Call INDEXCAJA(PRIMODCAJA&)
   PRIMODCAJA& = MAXDIA&
   '
   Close #NX4%: NX4% = 0
   Close #NXK5%: NXK5% = 0
   Close #NX6%: NX6% = 0
   Close #NX8%: NX8% = 0
   Close #NX10%: NX10% = 0
   '
   Screen.MousePointer = 1
   '
Exit Sub
'
'**********************************************************
'
189 If REGCUE% > LOF(NX8%) / 112 Then
      For J% = LOF(NX8%) / 112 + 1 To REGCUE% + 1
         LSet C19$ = String$(112, 0)
         Put #NX8%, J%, C19$
      Next J%
   End If
   '
   If REGCUE% > LOF(NX4%) / 80 Then
      For J% = LOF(NX4%) / 80 + 1 To REGCUE% + 1
         LSet C20$ = String$(80, 0)
         Put #NX8%, J%, C20$
      Next J%
   End If
   '
   Get #NX8%, REGCUE%, C19$
   SALMEN$ = C19$: SALMENUE$ = ""
   For J% = 1 To 14
      SALME# = CVD(Mid$(SALMEN$, 8 * J% - 7))
      If J% >= COLMES%(FFI%) Then
         SALME# = SALME# + (Int(100 * IMPO# + 0.5)) / 100
      End If
      SALMENUE$ = SALMENUE$ + MKD$(SALME#)
   Next J%
   '
   LSet C19$ = SALMENUE$
   Put #NX8%, REGCUE%, C19$
   '
   COLSE% = 1 + 8 * XVALO(SE$): If COLSE% < 1 Or COLSE% > 73 Then COLSE% = 1
   Get #NX4%, REGCUE%, C20$
   CX$ = C20$
   SALME# = CVD(Mid$(CX$, COLSE%)) + (Int(100 * IMPO# + 0.5)) / 100
   Mid$(CX$, COLSE%, 8) = MKD$(SALME#)
   LSet C20$ = CX$
   Put #NX4%, REGCUE%, C20$
   '
Return
'
End Sub
'

Sub GRASIDIAR(NROAS, FEAS%, REFE$, CONTA%, CTA%(), IMPA#(), DETA$())
   '
   Screen.MousePointer = 11
   '
   Call ABRENCDIARIO
   Call ABREDIARIO
   Call ABRESALMEN
   '
   If NROAS < 1 Or NROAS > 9999.9 Then
      Call MENSERR(24, "Imposible Grabar.\Número de Asiento Fuera de Rango.")
      Call FINAL("")
   End If
   '
   If FEAS% < COMEJ% Or FEAS% > FINEJ% Then
      Call MENSERR(24, "Imposible Grabar.\Fecha de Asiento Fuera de Ejercicio Actual")
      Call FINAL("")
   End If
   '
   REGENCA% = 0: REGENC% = 0
   REGINIA% = 0: REGINI% = 0
   CONTANT% = 0
   FEASA% = 0
   '
   If NROAS <= ULTASDIA + 0.05 Then
      LI% = 0: LS% = MAXENCDIARIO% + 1
81    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then GoTo 819 '    INTERCA% = 1: Exit Function
      L% = LI% + E%
      If L% < 1 Or L% > MAXENCDIARIO% Then GoTo 819 ': INTERCA% = 1: Exit Function
      Call LEENCDIARIO(L%)
      If CVS(E1$) < NROAS - 0.01 Then LI% = L%: GoTo 81
      If CVS(E1$) > NROAS + 0.01 Then LS% = L%: GoTo 81
      REGENCA% = L%
      '
      Call LEENCDIARIO(REGENCA%)
      FEASA% = CVI(E3$)
      '
819   LI% = 0: LS% = MAXDIARIO% + 1
82    E% = Int((LS% - LI%) / 2)
      If E% < 1 Then GoTo 84
      L% = LI% + E%
      If L% < 1 Or L% > MAXDIARIO% Then GoTo 84
      Call LEEDIARIO(L%)
      If CVS(E12$) < NROAS - 0.01 Then LI% = L%: GoTo 82
      If CVS(E12$) > NROAS + 0.01 Then LS% = L%: GoTo 82
      '
      For K1% = L% To 1 Step -1
         Call LEEDIARIO(K1%)
         If Abs(CVS(E12$) - NROAS) > 0.01 Then GoTo 83
         REGINIA% = K1%
      Next K1%
      '
83    For K1% = 1 To MAXDIARIO% - REGINIA% + 1
         Call LEEDIARIO(REGINIA% + K1% - 1)
         If Abs(CVS(E12$) - NROAS) > 0.01 Then GoTo 84
         CONTANT% = K1%
      Next K1%
   End If
   '
84 If CONTA% > 0 Then
      If REGENCA% > 0 Then
         REGENC% = REGENCA%
         GoTo 85
       ElseIf REGENCA% = 0 Then
         For RGX% = MAXENCDIARIO% To 1 Step -1
            Call LEENCDIARIO(RGX%)
            If CVS(E1$) < NROAS - 0.01 Then
               REGENC% = RGX% + 1
               GoTo 85
            End If
            Put #NX9%, RGX% + 1, E0$
         Next RGX%
         REGENC% = 1
      End If
    ElseIf CONTA% = 0 Then
      If REGENCA% > 0 Then
         For RGX% = REGENCA% + 1 To MAXENC% + 1
            Call LEENCDIARIO(RGX%)
            Put #NX9%, RGX% - 1, E0$
         Next RGX%
      End If
      REGENC% = 0
   End If
   '
85 REGINI% = REGINIA%
   If REGINI% < 1 Then
      For RGE% = MAXDIARIO% To 1 Step -1
         Call LEEDIARIO(RGE%)
         If CVS(E12$) < NROAS - 0.01 Then
            REGINI% = RGE% + 1
            GoTo 86
         End If
      Next RGE%
      REGINI% = 1
   End If
   '
86 If REGENC% > 1 Then
      Call LEENCDIARIO(REGENC% - 1)
      If CVI(E3$) > FEAS% Then
         FEAS% = CVI(E3$)
         FEX = FEAS%
         Call MENSERR(24, "Imposible Grabar este Asiento\en la Fecha Elegida.\  \Se Grabará con Fecha " + FECHATEX$(FEX))
      End If
   End If
   '
87 If REGENC% > 0 Then
      If REGENC% < MAXENCDIARIO% Then
         Call LEENCDIARIO(REGENC% + 1)
         If CVI(E3$) < FEAS% Then
            FEAS% = CVI(E3$)
            FEX = FEAS%
            Call MENSERR(24, "Imposible Grabar este Asiento\en la Fecha Elegida.\  \Se Grabará con Fecha " + FECHATEX$(FEX))
         End If
      End If
      '
      Get #NX9%, REGENC%, E0$
      Call REPLA(E0$, MKS$(NROAS), 1, 4)
      Call REPLA(E0$, MKI$(0), 5, 2)
      Call REPLA(E0$, MKI$(FEAS%), 7, 2)
      Call REPLA(E0$, MKD$(SUMD#), 9, 8)
      Call REPLA(E0$, MKI$(0), 17, 2)
      Call REPLA(E0$, MKI$(CONTA%), 19, 2)
      Call REPLA(E0$, REFE$, 21, 40)
      Call REPLA(E0$, Chr$(TIASI%) + String$(3, 0), 61, 4)
      Put #NX9%, REGENC%, E0$
   End If
   '
90 For I% = REGINIA% To REGINIA% + CONTANT% - 1
      Call LEEDIARIO(I%)
      REGCUE% = CVI(E11$)
      FFI% = CVI(E15$)
      IMPO# = (-1) * CVD(E16$)
      GoSub 189            ' AJUSTAR SALDOS MENSUALES
      If I% < PRIMODIAR% Then PRIMODIAR% = I%
   Next I%
   '
   If CONTA% < CONTANT% Then
      '
      DIFE% = CONTANT% - CONTA%
      For I% = REGINI% + CONTANT% To MAXDIARIO%
         Get #NX10%, I%, E10$
         Put #NX10%, I% - DIFE%, E10$
         If PRIMODIAR% < 1 Or (I% - DIFE%) < PRIMODIAR% Then
            PRIMODIAR% = I% - DIFE%
         End If
         If LOF(NX12%) / 32 >= I% Then
            Get #NX12%, I%, E20$
            Put #NX12%, I% - DIFE%, E20$
         End If
      Next I%
      '
      For I% = MAXDIARIO% - DIFE% + 1 To LOF(NX10%) / 32
         LSet E10$ = String$(32, 0)
         Put #NX10%, I%, E10$
         If PRIMODIAR% < 1 Or I% < PRIMODIAR% Then
            PRIMODIAR% = I%
         End If
      Next I%
      INI% = MAXDIARIO% - DIFE%
      If INI% > LOF(NX12%) / 32 Then INI% = LOF(NX12%) / 32
      If INI% < 0 Then INI% = 0
      For I% = INI% + 1 To LOF(NX10%) / 32
         LSet E20$ = Space$(32)
         Put #NX12%, I%, E20$
      Next I%
      
      '
   ElseIf CONTA% > CONTANT% Then
      '
      DIFE% = CONTA% - CONTANT%
      For I% = MAXDIARIO% To REGINI% + CONTANT% Step -1
         Get #NX10%, I%, E10$
         Put #NX10%, I% + DIFE%, E10$
         If PRIMODIAR% < 1 Or (I% + DIFE%) < PRIMODIAR% Then
            PRIMODIAR% = I% + DIFE%
         End If
         If LOF(NX12%) / 32 >= I% Then
            Get #NX12%, I%, E20$
            Put #NX12%, I% + DIFE%, E20$
         End If
      Next I%
      '
      For I% = MAXDIARIO% + DIFE% + 1 To LOF(NX10%) / 32
         LSet E10$ = String$(32, 0)
         Put #NX10%, I%, E10$
         If PRIMODIAR% < 1 Or I% < PRIMODIAR% Then
            PRIMODIAR% = I%
         End If
      Next I%
      INI% = MAXDIARIO% + DIFE%
      If INI% > LOF(NX12%) / 32 Then INI% = LOF(NX12%) / 32
      For I% = INI% + 1 To LOF(NX10%) / 32
         LSet E20$ = Space$(32)
         Put #NX12%, I%, E20$
      Next I%
      '
   End If
   '
   FFI% = FEAS%
   For I% = 1 To CONTA%
      Get #NX10%, REGINI% + I% - 1, E10$
      Get #NX12%, REGINI% + I% - 1, E20$
      Call REPLA(E10$, MKI$(CTA%(I%)), 1, 2)
      Call REPLA(E10$, MKS$(NROAS), 3, 4)
      Call REPLA(E10$, MKI$(0), 7, 2)
      Call REPLA(E10$, MKI$(0), 9, 2)
      Call REPLA(E10$, MKI$(FEAS%), 11, 2)
      Call REPLA(E10$, MKD$(IMPA#(I%)), 13, 8)
      Call REPLA(E10$, MKI$(REGENC%), 21, 2)
      Call REPLA(E20$, DETA$(I%), 1, 32)
      Call REPLA(E10$, String$(10, 0), 23, 10)
      If TIASI% = 3 Then
         Call REPLA(E10$, Chr$(1), 23, 1)
      End If
      Put #NX10%, REGINI% + I% - 1, E10$
      Put #NX12%, REGINI% + I% - 1, E20$
      If REGINI% + I% - 1 < PRIMODIAR% Then
         PRIMODIAR% = REGINI% + I% - 1
      End If
      '
      REGCUE% = CTA%(I%)
      IMPO# = IMPA#(I%)
      GoSub 189          ' ACTUALIZAR SALDOS MENSUALES
   Next I%
   '
   Call INDEXDIARIO(PRIMODIAR%)
   PRIMODIAR% = MAXDIARIO%
   '
   Close #NX8%: NX8% = 0
   Close #NX9%: NX9% = 0
   Close #NX10%: NX10% = 0
   Close #NX12%: NX12% = 0
   '
   Screen.MousePointer = 1
   '
Exit Sub
'
'**********************************************************
'
189 If REGCUE% > LOF(NX8%) / 112 Then
      For J% = LOF(NX8%) / 112 + 1 To REGCUE% + 1
         LSet C19$ = String$(112, 0)
         Put #NX8%, J%, C19$
      Next J%
   End If
   '
   Get #NX8%, REGCUE%, C19$
   SALMEN$ = C19$: SALMENUE$ = ""
   For J% = 1 To 14
      SALME# = CVD(Mid$(SALMEN$, 8 * J% - 7))
      If J% >= COLMES%(FFI%) Then
         SALME# = SALME# + (Int(100 * IMPO# + 0.5)) / 100
      End If
      SALMENUE$ = SALMENUE$ + MKD$(SALME#)
   Next J%
   '
   LSet C19$ = SALMENUE$
   Put #NX8%, REGCUE%, C19$
   '
Return
'
End Sub
'
Sub CARBACHEQ()
   '
   Static CAMP%, CUEMP%(2048), MEPA%(2048), DEVAL$(2048)
   '
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUEINT7%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICHEQ% < 1 Or CUEICHEQ% > NUCUEN% Then
      Call MENSERR(24, "Falta Definir Cuenta Contable\Asociada a Cartera de Valores")
      OPASICONT.Show 1
      GoTo 99
   End If
   '
   If CAMP% <= 0 Then
      For I& = 1 To ULTREG&("MECOBRA")
         F0$ = REGLEIDO$("MECOBRA", I&)
         CAMP% = I&
         MEPA%(CAMP%) = CVI(Left$(F0$, 2))
         DEVAL$(CAMP%) = Mid$(F0$, 3, 12)
         CUE$ = Mid$(F0$, 45, 12)
         CUEMP%(CAMP%) = CUEINT7%(CUE$)
         If TRIM$(CODCUE$(CUEMP%(CAMP%))) <> TRIM$(CUE$) Then
            Call MENSERR(24, "Cuenta Contable Incorrecta\para Medio de Cobranza '" + TRIM$(Mid$(F0$, 3, 12)) + "'\-\Resultados Imprevisibles.")
            Call FINAL("")
         End If
     Next I&
   End If
   '
   FICHECA& = ULTREG&("CHECAR")
   FBC& = ULTREG&("BADACHE")
   For U& = 1 To ASICAJA.List1.ListCount
      Z$ = ASICAJA.List1.List(U& - 1)
      RECHECA& = XVALO(Mid$(Z$, 121, 8))
      If RECHECA& > 0 Then
         If RECHECA& <= FICHECA& Then
            X$ = REGLEIDO$("CHECAR", RECHECA&)
            RBC& = CVS(Mid$(X$, 51, 4))
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FERVI") > 0 Then
   If RBC& < 0 Then
      RBC& = 0
   End If
End If
            If RBC& <> (-1) Then
               If RBC& < 1 Or RBC& > FBC& Then
                  QUE% = CVI(Left$(X$, 2))
                  BANCO$ = Mid$(X$, 17, 16)
                  '
                  CARBADACHE% = 0
                  For UX% = 1 To CAMP%
                     If MEPA%(UX%) = QUE% Then
                        BANCO$ = TRIM$(DEVAL$(UX%))
                        If CUEMP%(UX%) = CUEICHEQ% Then
                           CARBADACHE% = 1
                        End If
                        GoTo 18
                     End If
                  Next UX%
                  '
18                ISCHE% = 0
                  If UCase$(Left$(BANCO$, 3)) = "CH." Then
                     ISCHE% = 1
                     BANCO$ = Mid$(BANCO$, 4)
                  End If
                  BANCO$ = AJUSTI$(BANCO$, 16)
                  '
                  If CARBADACHE% = 1 Then
                     NUCHE$ = Mid$(X$, 5, 12)
                     If TRIM$(NUCHE$) <> "" Or ISCHE% > 0 Then
                        FEX = CVI(Mid$(X$, 41, 4))
                        FECHE$ = FECHATEX$(FEX)
                        IMCHE# = CVD(Mid$(X$, 33, 8))
                        FEMI% = CVI(Mid$(X$, 43, 2))
                        FEVE% = CVI(Mid$(X$, 41, 4))
                        FERE% = CVI(Mid$(X$, 45, 2))
                        DORE$ = "RB." + TRIM$(Str$(CVS(Mid$(X$, 47, 4))))
                        NC% = CVI(Mid$(X$, 3, 2))
                        ORI$ = RASOCLI$(NC%)
                        '
                        Y$ = REGBLAN$("BADACHE")
                        Call REPLA(Y$, BANCO$, 1, 16)
                        Call REPLA(Y$, NUCHE$, 19, 12)
                        Call REPLA(Y$, FECHE$, 33, 8)
                        Call REPLA(Y$, MKD$(IMCHE#), 49, 8)
                        Call REPLA(Y$, MKI$(FEMI%), 57, 2)
                        Call REPLA(Y$, MKI$(FEVE%), 59, 2)
                        Call REPLA(Y$, MKI$(FERE%), 61, 2)
                        Call REPLA(Y$, DORE$, 63, 10)
                        Call REPLA(Y$, ORI$, 73, 30)
                        '
                        Call REGAPP("BADACHE", Y$)
                        RBC& = ULTREG&("BADACHE")
                        Call REPLA(X$, MKS$(RBC&), 51, 4)
                        Call GRAREG("CHECAR", X$, RECHECA&)
                        '
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   If TRIM$(CONTROL$("", "CARBACHE")) = "" Then
      HOII% = HOY(HOS$)
      '
      FICHECA& = ULTREG&("CHECAR")
      FBC& = ULTREG&("BADACHE")
      For U& = 1 To FICHECA&
         Call TRACE(20, U&, FICHECA&)
         X$ = REGLEIDO$("CHECAR", U&)
         RBC& = CVS(Mid$(X$, 51, 4))
         If RBC& <> (-1) Then
            If RBC& < 1 Or RBC& > FBC& Then
               QUE% = CVI(Left$(X$, 2))
               FEVE% = CVI(Mid$(X$, 41, 4))
               If FEVE% < HOII% Then GoTo 89
               '
               CARBADACHE% = 0
               For UX% = 1 To CAMP%
                  If MEPA%(UX%) = QUE% Then
                     If CUEMP%(UX%) = CUEICHEQ% Then
                        BANCO$ = TRIM$(DEVAL$(UX%))
                        CARBADACHE% = 1
                     End If
                     GoTo 78
                  End If
               Next UX%
               '
78             If Left$(BANCO$, 3) = "CH." Then
                  BANCO$ = Mid$(BANCO$, 4)
               End If
               BANCO$ = AJUSTI$(BANCO$, 16)
               '
               If CARBADACHE% = 1 Then
                  NUCHE$ = Mid$(X$, 5, 12)
                  If TRIM$(NUCHE$) <> "" Then
                     FEX = CVI(Mid$(X$, 41, 4))
                     FECHE$ = FECHATEX$(FEX)
                     IMCHE# = CVD(Mid$(X$, 33, 8))
                     FEMI% = CVI(Mid$(X$, 43, 2))
                     FEVE% = CVI(Mid$(X$, 41, 4))
                     FERE% = CVI(Mid$(X$, 45, 2))
                     DORE$ = "RB." + TRIM$(Str$(CVS(Mid$(X$, 47, 4))))
                     NC% = CVI(Mid$(X$, 3, 2))
                     ORI$ = RASOCLI$(NC%)
                     '
                     Y$ = REGBLAN$("BADACHE")
                     Call REPLA(Y$, BANCO$, 1, 16)
                     Call REPLA(Y$, NUCHE$, 19, 12)
                     Call REPLA(Y$, FECHE$, 33, 8)
                     Call REPLA(Y$, MKD$(IMCHE#), 49, 8)
                     Call REPLA(Y$, MKI$(FEMI%), 57, 2)
                     Call REPLA(Y$, MKI$(FEVE%), 59, 2)
                     Call REPLA(Y$, MKI$(FERE%), 61, 2)
                     Call REPLA(Y$, DORE$, 63, 10)
                     Call REPLA(Y$, ORI$, 73, 30)
                     '
                     Call REGAPP("BADACHE", Y$)
                     RBC& = ULTREG&("BADACHE")
                     Call REPLA(X$, MKS$(RBC&), 51, 4)
                     Call GRAREG("CHECAR", X$, U&)
                     '
                  End If
               End If
               '
               Call REPLA(X$, MKS$(-1), 51, 4)
               Call GRAREG("CHECAR", X$, U&)
               '
            End If
         End If
         '
89    Next U&
      Call GRACONTROL("", "CARBACHE", HOS$)
   End If
   '
99 Call CIERRARCH("CHECAR")
   Call CIERRARCH("BADACHE")
   '
End Sub

Sub BLABACHEQ()
   '
   FBC& = ULTREG&("BADACHE")
   FF% = FECHANUM(ASICAJA.Text1.TEXT)
   NROAS = XVALO(ASICAJA.Text2.TEXT)
   DENOBAN$ = TRIM$(Left$(ASICAJA.List1.List(0), 12))
   CUEBA$ = DENOBAN$
   CUII% = CUEINT7%(CUEBA$)
   DENBA$ = DENOCUE$(CUII%)
   If TRIM$(DENBA$) <> "" Then DENOBAN$ = TRIM$(DENBA$)
   If Left$(UCase$(DENOBAN$), 6) = "BANCO " Then
      DENOBAN$ = "B." + Mid$(DENOBAN$, 7)
   End If
   REFE$ = TRIM$(Mid$(ASICAJA.List1.List(0), 14, 27))
   '
   HOII% = HOY(HOS$)
   For U& = 2 To ASICAJA.List1.ListCount
      Z$ = ASICAJA.List1.List(U& - 1)
      REGBA& = XVALO(Mid$(Z$, 121, 8))
      If REGBA& > 0 Then
         If REGBA& <= FBC& Then
            TTXX$ = REGLEIDO$("BADACHE", REGBA&)
            Call REPLA(TTXX$, REFE$ + " " + DENOBAN$, 103, 30)
            Call REPLA(TTXX$, MKI$(FF%), 133, 2)
            Call REPLA(TTXX$, "Asiento Contable C&B " + Str$(NROAS), 135, 40)
            Call REPLA(TTXX$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 175, 40)
            Call REPLA(TTXX$, "Usuario: " + TRIM$(USERNAME$), 215, 40)
            Call GRAREG("BADACHE", TTXX$, REGBA&)
         End If
      End If
   Next U&
   '
   Call CIERRARCH("BADACHE")
   '
End Sub
'
Sub RECARCHE(FF%, NUMEASIEN)
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BADACHE")
     TTXX$ = REGLEIDO$("BADACHE", U&)
     If CVI(Mid$(TTXX$, 133, 2)) = FF% Then
       If Abs(XVALO(Mid$(TTXX$, 155, 10)) - NUMEASIEN) < 0.05 Then
          Call REPLA(TTXX$, String$(152, 0), 103, 152)
          Call GRAREG("BADACHE", TTXX$, U&)
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   Call CIERRARCH("BADACHE")
   '
End Sub
'

Sub CARCUEBANC()
   '
   JJ& = 0
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("CUEBANC")
      XX$ = REGLEIDO$("CUEBANC", U&)
      CXI% = CUEINT7%(TRIM$(Left$(XX$, 12)))
      If CXI% > 0 Then
        JJ& = JJ& + 1
        Call GRAREG("!CUEBANC", XX$, JJ&)
      End If
   Next U&
   Call SETULTREG("!CUEBANC", JJ&)
   Call CIERRARCH("!CUEBANC")
   Call CIERRARCH("CUEBANC")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("CUEBANC")
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To ULTREG&("!CUEBANC")
         XX$ = REGLEIDO$("!CUEBANC", U&)
         CXI% = CUEINT7%(TRIM$(Left$(XX$, 12)))
         If CXI% > 0 Then
           JJ& = JJ& + 1
           Call GRAREG("CUEBANC", XX$, JJ&)
         End If
      Next U&
      Call SETULTREG("CUEBANC", JJ&)
      Call CIERRARCH("CUEBANC")
      Call FRESHECHO("CUEBANC")
      Call CIERRARCH("!CUEBANC")
      Screen.MousePointer = 1
   End If
   '
End Sub

Sub CARCUEINGAS()
    '
    If App.EXEName = "CONTAB04" Then
        FPLANCUEN.Show 1
        Exit Sub
      ElseIf EXISTE%("CONTAB04.EXE") > 0 Then
        Call CONECRUN("CONTAB04/PLANCUEN", "Contabilidad General")
        Exit Sub
    End If
    '
    Dim FCUEN As String * 64
    Dim FDEN As String * 128
    Dim ICUEN As String * 16
    Dim CODCUECO$(4096), RECUEN&(4096)
    UCUEN& = 0
    REMAX& = 0
    XMAX% = 0
    '
    EJERCI$ = EJACT$
    LU$ = LUDAT$
    ACUEN% = FILEOPEN%(LU$ + "CUENTAS." + EXTE$, 0, 64)
    ADEN% = FILEOPEN%(LU$ + "DENOCUE." + EXTE$, 0, 128)
    NX3% = FILEOPEN%(LUDAT$ + "INVECUE" + "." + EXTE$, 0, 16)
    '
    Call BLANARCH("!CUEINGAS")
    '
    Screen.MousePointer = 11
    For UI% = LOF(ACUEN%) / 64 To 1 Step -1
       Get #ACUEN%, UI%, FCUEN$
       If Asc(Left$(FCUEN$, 1)) > 0 Then
          NUCUEN% = UI%
          GoTo 5
       End If
    Next UI%
    NUCUEN% = 0
    '
5   XMAX% = 0
    FIN& = NUCUEN%
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       Get #ACUEN%, I&, FCUEN$
       X$ = FCUEN$
       If Left$(X$, 2) = "X_" Then
          XMAX% = XVALO(Mid$(X$, 3, 4))
       End If
    Next I&
    '
    FORMALTER.LINOSORT.Clear
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = "A" + Space$(32)
       Get #ACUEN%, U&, FCUEN$
       Call REPLA(X$, Left$(FCUEN$, 12), 2, 12)
       Call REPLA(X$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 15, 6)
       If TRIM$(Left$(FCUEN$, 12)) <> "" Then
          Call AGRELISCUEN(X$)
       End If
    Next U&
    '
    FIN& = FORMALTER.LINOSORT.ListCount
    For I1& = 1 To FORMALTER.LINOSORT.ListCount
       Call TRACE(20, I1&, FIN&)
       FORMALTER.LINOSORT.ListIndex = I1& - 1
       I& = XVALO(Mid$(FORMALTER.LINOSORT.TEXT, 15, 6))
       Get #ACUEN%, I&, FCUEN$
       X$ = FCUEN$
       TICUE$ = UCase$(Mid$(X$, 13, 1))
       If TICUE$ <> "I" Then
         If TICUE$ <> "G" Then
           If TICUE$ <> "H" Then
             TICUE$ = "D"
           End If
         End If
       End If
       CODCC$ = AJUSTI$(Left$(X$, 12), 12)
       CUEI% = CUEINT7%(CODCC$)
       DENCUE$ = DENOCUE$(CUEI%)
       If TRIM$(DENCUE$) = "" Then
         If CUEI% > 0 Then
           Get #ADEN%, CUEI%, FDEN$
           DENCUE$ = Mid$(FDEN$, 101, 28)
         End If
       End If
       '
       If TRIM$(CODCC$) = "" Then
         XMAX% = XMAX% + 1
         CODCC$ = "X_" + TRIM$(Str$(XMAX%))
       End If
       '
       Y$ = REGBLAN$("CUEINGAS")
       Call REPLA(Y$, CODCC$, 1, 12)
       Call REPLA(Y$, DENCUE$, 13, 28)
       Call REPLA(Y$, TICUE$, 41, 2)
       Call REPLA(Y$, MKS$(I&), 45, 4)
       Call REGAPP("!CUEINGAS", Y$)
       If I& > REMAX& Then
          REMAX& = I&
       End If
    Next I1&
    Call CIERRARCH("!CUEINGAS")
    '
    VTB% = VENTABU%("CUEINGAS")
    If VTB% < 0 Then Exit Sub
    '
    Screen.MousePointer = 11
    '
    Dim YAPLAN%(32767, 2)
    Dim IVP As String * 4
    NOHAYPLAN% = 0
    EXTEPLAN$ = "DAT": IPLAN% = 1: LPLAN$ = "Actual"
    '
6   NXP% = FILEOPEN%(LUDAT$ + "PLAGRAL." + EXTEPLAN$, 0, 4)
    If LOF(NXP%) < 4 Then NOHAYPLAN% = NOHAYPLAN% + 1: GoTo 7
    Get #NXP%, 1, IVP$
    If IVP$ = String$(4, 0) Then NOHAYPLAN% = NOHAYPLAN% + 1
    '
7   For U& = 1 To LOF(NXP%) / 4
       Get #NXP%, U&, IVP$
       ICUE% = CVI(Left$(IVP$, 2))
       If ICUE% > 0 Then
          If ICUE% <= 32767 Then
             YAPLAN%(ICUE%, IPLAN%) = YAPLAN%(ICUE%, IPLAN%) + 1
          End If
       End If
    Next U&
    Close #NXP%: NXP% = 0
    '
    For KKI% = 1 To NUCUEN%
       If YAPLAN%(KKI%, IPLAN%) > 1 Then
          Call MENSERR(24, "Cuenta '" + TRIM$(CODCUE$(KKI%)) + "'\Duplicada en el Plan de Cuentas Ejercicio " + LPLAN$ + ".")
       End If
    Next KKI%
    '
    If EXTEPLAN$ = "DAT" Then
       EXTEPLAN$ = "OLD"
       IPLAN% = 2
       LPLAN$ = "Anterior"
       GoTo 6
    End If
    '
    If EXISTE%("CONTAGEN.EXE") < 1 Then NOHAYPLAN% = 2
    '
    FIN& = ULTREG&("!CUEINGAS")
    REMAX& = NUCUEN%: UCUEN& = 0
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("!CUEINGAS", U&)
       CODCC$ = AJUSTI$(Left$(X$, 12), 12)
       '
       If TRIM$(CODCC$) = "" Then
         XMAX% = XMAX% + 1
         CODCC$ = AJUSTI$("X_" + TRIM$(Str$(XMAX%)), 12)
         GoTo 10
       End If
       '
       For KKJ& = 1 To UCUEN&
         If CODCUECO$(KKJ&) = CODCC$ Then
           XMAX% = XMAX% + 1
           CODCC$ = AJUSTI$("X_" + TRIM$(Str$(XMAX%)), 12)
           GoTo 10
         End If
       Next KKJ&
       '
10     If TRIM$(CODCC$) <> TRIM$(Left$(X$, 12)) Then
          Call MENSERR(24, "Codigo de Cuenta '" + TRIM$(Left$(X$, 12)) + "' no Válido\o Repetido. Recodificado como '" + TRIM$(CODCC$) + "'.")
       End If
       '
       RECU& = CVS(Mid$(X$, 45, 4))
       If RECU& < 1 Or RECU& > REMAX& Then
         REMAX& = REMAX& + 1
         RECU& = REMAX&
         GoTo 20
       End If
       '
       For KKJ& = 1 To UCUEN&
         If RECU& = RECUEN&(KKJ&) Then
           REMAX& = REMAX& + 1
           RECU& = REMAX&
           GoTo 20
         End If
       Next KKJ&
       '
20     UCUEN& = UCUEN& + 1
       CODCUECO$(UCUEN&) = CODCC$
       RECUEN&(UCUEN&) = RECU&
       '
       DENCUE$ = Mid$(X$, 13, 28)
       TICUE$ = UCase$(Mid$(X$, 41, 1))
       If TICUE$ <> "I" Then
         If TICUE$ <> "G" Then
           If TICUE$ <> "H" Then
             TICUE$ = "D"
           End If
         End If
       End If
       '
       Call REPLA(X$, CODCC$, 1, 12)
       Call REPLA(X$, TICUE$, 41, 1)
       Call REPLA(X$, MKS$(RECU&), 45, 4)
       Call GRAREG("!CUEINGAS", X$, U&)
       '
    Next U&
    '
    For K& = NUCUEN% + 1 To REMAX&
30     RECUENS$ = Space$(13) + String$(51, 0)
       REDEN$ = Space$(128)
       '
       LSet FCUEN$ = RECUENS$
       Put #ACUEN%, K&, FCUEN$
       '
       LSet FDEN$ = Space$(128)
       Put #ADEN%, K&, FDEN$
    Next K&
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("!CUEINGAS", U&)
       CODCC$ = AJUSTI$(Left$(X$, 12), 12)
       DENCUE$ = Mid$(X$, 13, 28)
       TICUE$ = UCase$(Mid$(X$, 41, 1))
       RECU& = CVS(Mid$(X$, 45, 4))
       '
       If RECU& > 0 Then
         If RECU& <= REMAX& Then
            '
            Get #ACUEN%, RECU&, FCUEN$
            RECUENS$ = FCUEN$
            Get #ADEN%, RECU&, FDEN$
            REDEN$ = FDEN$
            '
            Call REPLA(RECUENS$, CODCC$, 1, 12)
            Call REPLA(RECUENS$, TICUE$, 13, 1)
            LSet FCUEN$ = RECUENS$
            Put #ACUEN%, RECU&, FCUEN$
            '
            Call REPLA(REDEN$, DENCUE$, 101, 28)
            If TRIM$(Left$(REDEN$, 100)) = "" Then
               Call REPLA(REDEN$, DENCUE$, 1, 28)
            End If
            LSet FDEN$ = REDEN$
            Put #ADEN%, RECU&, FDEN$
            '
         End If
       End If
    Next U&
    '
    For U& = REMAX& + 1 To LOF(ACUEN%) / 64
       LSet FCUEN$ = String$(64, 0)
       Put #ACUEN%, U&, FCUEN$
    Next U&
    If REMAX& > NUCUEN% Then NUCUEN% = REMAX&
    '
    FORMALTER.LINOSORT.Clear
    For U& = 1 To REMAX&
       X$ = "A" + Space$(32)
       Get #ACUEN%, U&, FCUEN$
       Call REPLA(X$, Left$(FCUEN$, 12), 2, 12)
       Call REPLA(X$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 15, 6)
       Call AGRELISCUEN(X$)
    Next U&
    '
    J& = 0: K1& = 0: K2& = 0: K3& = 0
    FIN& = FORMALTER.LINOSORT.ListCount
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       FORMALTER.LINOSORT.ListIndex = U& - 1
       X$ = FORMALTER.LINOSORT.TEXT
       CUE$ = Mid$(X$, 2, 12)
       ICUE% = XVALO(Mid$(X$, 15, 6))
       If TRIM$(CUE$) <> "" Then
          Z$ = String$(16, 0)
          Call REPLA(Z$, CUE$, 1, 12)
          Call REPLA(Z$, MKI$(ICUE%), 13, 2)
          J& = J& + 1
          LSet ICUEN$ = Z$
          Put #NX3%, J&, ICUEN$
          '
          Get #ACUEN%, ICUE%, FCUEN$
          Get #ADEN%, ICUE%, FDEN$
          Y1$ = FCUEN$
          Y2$ = FDEN$
          TICUE$ = UCase$(Mid$(Y1$, 13, 1))
          If EXTE$ = "DAT" Then
            If TICUE$ = "I" Or TICUE$ = "G" Then
               Get #ADEN%, ICUE%, FDEN$
               Y2$ = FDEN$
               Z1$ = REGBLAN$("CUECON")
               Call REPLA(Z1$, MKI$(ICUE%), 1, 2)
               Call REPLA(Z1$, Mid$(Y2$, 101, 28), 3, 30)
               Call REPLA(Z1$, CUE$, 33, 12)
               Call REPLA(Z1$, CUE$, 45, 12)
               K1& = K1& + 1
               Call GRAREG("CUECON", Z1$, K1&)
            End If
          End If
          '
          Z2$ = REGBLAN$("ECUECON")
          Call REPLA(Z2$, CUE$, 1, 12)
          Call REPLA(Z2$, Mid$(Y2$, 101, 28), 13, 28)
          Call REPLA(Z2$, MKI$(ICUE%), 41, 2)
          K3& = K3& + 1
          Call GRAREG("ECUECOT", Z2$, K3&)
          '
          If YAPLAN%(ICUE%, 1) Or YAPLAN%(ICUE%, 2) > 0 Or NOHAYPLAN% = 2 Then
             K2& = K2& + 1
             Call GRAREG("ECUECON", Z2$, K2&)
          End If
          '
       End If
    Next U&
    '
    If K1& > 0 Then Call SETULTREG&("CUECON", K1&)
    Call SETULTREG&("ECUECON", K2&)
    Call SETULTREG&("ECUECOT", K3&)
    Call CIERRARCH("CUECON")
    Call CIERRARCH("ECUECON")
    Call FRESHECHO("ECUECON")
    Call CIERRARCH("ECUECOT")
    Call FRESHECHO("ECUECOT")
    Call CIERRARCH("*.*")
    '
    Close #ACUEN%: ACUEN% = 0
    Close #ADEN%: ADEN% = 0
    Close #NX3%: NX3% = 0
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub AGRELISCUEN(RAGRE$)
   '
   ULIS& = FORMALTER.LINOSORT.ListCount
   If ULIS& < 1 Or Left$(RAGRE$, 17) > FORMALTER.LINOSORT.List(ULIS& - 1) Then
       FORMALTER.LINOSORT.AddItem RAGRE$
       GoTo 10
     Else
       For J& = ULIS& - 1 To 0 Step -1
         If Left$(FORMALTER.LINOSORT.List(J&), 17) <= Left$(RAGRE$, 17) Then
           FORMALTER.LINOSORT.AddItem RAGRE$, J& + 1
           GoTo 10
         End If
       Next J&
       FORMALTER.LINOSORT.AddItem RAGRE$, 0
10 End If
   '
End Sub
'
'\\\OT-06-0244-WAR-28/06/06///
'Valida si el Proyecto se Encuentra en Vigencia
'1= PROYECTO VIGENTE
' 2= PROYECTO NO VIGENTE
Function VALIPROY%(NUPROY$, FF1%)
  '
  If ULTREG&("TACENCOS") < 1 Or TRIM$(NUPROY$) = "*" Then
     VALIPROY% = 1
     Exit Function
  End If
  '
  XZ$ = RECFILT$("TACENCOS", 1, NUPROY$)
  NR& = CVS(XZ$)
  YY$ = REGLEIDO$("TACENCOS", NR&)
  '
  FD1% = CVI(Mid$(YY$, 45, 2))
  FH1% = CVI(Mid$(YY$, 47, 2))
  If FH1% = 8224 Then ' Solo tiene Fecha Desde Vigencia
    If FF1% >= FD1% Then
      VALIPROY% = 1
      Exit Function
    End If
  End If
  '
  If FF1% >= FD1% Then
    If FF1% <= FH1% Then
      VALIPROY% = 1
      Exit Function
    End If
  End If
  VALIPROY% = -1
  '
End Function
  
  '
'\\\OT-06-0244-WAR-FIN///

Sub RECUMOCAJ(NUMEFA$, FECONTAFA, OKEYY%)
   '
   ' RECUPERA MOVIMIENTOS DE C & BANCOS NO BALANCEADOS HASTA $ 0,50
   OKEYY% = 0
   FECONTAD = Int(CDbl(FECONTAFA))
   
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where (CodiCom_Cor = '" & NUMEFA$ & "' OR CodiCom_Lar = '" & NUMEFA$ & "') "
   SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) = " & FECONTAD & " "
   SQLX$ = SQLX$ + "Order by ABS(IDebeComp - IHabeComp) ASC "
   '
   Set MOFASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   SUMACOMP# = 0
   On Error Resume Next
   With MOFASEL
     .MoveFirst
     If Err > 0 Then
       On Error GoTo 0
       Exit Sub
     End If
     Do While Not .EOF
       IDEBAS# = ROUND#(!IDEBECOMP, 2)
       IHABAS# = ROUND#(!IHABECOMP, 2)
       '
       .Edit
       !IDEBECOMP = IDEBAS#
       !IHABECOMP = IHABAS#
       !ValAbsComp = Abs(IDEBAS# - IHABAS#)
       .Update
       SUMACOMP# = SUMACOMP# + IDEBAS# - IHABAS#
     .MoveNext
     Loop
     If Abs(SUMACOMP#) < 0.005 Then OKEYY% = 1
   End With
   On Error GoTo 0
   '
   If Abs(SUMACOMP#) >= 0.005 Then
     If Abs(SUMACOMP#) < 0.5 Then
       TTXX$ = "Diferencia de Redondeo (" + TRIM$(FORMATNUM$(SUMACOMP#, "F12.2")) + ") en Comprobante " + TRIM$(NUMEFA$)
       TTXX$ = TTXX$ + "\   \Desea Recuperar Esta Diferencia Automáticamente ? \\Si, Recuperar\No, Dejar como está"
       If COMALTER%(TTXX$) = 1 Then
         MOFASEL.MoveLast
         MOFASEL.Edit
         If MOFASEL!IDEBECOMP >= 0.005 Then
             MOFASEL!IDEBECOMP = MOFASEL!IDEBECOMP - SUMACOMP#
           Else
             MOFASEL!IHABECOMP = MOFASEL!IHABECOMP + SUMACOMP#
         End If
         MOFASEL.Update
         OKEYY% = 1
       End If
     End If
   End If
   '
End Sub



