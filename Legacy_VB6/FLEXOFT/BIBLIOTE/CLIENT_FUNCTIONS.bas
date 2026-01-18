Attribute VB_Name = "CLIENT_FUNCTIONS"
' ACCDEUDO.QBS                         ' accesos varios a archivo de deudores
'
Public ARCHE$(16)
Public NC As Integer
Public MODO As Integer
Public NROPRO&
Public NCLILLAMA%
'
Public VERICUIT%
'
Dim PUCLI(1) As String
Dim PUBAC(1) As String
'
Sub DATCLIEN(NCLIE, VDEV$, OP%)
'
   Static REBLA$(1)
   Static CAPOS%, PIVA$(), PIVA2$(), CTX%, REG&
   Static NCA%, TIBUA%, TANTU, XX$, XBA$, XBC$
   Static SIGANT%
   Static URECUEN&
   '
   VDEV$ = ""
   If CTX% = 0 Or Sgn(NCLIE) <> SIGANT% Then
       URECUEN& = ULTREG&("CUENTAS")
       ARX1$ = "DEUDOR1"
       If NCLIE < 0 Then
         ARX1$ = "PROVED1"
       End If
       On Error Resume Next
       TTX1# = CDbl(FileDateTime(LUDAT$ + ARX1$ + ".DAT"))
       TTX2# = CDbl(FileDateTime(LUDAT$ + ARX1$ + ".X01"))
       On Error GoTo 0
       'If TTX1# > TTX2# Then
       '  If ULTREG&(ARX1$) > 0 Then
       '    X$ = REGLEIDO$(ARX1$, 1)      ' fuerza la
       '    Call GRAREG(ARX1$, X$, 1)     ' reindexacion
       '    Call CIERRARCH(ARX1$)         ' del archivo
       '    TTX1# = 0: TTX2# = 0
       '    On Error Resume Next
       '    TTX1# = CDbl(FileDateTime(LUDAT$ + ARX1$ + ".DAT"))
       '    TTX2# = CDbl(FileDateTime(LUDAT$ + ARX1$ + ".X01"))
       '    On Error GoTo 0
       '    If TTX1# > TTX2# Then
       '      Call MENSERR(24, "Error de Indices '" + ARX1$ + ".DAT' - Imposible Continuar")
       '      Call FINAL("")
       '    End If
       '  End If
       'End If
       '
       Call ACTUDESC           ' ACTUALIZAR TABLA DE DESCUENTOS
       CTX% = 1
   End If
   '
   SIGANT% = Sgn(NCLIE)
   If Abs(NCLIE) < 1 Then
       If OP% <> 0 Then
           Call MENSERR(24, "Numero de Cuenta Incorrecto Invocando Rutina 'DATCLIEN'.")
       End If
       Exit Sub
   End If
   '
   TIBU% = 0: If NCLIE < 0 Then TIBU% = 1
   If TIBU% = TIBUA% Then
     If NCLIE = NCA% Then
       If Timer < TANTU + 1 Then
         GoTo SELECTAR
       End If
     End If
   End If
   '
   ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2": ARX3$ = "BADACLIE": TIBU% = 0
   If NCLIE < 0 Then
     ARX1$ = "PROVED1": ARX2$ = "PROVED2": ARX3$ = "BADAPROV": TIBU% = 1
   End If
   If REBLA$(TIBU%) = "" Then
      REBLA$(TIBU%) = REGBLAN$(ARX1$) + REGBLAN$(ARX2$)
   End If
   '
10 If PUCLI$(TIBU%) = "" Then
     Call CARPUNCLI(TIBU%)
   End If
   '
3  ABUS$ = MKI$(Abs(NCLIE))
   For U& = 1 To Len(PUCLI$(TIBU%)) Step 2
     If CVI(Mid$(PUCLI$(TIBU%), U&, 2)) < 1 Then
       X$ = REGLEIDO$(ARX1$, (U& + 1) / 2)
       If CVI(Left$(X$, 2)) <> CVI(Mid$(PUCLI$(TIBU%), U&, 2)) Then
         X$ = REGLEIDO$(ARX1$, 1)
         Call GRAREG(ARX1$, X$, 1)
         Call CIERRARCH(ARX1$)
         PUCLI$(TIBU%) = ""
         GoTo 10
       End If
     End If
     '
     If Mid$(PUCLI$(TIBU%), U&, 2) = ABUS$ Then
       X$ = REGLEIDO$(ARX1$, (U& + 1) / 2)
       If Left$(X$, 2) = ABUS$ Then
         REG& = (U& + 1) / 2 '  REGNUM(ARX1$)
         GoTo CONTINUAR
       End If
       X$ = REGLEIDO$(ARX1$, 1)
       Call GRAREG(ARX1$, X$, 1)
       Call CIERRARCH(ARX1$)
       PUCLI$(TIBU%) = ""
       GoTo 10
     End If
   Next U&
'AAA = Len(PUCLI$(TIBU%)) / 2
'bbb = ULTREG&(ARX1$)
'CCC = CVI(Mid$(PUCLI$(TIBU%), 1309, 2))
'ddd = CVI(Mid$(PUCLI$(TIBU%), 1311, 2))
'PUCLI$(TIBU%) = ""
   If Len(PUCLI$(TIBU%)) / 2 < ULTREG&(ARX1$) Then
     PUCLI$(TIBU%) = ""
     GoTo 10
   End If
   REG& = 0: GoTo CONTINUAR
   '
CONTINUAR:
   If REG& > 0 And REG& <= ULTREG&(ARX1$) Then
      XX$ = REGLEIDO$(ARX1$, REG&) + REGLEIDO$(ARX2$, REG&)
      'If TIBU% = 0 Then
      '  If URECUEN& > 0 Then
      '    VDL& = CVI(Mid$(XX$, 159, 2))
      '    If TRIM$(Mid$(XX$, 33, 25)) = "" And VDL& < 1 Or VDL& > URECUEN& Then
      '      For KK& = 1 To ULTREG&("BKDATCLI")
      '        XXYY$ = REGLEIDO$("BKDATCLI", KK&)
      '        If CVI(Left$(XXYY$, 2)) = NCLIE Then
      '          XX$ = XXYY$
      '          GoTo 23
      '        End If
      '      Next KK&
      '    End If
      '  End If
      'End If
    Else
      REG& = 0
      XX$ = REBLA$(TIBU%)
   End If
23 NCA% = NCLIE: TIBUA% = TIBU%: TANTU = Timer
   '
   If OP% = 12 Then GoTo SELECTAR
   XBC$ = ""
   ABADA$ = "BADACLIE": If NCLIE < 0 Then ABADA$ = "BADAPROV"
   If Abs(NCLIE) > 0 Then
      RFF$ = RECFILT$(ABADA$, 1, MKI$(Abs(NCLIE)))
      REBACLI& = 0
      If Len(RFF$) >= 4 Then
         REBACLI& = CVS(Left$(RFF$, 4))
      End If
      If REBACLI& > 0 Then
         If REBACLI& <= ULTREG&(ABADA$) Then
            XBC$ = REGLEIDO$(ABADA$, REBACLI&)
            If CVI(Left$(XBC$, 2)) <> Abs(NCLIE) Then
               XBC$ = ""
            End If
         End If
      End If
   End If
   '
SELECTAR:
   Select Case OP%
      '
      Case 0                                     ' numero de registro
         VDSS$ = Str$(REG&)
      Case 1                                     ' numero de cuenta
         VDSS$ = Str$(NCLIE)
      Case 2                                     ' denominacion - razon social
         VDSS1$ = Mid$(XX$, 3, 30)
         If XBC$ <> "" Then VDSS1$ = Mid$(XBC$, 3, 62)
         VDSS$ = REPLACAR$(VDSS1$, "@", " ")
      Case 3                                     ' domicilio: calle y numero
         VDSS$ = Mid$(XX$, 33, 25)
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 97, 64)
      Case 4                                     ' numero de calle
         VDSS$ = ""
      Case 5                                     ' codigo postal
         VDSS$ = Str$(CVI(Mid$(XX$, 60, 2)))
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 161, 16)
      Case 6                                     ' localidad
         VDSS$ = Mid$(XX$, 62, 25)
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 177, 48)
      Case 7                                     ' telefono 1
         VDSS$ = Mid$(XX$, 87, 10)
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 257, 24)
      Case 8                                     ' telefono 2
         VDSS$ = Mid$(XX$, 97, 10)
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 65, 32)
      Case 9                                     ' numero de C.U.I.T.
         VDSS$ = Mid$(XX$, 107, 15)
      Case 10                                    ' numero de I.Brutos
         VDSS$ = Mid$(XX$, 122, 13)
      Case 11                                    ' numero de proveedor
         VDSS$ = ""
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 433, 16)
         If TRIM$(VDSS$) = "" Then
           If CVS(Mid$(XX$, 137, 4)) >= 1 Then
             VDSS$ = Str$(CVS(Mid$(XX$, 137, 4)))
           End If
         End If
      Case 12                                    ' saldo de la cuenta
         VDSS$ = NUMSTRI$(CVD(Mid$(XX$, 141, 8)), 15, 2, 0)
      Case 13                                    ' codigo de provincia
         VDSS$ = Mid$(XX$, 149, 1)
      Case 14                                    ' codigo condicion de pago
         VDSS$ = Str$(Asc(Mid$(XX$, 150, 1)))
      Case 15                                    ' codigo lista de precios
         VDSS$ = Str$(Asc(Mid$(XX$, 151, 1)))
      Case 16                                    ' codigo de vendedor
         VDSS$ = Str$(Asc(Mid$(XX$, 152, 1)))
      Case 17                                    ' codigo posicion iva
         VDSS$ = Str$(Asc(Mid$(XX$, 136, 1)))
         If NCLIE < 0 Then VDSS$ = Str$(Asc(Mid$(XX$, 152, 1)))
      Case 18
         VDSS$ = "N/Def."
         VD% = Asc(Mid$(XX$, 136, 1))
         If NCLIE < 0 Then VD% = Asc(Mid$(XX$, 152, 1))
         If VD% >= 0 Then If VD% <= CAPOS% Then VDSS$ = PIVA$(VD%)
      Case 19
         VDSS$ = "Pos. IVA no Definida"
         VD% = Asc(Mid$(XX$, 136, 1))
         If NCLIE < 0 Then VD% = Asc(Mid$(XX$, 152, 1))
         If VD% >= 0 Then If VD% <= CAPOS% Then VDSS$ = PIVA$(VD%)
      Case 20
         VDSS$ = "Descripcion Lista de Precios"
      Case 21
         VDSS$ = "Descripcion Condicion de pago"
      Case 22
         VDSS$ = "Denominacion del Vendedor"
      Case 23
           VDD# = CVS(Mid$(XX$, 153, 4))    ' LIMITE DE CREDITO
           VDSS$ = Str$(VDD#)
      Case 24
           VD% = CVI(Mid$(XX$, 159, 2))
           If NUCUEN% < 1 Then NUCUEN% = ULTREG&("CUENTAS")
           If VD% < 1 Or VD% > NUCUEN% Then
             CQXX$ = "AMACLI": If NCLIE < 0 Then CQXX$ = "AMAPRO"
             VD% = CUECOINT%(CONTROL$(CQXX$, "ECUECON"))
           End If
           VDSS$ = Str$(VD%)
      Case 25
           VDL& = CVI(Mid$(XX$, 159, 2))
           NCUECOI& = ULTREG&("CUENTAS")
           If VDL& > 0 And VDL& <= NCUECOI& Then
               VDSS$ = Left$(REGLEIDO$("CUENTAS", VDL&), 12)
             Else
               CQXX$ = "AMACLI": If NCLIE < 0 Then CQXX$ = "AMAPRO"
               VDSS$ = CONTROL$(CQXX$, "ECUECON")
           End If
      Case 26
           VD% = CVI(Mid$(XX$, 157, 2))
           VDSS$ = Str$(VD%)
      Case 27
         VDSS$ = ""
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 225, 32)
      Case 28
         VDSS$ = ""
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 281, 24)
      Case 29
         VDSS$ = ""
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 305, 48)
      Case 30
         VDSS$ = ""                    ' persona de contacto
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 385, 48)
      Case 31                                    ' CATEGORIA de I.Brutos
         VDSS$ = Mid$(XX$, 135, 1)
      Case 32
         VDSS$ = ""                    ' HORARIO
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 449, 32)
      Case 33
         VDSS$ = ""                    ' ZONA
         If XBC$ <> "" Then VDSS$ = Mid$(XBC$, 481, 12)
   End Select
   '
   VDEV$ = TRIM$(VDSS$)
   '
End Sub
'
Sub CARPUNCLI(TIBU%)
   '
   If TIBU% = 0 Then
     ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2": ARX3$ = "BADACLIE"
   ElseIf TIBU% = 1 Then
     ARX1$ = "PROVED1": ARX2$ = "PROVED2": ARX3$ = "BADAPROV"
   Else
     Call MENSERR(24, "Error de Invocación Rutina 'CARPUNCLI'\en Módulo '" + TRIM$(App.EXEName) + "'.")
     Call FINAL("")
   End If
   '
   PUCLI$(TIBU%) = ""
   PUBAC$(TIBU%) = ""
   '
   PUCLI$(TIBU%) = LOADFILE(LUDAT$ + ARX1$ + ".X01")
   PUBAC$(TIBU%) = LOADFILE(LUDAT$ + ARX3$ + ".X01")
   '
   If PUCLI$(TIBU%) = "" Then
     If ULTREG&(ARX1$) < 1 Then GoTo 20
     X$ = REGLEIDO$(ARX1$, 1)
     Call GRAREG(ARX1$, X$, 1)
     Call CIERRARCH(ARX1$)
     PUCLI$(TIBU%) = LOADFILE(LUDAT$ + ARX1$ + ".X01")
     If PUCLI$(TIBU%) = "" Then
       Call MENSERR(24, "Error de Configuración de Indices\de Archivo '" + ARX1$ + "'. - Consulte.")
       Call FINAL("")
     End If
   End If
   '
20 If PUBAC$(TIBU%) = "" Then
     If ULTREG&(ARX3$) < 1 Then GoTo 30
     X$ = REGLEIDO$(ARX3$, 1)
     Call GRAREG(ARX3$, X$, 1)
     Call CIERRARCH(ARX3$)
     PUBAC$(TIBU%) = LOADFILE(LUDAT$ + ARX3$ + ".X01")
     If PUBAC$(TIBU%) = "" Then
       Call MENSERR(24, "Error de Configuración de Indices\de Archivo '" + ARX3$ + "'. - Consulte.")
       Call FINAL("")
     End If
   End If
   '
30 End Sub
'
Function REGICLI&(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 0)
   REGICLI& = Val(VDEV$)
End Function
'
Function CODICLI$(NCLIE)   ' agregado por compatibilidad con versión en access
   If NCLIE > 0 Then
       CODICLI$ = "CL-" + TRIM$(Str$(Abs(NCLIE)))
     Else
       CODICLI$ = "PR-" + TRIM$(Str$(Abs(NCLIE)))
   End If
End Function
'
Function RASOCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 2)
   RASOCLI$ = VDEV$
End Function
'
Function DOMICLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 3)
   DOMICLI$ = VDEV$
End Function
'
Function CPOSCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 5)
   CPOSCLI$ = VDEV$
End Function
'
Function LOCACLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 6)
   LOCACLI$ = VDEV$
End Function
'
Function CPLOCLI$(NCLIE)
   CPX$ = TRIM$(CPOSCLI$(NCLIE))
   If CPX$ <> "" Then CPX$ = "(" + CPX$ + ") "
   CPLOCLI$ = TRIM$(CPX$ + LOCACLI$(NCLIE))
End Function
'
Function PAISCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 27)
   PAISCLI$ = VDEV$
End Function
'
Function TEL1CLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 7)
   TEL1CLI$ = VDEV$
End Function
'
Function TEL2CLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 8)
   TEL2CLI$ = VDEV$
End Function
'
Function FAXCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 28)
   FAXCLI$ = VDEV$
End Function
'
Function EMAILCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 29)
   EMAILCLI$ = VDEV$
End Function
'
Function CONTACTCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 30)
   CONTACTCLI$ = VDEV$
End Function
'
Function HORARCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 32)
   HORARCLI$ = VDEV$
End Function
'
Function DOMIENT$(NCLIE)
   '
   NCLIE = NCLIE
   DOMIENT$ = ""
   If NCLIE > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCLIE))
     DOMIENT$ = TRIM$(Mid$(XX$, 3, 32))
   End If
   '
End Function
'
Function LOCAENT$(NCLIE)
   '
   LOCAENT$ = ""
   If NCLIE > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCLIE))
     LOCAENT$ = TRIM$(Mid$(XX$, 35, 32))
   End If
   '
End Function
'
Function TRANENT$(NCLIE)
   '
   TRANENT$ = ""
   If NCLIE > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCLIE))
     TRANENT$ = TRIM$(Mid$(XX$, 67, 32))
   End If
   '
End Function
'
Function NUTRANCLI%(NCLIE)
   '
   NUTRANCLI% = 0
   If NCLIE > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCLIE))
     NUTRANCLI% = CVI(Mid$(XX$, 99, 32))
   End If
   '
End Function
'
Function DENOTRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    DENOTRANS$ = TRIM$(Mid$(XX$, 3, 32))
End Function

Function RASOTRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    RASOTRANS$ = TRIM$(Mid$(XX$, 35, 48))
End Function

Function DOMITRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    DOMITRANS$ = TRIM$(Mid$(XX$, 83, 48))
End Function

Function LOCATRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    LOCATRANS$ = TRIM$(Mid$(XX$, 131, 48))
End Function

Function TELETRANS$(NTRANI%)
    XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
    TELETRANS$ = TRIM$(Mid$(XX$, 195, 48))
End Function

Function NTRANCLI$(NCLIE)           ' NOMBRE DEL TRANSPORTISTA
    NTRANCLI$ = ""
    NCLIE = NCLIE
    If NCLIE > 0 Then
      NTR% = NUTRANCLI%(NCLIE)
      If NTR% > 0 Then
        NTRANCLI$ = DENOTRANS$(NTR%)
      End If
    End If
End Function

Function RTRANCLI$(NCLIE)           ' R.SOCIAL DEL TRANSPORTISTA
    RTRANCLI$ = ""
    NCLIE = NCLIE
    If NCLIE > 0 Then
      NTR% = NUTRANCLI%(NCLIE)
      If NTR% > 0 Then
        RTRANCLI$ = RASOTRANS$(NTR%)
      End If
    End If
End Function

Function DTRANCLI$(NCLIE)           ' DOMICILIO DEL TRANSPORTISTA
    DTRANCLI$ = ""
    If NCLIE > 0 Then
      NTR% = NUTRANCLI%(NCLIE)
      If NTR% > 0 Then
        DTRANCLI$ = DOMITRANS$(NTR%)
      End If
    End If
End Function

Function LTRANCLI$(NCLIE)           ' LOCALIDAD DEL TRANSPORTISTA
    LTRANCLI$ = ""
    If NCLIE > 0 Then
      NTR% = NUTRANCLI%(NCLIE)
      If NTR% > 0 Then
        LTRANCLI$ = LOCATRANS$(NTR%)
      End If
    End If
End Function

Function HORAENT$(NCLIE)
   '
   HORAENT$ = ""
   If NCLIE > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NCLIE))
     HORAENT$ = TRIM$(Mid$(XX$, 101, 24))
   End If
   '
End Function
'
Function ZONACLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 33)
   ZONACLI$ = VDEV$
End Function
'
Function CUITCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 9)
   CUITCLI$ = VDEV$
End Function
'
Function NIBRCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 10)
   NIBRCLI$ = VDEV$
End Function
'
Function NPROCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 11)
   NPROCLI$ = VDEV$
End Function
'
Function SALDCLI#(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 12)
   SALDCLI# = Val(VDEV$)
End Function
'
Function CHEPEACRE#(NCLIE)
   RFF$ = RECFILT$("CHECAR", 3, MKI$(NCLIE))
   SUCHEPE# = 0
   HOII% = HOY(HOS$)
   For UI& = 1 To Len(RFF$) Step 4
      RECHE& = CVS(Mid$(RFF$, UI&, 4))
      If RECHE& > 0 Then
         If RECHE& <= ULTREG&("CHECAR") Then
            XX$ = REGLEIDO$("CHECAR", RECHE&)
            If CVI(Mid$(XX$, 3, 2)) = NCLIE Then
               If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                  SUCHEPE# = SUCHEPE# + CVD(Mid$(XX$, 33, 8))
               End If
            End If
         End If
      End If
   Next UI&
   CHEPEACRE# = SUCHEPE#
End Function
'
Function PROVCLI$(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 13)
   PROVCLI$ = VDEV$
End Function
'
Function CPAGCLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 14)
   CPAGCLI% = Val(VDEV$)
End Function
'
Function LIPRCLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 15)
   LIPRCLI% = Val(VDEV$)
End Function
'
Function TICUEPRO%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 15)
   TICUEPRO% = Val(VDEV$)
End Function
'
Function VENDCLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 16)
   VENDCLI% = Val(VDEV$)
End Function
'
Function PIVACLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 17)
   PIVACLI% = Val(VDEV$)
End Function
'
Function CATIBRU$(NCLIE)
   '
   CATIBRU$ = ""
   If NCLIE > 0 Then
      Call DATCLIEN(NCLIE, VDEV$, 31)
      CATIBRU$ = VDEV$
      Exit Function
   End If
   '
   For UU& = 1 To ULTREG&("TARETIB")
      XX$ = REGLEIDO$("TARETIB", UU&)
      If CVI(Left$(XX$, 2)) = (-1) * NCLIE Then
          CATIBRU$ = Mid$(XX$, 55, 1)
          Exit Function
      End If
   Next UU&
   '
End Function
'
Function DESCOMBINA$(NCX%, cix%)
   '
   DCX$ = ""
   If NCX% > 0 Then
     If cix% > 0 Then
       GRX% = GRUCOCLI%(NCX%)
       LNX% = GRUCOVEI%(cix%)
       '
       DCX$ = ""
       For KU& = 1 To ULTREG&("DESCOMBI")
         RECORD$ = REGLEIDO$("DESCOMBI", KU&)
         NCY% = CVI(Mid$(RECORD$, 17, 2))
         If NCY% = 0 Or NCY% = NCX% Then
           GRY% = CVI(Mid$(RECORD$, 19, 2))
           If GRY% = 0 Or GRY% = GRX% Then
             CIY% = CVI(Mid$(RECORD$, 21, 2))
             If CIY% = 0 Or CIY% = cix% Then
                LNY% = CVI(Mid$(RECORD$, 23, 2))
                If LNY% = 0 Or LNY% = LNX% Then
                  DCY$ = TRIM$(Mid$(RECORD$, 49, 12))
                  If DCY$ <> "" Then
                     If DCX$ <> "" Then DCX$ = DCX$ + "+"
                     DCX$ = DCX$ + DCY$
                  End If
                End If
             End If
           End If
         End If
       Next KU&
     End If
   End If
   '
   DESCOMBINA$ = DCX$
End Function

Function PORCEDESCU(TDESCUEN$)
    '
    COEFI = 1
    DESCX$ = TRIM$(REPLACAR$(TDESCUEN$, "%", ""))
    '
    Do While Len(DESCX$) > 0
       PPXX% = InStr(DESCX$, "+")
       If PPXX% < 1 Then PPXX% = 1 + Len(DESCX$)
       AKX1$ = Left$(DESCX$, PPXX% - 1)
       DESCX$ = TRIM$(Mid$(DESCX$, PPXX% + 1))
       '
       PDC = XVALO(AKX1$)
       COEFI = COEFI * (1 - PDC / 100)
    Loop
    '
99  PORCEDESCU = 100 * (1 - COEFI)
    '
End Function
'

Function DESCUCLI$(NCLIE)
   DSCX$ = ""
   For U& = 1 To ULTREG&("DESCUCLI")
       X$ = REGLEIDO$("DESCUCLI", U&)
       If CVI(Left$(X$, 2)) = NCLIE Then
           PDX0$ = TRIM$(Mid$(X$, 3, 14))
           DSCX$ = REPLACAR$(PDX0$, "%", " ")
           GoTo 99
       End If
   Next U&
   '
99 DESCUCLI$ = DSCX$
   '
End Function
'
Function PORDESCLI(NCLIE)
    '
    Static PDCLI(16384), CTCLI%(16384)
    '
    If NCLIE < 1 Then
       PDXC = 0
       GoTo 99
    End If
    '
    If NCLIE < 16385 Then
       If CTCLI%(NCLIE) > 0 Then
          PDXC = PDCLI(NCLIE)
          GoTo 99
       End If
    End If
    '
    DESC$ = DESCUCLI$(NCLIE)
    PDXC = 0
    DESCX$ = REPLACAR$(DESC$, "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = Val(DESCX$)
        GoTo 95
    End If
    '
    XDC% = InStr(DESCX$, "+")
    DESC1 = Val(Left$(DESCX$, XDC% - 1))
    DESC2 = Val(Mid$(DESCX$, XDC% + 1))
    PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    '
95  If NCLIE < 16385 Then
       CTCLI%(NCLIE) = 1
       PDCLI(NCLIE) = PDXC
    End If
    '
99  PORDESCLI = PDXC
    '
End Function
'
Function LIMICRE(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 23)
   LIMICRE = Val(VDEV$)
End Function
'
Function CUECOCLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 24)
   CUECOCLI% = Val(VDEV$)
End Function
'
Function CUEMADRE$(NCLIE)
   '
   Static KKXX%, CUEDEUVEN$
   If KKXX% < 1 Then
     CUEDEUVEN$ = TRIM$(CONTROL$("AMACLI", "ECUECON"))
     KKXX% = 1
   End If
   '
   Call DATCLIEN(NCLIE, VDEV$, 25)
   CUEMADRE$ = VDEV$
   '
   If TRIM$(VDEV$) = "" Then
     If NCLIE > 0 Then
       CUEMADRE$ = CUEDEUVEN$
     End If
   End If
   '
End Function
'
Function GRUCOCLI%(NCLIE)
   Call DATCLIEN(NCLIE, VDEV$, 26)
   GRUCOCLI% = Val(VDEV$)
End Function

Function DEGRUVE$(NCLIE)
   GRC% = GRUCOCLI%(NCLIE)
   DEGRUVE$ = TRIM$(ECOARCH$("GRUCOCLI", MKI$(GRC%)))
End Function

Sub NOTACLI(NCLIE, MODO%)
   '
   ' MODO% = 1    abre automaticamente si hay nota para el cliente
   ' MODO% = 2    se pidió apertura por botón de acceso
   '
   If ECOARCH$("DEUDOR1", MKI$(NCLIE)) = "" Then Exit Sub
   '
   RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCLIE))
   If Len(RNC$) > 0 Or MODO% = 2 Then
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To Len(RNC$) Step 4
         RENOTA& = CVS(Mid$(RNC$, U&, 4))
         If RENOTA& > 0 Then
            If RENOTA& <= ULTREG&("NOTACLI") Then
               X$ = REGLEIDO$("NOTACLI", RENOTA&)
               JJ& = JJ& + 1
               Call GRAREG("!NOTACLI", X$, JJ&)
            End If
         End If
      Next U&
      Call SETULTREG("!NOTACLI", JJ&)
      Call CIERRARCH("!NOTACLI")
      Screen.MousePointer = 1
      VTB% = VENTABU%("NOTACLI")
      If VTB% >= 0 Then
         Screen.MousePointer = 11
         URN& = ULTREG&("!NOTACLI")
         JJ& = ULTREG&("NOTACLI")
         While Len(RNC$) < 4 * URN&
            JJ& = JJ& + 1
            RNC$ = RNC$ + MKS$(JJ&)
         Wend
         KK& = (-3)
         For JJ& = 1 To ULTREG&("!NOTACLI")
            X$ = REGLEIDO$("!NOTACLI", JJ&)
            Call REPLA(X$, MKI$(NCLIE), 1, 2)
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", X$, RENOTA&)
         Next JJ&
         While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", MKI$(NCLIE) + Space$(78), RENOTA&)
         Wend
         Call CIERRARCH("NOTACLI")
         Screen.MousePointer = 1
      End If
   End If
End Sub
'
Function DEPROVI$(PR$)
   DEPROVI$ = ECOARCH$("PROVINC", PR$)
End Function
'
Function DELIPRE$(CL%)
   DX$ = "PRECIOS LIBRES"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("LISTAS", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "PRECIOS LIBRES"
   DELIPRE$ = DX$
End Function
'
Function DECONPA$(CL%)
   DX$ = "CONTADO"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("CONPAG", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "CONTADO"
   DECONPA$ = DX$
End Function
'
Function DEVENDE$(CL%)
   DX$ = "GERENCIA"
   If CL% > 0 Then DX$ = ECOARCH$("VENDEDOR", Chr$(VENDCLI%(CL%)))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "GERENCIA"
   DEVENDE$ = DX$
End Function
'
Function DEGRUCLI$(CL%)
   DX$ = "NO CLASIFICADO"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOCLI", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "NO CLASIFICADO"
   DEGRUCLI$ = DX$
End Function
'
Function POSIVAC$(NCLIE)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "CONS.FIN."
    PIVA$(1) = "RESP.INSC"
    PIVA$(2) = "R.NO INSC"
    PIVA$(3) = "NO RESP."
    PIVA$(4) = "EXENTO"
    PIVA$(5) = "COM.EXT. "
    PIVA$(6) = "MONOTR."
    PIVA$(7) = "EXEN-PRO"
    '
    APV% = PIVACLI%(NCLIE)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
    POSIVAC$ = PS$
End Function
'
Function POSIVAL$(NCLIE)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "A CONSUMIDOR FINAL"
    PIVA$(1) = "IVA RESPONSABLE INSCRIPTO"
    PIVA$(2) = "IVA RESPONSABLE NO INSCRIPTO"
    PIVA$(3) = "IVA NO RESPONSABLE"
    PIVA$(4) = "IVA EXENTO"
    PIVA$(5) = "OPs.COMERCIO EXTERIOR"
    PIVA$(6) = "CONTRIBUYENTE MONOTRIBUTO"
    PIVA$(7) = "EXENTO - ZONA PROMOVIDA"
    '
    If InStr(EMPREAC$, "SABO") > 0 Then
      PIVA$(5) = "EXENTO OPs. EXPORTACION"
    End If
    '
    APV% = PIVACLI%(NCLIE)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
POSIVAL$ = PS$
End Function
'
Function VALICUIT%(CUIT$)
    '
    CU$ = TRIM$(CUIT$)
    VC% = 0
    If Len(CU$) < 11 Then GoTo 999                                    ' no valida
    '
    VX% = Asc(Mid$(CU$, 3))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 2) + "-" + Mid$(CU$, 3)
        End If
    End If
    '
    If Len(CU$) < 12 Then GoTo 999
    VX% = Asc(Mid$(CU$, 12))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 11) + "-" + Mid$(CU$, 12)
        End If
    End If
    '
    If Len(CU$) <> 13 Then GoTo 999                                   ' no valida
    '
    For J% = 1 To 13
        If J% <> 3 Then
            If J% <> 12 Then
                VX% = Asc(Mid$(CU$, J%, 1))
                If VX% < 48 Or VX% > 57 Then GoTo 999                 ' no valida
            End If
        End If
    Next J%
    '
    Mid$(CU$, 3, 1) = "-"
    Mid$(CU$, 12, 1) = "-"
    '
    SUMA% = 0
    SUMA% = SUMA% + 5 * Val(Mid$(CU$, 1, 1))
    SUMA% = SUMA% + 4 * Val(Mid$(CU$, 2, 1))
    SUMA% = SUMA% + 3 * Val(Mid$(CU$, 4, 1))
    SUMA% = SUMA% + 2 * Val(Mid$(CU$, 5, 1))
    SUMA% = SUMA% + 7 * Val(Mid$(CU$, 6, 1))
    SUMA% = SUMA% + 6 * Val(Mid$(CU$, 7, 1))
    SUMA% = SUMA% + 5 * Val(Mid$(CU$, 8, 1))
    SUMA% = SUMA% + 4 * Val(Mid$(CU$, 9, 1))
    SUMA% = SUMA% + 3 * Val(Mid$(CU$, 10, 1))
    SUMA% = SUMA% + 2 * Val(Mid$(CU$, 11, 1))
    '
    SUMB% = SUMA% Mod 11
    DIVER% = 0
    If SUMB% Mod 11 <> 0 Then DIVER% = 11 - SUMB%
    If DIVER% <> Val(Mid$(CU$, 13, 1)) Then GoTo 999                  ' no valida
    '
    VC% = 1
    CUIT$ = CU$

999 VALICUIT% = VC%
    '
End Function

Function REPICUIT%(CUTT$, NC1%)
   '
   RPP% = 0
   If Abs(NC1%) < 1 Then GoTo 99
   '
   Screen.MousePointer = 11
   ADEU1$ = "DEUDOR1": ADEU2$ = "DEUDOR2"
   If NC1% < 0 Then
      ADEU1$ = "PROVED1": ADEU2$ = "PROVED2"
   End If
   '
   For U& = 1 To ULTREG&(ADEU1$)
      X1$ = REGLEIDO$(ADEU1$, U&) + REGLEIDO$(ADEU2$, U&)
      If TRIM$(Mid$(X1$, 107, 15)) = CUTT$ Then
         If CVI(Left$(X1$, 2)) <> Abs(NC1%) Then
            RPP% = CVI(Left$(X1$, 2)) * Sgn(NC1%)
            GoTo 99
         End If
      End If
   Next U&
   '
99 Screen.MousePointer = 1
   REPICUIT% = RPP%
   '
End Function
'
Sub CARDATPRO(NCLIE)
    '
    Screen.MousePointer = 11
    '
    MODIPRO.Text2.TEXT = TEL2CLI$(NCLIE)
    MODIPRO.Text3.TEXT = RASOCLI$(NCLIE)
    MODIPRO.Text4.TEXT = DOMICLI$(NCLIE)
    MODIPRO.Text5.TEXT = CPOSCLI$(NCLIE)
    MODIPRO.Text6.TEXT = LOCACLI$(NCLIE)
    MODIPRO.Text10.TEXT = TEL1CLI$(NCLIE)
    MODIPRO.Text15.TEXT = CUITCLI$(NCLIE)
    MODIPRO.Text18.TEXT = NIBRCLI$(NCLIE)
    MODIPRO.Text9.TEXT = PAISCLI$(NCLIE)
    MODIPRO.Text11.TEXT = FAXCLI$(NCLIE)
    MODIPRO.Text12.TEXT = EMAILCLI$(NCLIE)
    '
    MODIPRO.COTEXT(0).TEXT = TRIM$(PROVCLI$(NCLIE))
    MODIPRO.COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NCLIE)))
    MODIPRO.COTEXT(2).TEXT = TRIM$(CATIBRU$(NCLIE))
    MODIPRO.COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NCLIE)))
    MODIPRO.COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NCLIE)))
    MODIPRO.COTEXT(5).TEXT = TRIM$(CUEMADRE$(NCLIE))
    '
    For II = 0 To 5
       ECO$ = ""
       If TRIM$(MODIPRO.COTEXT(II).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(II), 1)
          LF11% = LENFIELD(ARCHE$(II), 1)
          ARGU$ = CBIN$(TRIM$(MODIPRO.COTEXT(II).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(II), ARGU$))
       End If
       '
       If ECO$ <> "" Then
             MODIPRO.DETEXT(II).TEXT = ECO$
          Else
             MODIPRO.DETEXT(II).TEXT = "No Definido"
       End If
    Next II
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub CARTAVALI()
    '
    If ULTREG&("CARETIB") < 4 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "0I.Brutos No Responsable", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "1Contribuyente Puro", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "2Convenio Multilateral", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "3Conv.Multil. (Comerciales)", IQ&)
      Call SETULTREG("CARETIB", IQ&)
      Call CIERRARCH("CARETIB")
    End If
    '
    If ULTREG&("TICUEPRO") < 2 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(0) + "Cuenta Corriente", IQ&)
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(1) + "Cuenta de Gastos", IQ&)
      Call SETULTREG("TICUEPRO", IQ&)
      Call CIERRARCH("TICUEPRO")
    End If
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext.", 8) + String$(24, 0), IQ&)
          If InStr(EMPREAC$, "SABO") > 0 Then Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operaciones Exportacion", 31) + AJUSTI$("Export.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    For I& = ULTREG&("LISTAS") To 1 Step -1
       XX$ = REGLEIDO$("LISTAS", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 59
       Call GRAREG("LISTAS", String$(128, 0), I&)
    Next I&
59  Call CIERRARCH("LISTAS")
    '
     For I& = ULTREG&("CONPAG") To 1 Step -1
       XX$ = REGLEIDO$("CONPAG", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 69
       Call GRAREG("CONPAG", String$(128, 0), I&)
    Next I&
69  Call CIERRARCH("CONPAG")
    '
    For I& = ULTREG&("VENDEDOR") To 1 Step -1
       XX$ = REGLEIDO$("VENDEDOR", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 79
       Call GRAREG("VENDEDOR", String$(128, 0), I&)
    Next I&
79  Call CIERRARCH("VENDEDOR")
    '
     For I& = ULTREG&("PROVINC") To 1 Step -1
       XX$ = REGLEIDO$("PROVINC", I&)
       If Asc(Left$(XX$, 1)) > 32 Then GoTo 89
       Call GRAREG("PROVINC", String$(128, 0), I&)
    Next I&
89  Call CIERRARCH("PROVINC")
    '
    ARCHE$(0) = "PROVINC"
    ARCHE$(1) = "CAPOSIV"
    ARCHE$(2) = "CARETIB"
    ARCHE$(3) = "LISTAS"
    ARCHE$(4) = "CONPAG"
    ARCHE$(5) = "VENDEDOR"
    ARCHE$(6) = "ECUECON"
    ARCHE$(7) = "GRUCOCLI"
    '
End Sub

Sub GRADATPRO()
    '
    If NCLIE% > 0 Then
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(MODIPRO.Text15.TEXT)
    If CUIT$ <> "" Then
      If VALICUIT%(CUIT$) <> 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Número de C.U.I.T. no Válido")
        If MODO% = 2 Then MODIPRO.Text15.TEXT = CUITCLI$(NCLIE%)
        MODIPRO.Text15.SetFocus
        Exit Sub
      End If
    End If
    '
    If NCLIE% = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If MODO% = 1 Then
       If Abs(NCLIE%) > 9999 Or Abs(NCLIE%) = 8224 Then
          Call MENSERR(24, "Imposible Grabar - Número de Cuenta no Válido.")
          Exit Sub
       End If
    End If
    '
    ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
    If NCLIE% < 0 Then ARX1$ = "PROVED1": ARX2$ = "PROVED2"
    '
    If MODO% = 1 Then
           REGX& = 1 + ULTREG&(ARX1$)
           XX$ = REGBLAN$(ARX1$) + REGBLAN$(ARX2$)
           Call REPLA(XX$, MKI$(Abs(NCLIE%)), 1, 2)
           If NCLIE% = 0 Or REGICLI&(NCLIE%) > 0 Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Número de Cuenta Ocupado.")
               Exit Sub
           End If
        Else
           REGX& = REGICLI&(NCLIE%)
           If REGX& < 1 Or REGX& > ULTREG&(ARX1$) Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
               Exit Sub
           End If
           '
           XX$ = REGLEIDO$(ARX1$, REGX&) + REGLEIDO$(ARX2$, REGX&)
           If CVI(Left$(XX$, 2)) <> Abs(NCLIE%) Then
               Screen.MousePointer = 1
               Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
               Exit Sub
           End If
    End If
    '
    CARASO% = 0
    If TRIM$(MODIPRO.Text3.TEXT) <> TRIM$(Mid$(XX$, 3, 30)) Then CARASO% = 1
    '
    Call REPLA(XX$, TRIM$(MODIPRO.Text3.TEXT), 3, 30)
    Call REPLA(XX$, TRIM$(MODIPRO.Text4.TEXT), 33, 25)
    CPOST% = 0
    If Val(MODIPRO.Text5.TEXT) > 0 Then
        If Val(MODIPRO.Text5.TEXT) < 32768 Then
           CPOST% = Val(MODIPRO.Text5.TEXT)
        End If
    End If
    Call REPLA(XX$, MKI$(CPOST%), 60, 2)
    Call REPLA(XX$, TRIM$(MODIPRO.Text6.TEXT), 62, 25)
    Call REPLA(XX$, TRIM$(MODIPRO.Text10.TEXT), 87, 10)
    Call REPLA(XX$, TRIM$(MODIPRO.Text2.TEXT), 97, 10)
    Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
    Call REPLA(XX$, TRIM$(MODIPRO.Text18.TEXT), 122, 15)
    Call REPLA(XX$, TRIM$(MODIPRO.COTEXT(0).TEXT), 149, 1)
    '
    PIVA% = VALIMI(MODIPRO.COTEXT(1).TEXT, 0, 7)
    Call REPLA(XX$, Chr$(PIVA%), 152, 1)
    '
    CATIB% = VALIMI(MODIPRO.COTEXT(2).TEXT, 0, 2)
    '
    LIPRE% = VALIMI(MODIPRO.COTEXT(3).TEXT, 0, 1)
    If ECOARCH("TICUEPRO", Chr$(LIPRE%)) = "" Then
        Call MENSERR(24, "Tipo de Cuenta no Válida")
        LIPRE% = 0
    End If
    Call REPLA(XX$, Chr$(LIPRE%), 151, 1)
    '
    COPAGI% = VALIMI(MODIPRO.COTEXT(4).TEXT, 0, 255)
    If COPAGI% > 0 Then
       If ECOARCH("CONPAG", Chr$(COPAGI%)) = "" Then
          Call MENSERR(24, "Condición de Pago no Válida")
          COPAGI% = 0
       End If
    End If
    Call REPLA(XX$, Chr$(COPAGI%), 150, 1)
    '
    CCX% = REGBUS&("CUENTAS", 1, AJUSTI$(MODIPRO.COTEXT(5).TEXT, 12))
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
    CATI$ = TRIM$(Str$(CATIB%))
    For UU& = 1 To ULTREG&("TARETIB")
      XX$ = REGLEIDO$("TARETIB", UU&)
      If CVI(Left$(XX$, 2)) = (-1) * NCLIE% Then
          Call REPLA(XX$, CATI$, 55, 1)
          Call GRAREG("TARETIB", XX$, UU&)
          GoTo 999
      End If
    Next UU&
    '
    XX$ = REGBLAN$("TARETIB")
    Call REPLA(XX$, MKI$((-1) * NCLIE%), 1, 2)
    Call REPLA(XX$, CATI$, 55, 1)
    Call REGAPP("TARETIB", XX$)
    '
999 Call CIERRARCH("TARETIB")
    '
    If CARASO% = 1 Then
       Call FRESHECHO(ARX1$)
       'FORMUSEL.TABLA(ITX%).Clear
    End If
    '
    NCLIE% = Val(MODIPRO.Text1.TEXT)
    If NCLIE% > 0 Then
       RFF$ = RECFILT$("BADAPROV", 1, MKI$(NCLIE%))
       REBACLI& = 0
       If Len(RFF$) >= 4 Then
          REBACLI& = CVS(Left$(RFF$, 4))
       End If
       If REBACLI& < 1 Or REBACLI& > ULTREG&("BADAPROV") Then
         REBACLI& = 0
         For UUL& = 1 To ULTREG&("BADAPROV")
           XX11$ = REGLEIDO$("BADAPROV", UUL&)
           If Abs(CVI(Left$(XX11$, 2))) = NCLIE% Then
             REBACLI& = UUL&
             GoTo 120
           End If
         Next UUL&
         REBACLI& = 1 + ULTREG&("BADAPROV")
       End If
       '
120    XX$ = REGBLAN$("BADAPROV")
       Call REPLA(XX$, MKI$(NCLIE%), 1, 2)
       Call REPLA(XX$, MODIPRO.Text3.TEXT, 3, 62)
       Call REPLA(XX$, MODIPRO.Text2.TEXT, 65, 32)
       Call REPLA(XX$, MODIPRO.Text4.TEXT, 97, 64)
       Call REPLA(XX$, MODIPRO.Text5.TEXT, 161, 16)
       Call REPLA(XX$, MODIPRO.Text6.TEXT, 177, 48)
       Call REPLA(XX$, MODIPRO.Text9.TEXT, 225, 32)
       Call REPLA(XX$, MODIPRO.Text10.TEXT, 257, 24)
       Call REPLA(XX$, MODIPRO.Text11.TEXT, 281, 24)
       Call REPLA(XX$, MODIPRO.Text12.TEXT, 305, 48)
       Call GRAREG("BADAPROV", XX$, REBACLI&)
       Call CIERRARCH("BADAPROV")
    End If
    '
    Call CLEARFORMPRO
    If MODO% = 1 Then
        MODIPRO.Text1.TEXT = ""
        NCLIE% = 0
    End If
    MODIPRO.Text1.SetFocus
    Screen.MousePointer = 1
    '
End Sub

Sub CLEARFORMPRO()
    '
    MODIPRO.Text2.TEXT = ""
    MODIPRO.Text3.TEXT = ""
    MODIPRO.Text4.TEXT = ""
    MODIPRO.Text5.TEXT = ""
    MODIPRO.Text6.TEXT = ""
    MODIPRO.Text9.TEXT = ""
    MODIPRO.Text10.TEXT = ""
    MODIPRO.Text11.TEXT = ""
    MODIPRO.Text12.TEXT = ""
    MODIPRO.Text15.TEXT = ""
    MODIPRO.Text18.TEXT = ""
    '
    For II = 0 To 5
       MODIPRO.COTEXT(II).TEXT = ""
       MODIPRO.DETEXT(II).TEXT = ""
    Next II
    '
    MODIPRO.Refresh
    MODIPRO.Command22.Refresh
    NCLIE = 0
    '
End Sub
'

Sub SORTCLI(ARDEU%)
    '
    Dim DEU1 As String * 32
    Dim DEU2 As String * 128
    Dim REGDEU$(8192)
    Dim OCUCLI%(32767)
    '
    PCLI$ = String$(32767, 0)
    '
    LU$ = LUDAT$
    ARX1$ = LU$ + "DEUDOR1.DAT"
    ASELE$ = "DEUDOR1"
    If ARDEU% = 2 Then
        ARX1$ = LU$ + "PROVED1.DAT"
        ASELE$ = "PROVED1"
    End If
    On Error GoTo 100
    N1% = FreeFile
    Open ARX1$ For Random Access Read Write Lock Read Write As #N1% Len = 32
    '
    ARX2$ = LU$ + "DEUDOR2.DAT"
    If ARDEU% = 2 Then
        ARX2$ = LU$ + "PROVED2.DAT"
    End If
    On Error GoTo 200
    N2% = FreeFile
    Open ARX2$ For Random Access Read Write Lock Read Write As #N2% Len = 128
    '
    Screen.MousePointer = 11
    AGRECLI% = 0
    On Error GoTo 0
    Lock #N1%
    Lock #N2%
    '
    FORSELEC.LISORT.Clear
    For U& = 1 To LOF(N1%) / 32
        Get #N1%, U&, DEU1$
        Get #N2%, U&, DEU2$
        If CVI(Left$(DEU1$, 2)) > 0 Then
            '
            NCXI% = CVI(Left$(DEU1$, 2))
            If NCXI% <= 32767 Then
              If OCUCLI%(NCXI%) > 0 Then
                GoTo 19       ' descarta repetidos
              End If
              OCUCLI%(NCXI%) = 1
            End If
            '
            CLI$ = Mid$(DEU1$, 3, 30)
            CUIT$ = TRIM$(Mid$(DEU2$, 75, 15))
            If VERICUIT% <> (-1) Then
              If CUIT$ <> "" Then
                If VALICUIT%(CUIT$) <> 1 Then
                  Screen.MousePointer = 1
                  TECLI$ = "Cliente": If ARDEU% = 2 Then TECLI$ = "Proveedor"
                  If COMALTER%(TECLI$ + ": " + TRIM$(CLI$) + "\C.U.I.T.: " + CUIT$ + "\  \Número de C.U.I.T. no Válido.\\Continuar Revisión\Cancelar") = 2 Then
                     VERICUIT% = (-1)
                  End If
                  Screen.MousePointer = 11
                End If
              End If
            End If
            REX$ = Space$(36)
            Call REPLA(REX$, CLI$, 1, 30)
            Call REPLA(REX$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            FORSELEC.LISORT.AddItem REX$
            REGDEU$(U&) = DEU1$ + DEU2$
            '
          ElseIf CVI(Left$(DEU1$, 2)) < 0 Then
            '
            XX$ = DEU1$
            Call COPYSTRU("DEUDOR1", "DEUDOBAJ")
            Call REGAPP("DEUDOBAJ", XX$)
            Call CIERRARCH("DEUDOBAJ")
        End If
19  Next U&
    '
    URE& = 0
    For U& = 1 To FORSELEC.LISORT.ListCount
        FORSELEC.LISORT.ListIndex = U& - 1
        INDI& = Val(Mid$(FORSELEC.LISORT.TEXT, 31, 6))
        DEU1$ = Left$(REGDEU$(INDI&), 32)
        DEU2$ = Mid$(REGDEU$(INDI&), 33, 128)
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
        URE& = U&
    Next U&
    For U& = URE& + 1 To LOF(N1%) / 32
        DEU1$ = String$(32, 0)
        DEU2$ = String$(128, 0)
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
    Next U&
    '
    Unlock #N2%
    Unlock #N1%
    Close #N2%: N2% = 0
    Close #N1%: N1% = 0
    DEU1$ = REGLEIDO$(ASELE$, 1)
    Call GRAREG(ASELE$, DEU1$, 1)
    Call CIERRARCH(ASELE$)
    Screen.MousePointer = 1
    Exit Sub
    '
100 Resume 101
101 On Error GoTo 0
    Close #N1%: N1% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
200 Resume 201
201 On Error GoTo 0
    Close #N1%: N1% = 0
    Close #N2%: N2% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
End Sub
'

Sub GRACREDOR()
     '
1540 J1& = 0: J2& = 0
1550 For I& = 1 To ULTREG&("PROVED1")
        X1$ = REGLEIDO$("PROVED1", I&)
        X2$ = REGLEIDO$("PROVED2", I&)
        TC% = Asc(Mid$(X2$, 119, 1))
1565    If TC% = 0 Then
              J1& = J1& + 1
              Call GRAREG("ACREDOR", X1$, J1&)
            Else
              J2& = J2& + 1
              Call GRAREG("CUEGAST", X1$, J2&)
        End If
1580 Next I&
1590 '
1600 Call SETULTREG("ACREDOR", J1&)
     Call SETULTREG("CUEGAST", J2&)
     Call CIERRARCH("ACREDOR")
     Call CIERRARCH("CUEGAST")
     '
End Sub
'

Sub SELRACLI()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("DEUDOR1")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    FORLICLI.Text10.TEXT = TRIM$(Str$(NCMIN%))
    FORLICLI.Text9.TEXT = RASOCLI$(NCMIN%)
    FORLICLI.Text8.TEXT = TRIM$(Str$(NCMAX%))
    FORLICLI.Text7.TEXT = RASOCLI$(NCMAX%)
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
10 VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
   OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
   If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     Exit Sub
   End If
   '
   NCL1% = CVI(Left$(OBX$, 2))
   NCL2% = CVI(Mid$(OBX$, 3, 2))
   If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
   End If
   VENDSEL% = Asc(Mid$(OBX$, 5, 1))
   '
End Sub
'
Sub GELISCLI()
    '
    Screen.MousePointer = 11
    GRUCLI% = Val(FORLICLI.Text1.TEXT)
    GRUX$ = TRIM$(FORLICLI.Text2.TEXT)
    VENDE% = Val(FORLICLI.Text3.TEXT)
    VENDX$ = TRIM$(FORLICLI.Text4.TEXT)
    NCMIN% = Val(FORLICLI.Text10.TEXT)
    NCMAX% = Val(FORLICLI.Text8.TEXT)
    '
    MODORDE% = 1
    If FORLICLI.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLICLI.Option1(3) = True Then
        MODORDE% = 3
    End If
    '
    FIN& = ULTREG&("DEUDOR1")
    UJ& = 0
    FORLICLI.List1.Clear
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("DEUDOR1", U&)
        NCLIE = CVI(Left$(X$, 2))
        '
        If NCLIE >= NCMIN% Then
          If NCLIE <= NCMAX% Then
             If GRUCLI% = 0 Or GRUCOCLI%(NCLIE) = GRUCLI% Then
                If VENDE% = 0 Or VENDCLI%(NCLIE) = VENDE% Then
                  '
                  TLISTA$ = Space$(36)
                  If MODORDE% = 1 Then
                       Call REPLA(TLISTA$, CSTRING$(MKI$(NCLIE), 1, 6, 0, 2), 1, 6)
                    ElseIf MODORDE% = 2 Then
                       Call REPLA(TLISTA$, RASOCLI$(NCLIE), 1, 30)
                    ElseIf MODORDE% = 3 Then
                       Call REPLA(TLISTA$, "A" + AJUSTD$(CPOSCLI$(NCLIE), 8), 1, 9)
                  End If
                  Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                  FORLICLI.List1.AddItem TLISTA$
                  '
                End If
             End If
          End If
        End If
        '
    Next U&
    '
    ULI& = 0
    FIN& = FORLICLI.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        FORLICLI.List1.ListIndex = U& - 1
        REG& = Val(Mid$(FORLICLI.List1.TEXT, 31, 8))
        X$ = REGLEIDO$("DEUDOR1", REG&)
        NCLIE = CVI(Left$(X$, 2))
        '
        If NCLIE > 0 Then
          If InStr(EPAC$, "NEUMANN") > 0 Then
            If TRIM$(CUITCLI$(NCLIE)) = "" Then GoTo 99
          End If
        End If
        '
        RACLI$ = RASOCLI$(NCLIE)
        On Error Resume Next
        FORLICLI.Label3.Caption = RACLI$
        FORLICLI.Label3.Refresh
        On Error GoTo 0
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NCLIE), 1, 2)
        Call REPLA(Y$, RACLI$, 3, 32)
        Call REPLA(Y$, TEL1CLI$(NCLIE), 35, 16)
        Call REPLA(Y$, CUITCLI$(NCLIE), 51, 16)
        Call REPLA(Y$, POSIVAC$(NCLIE), 67, 10)
        CUECOX% = GRUCOCLI%(NCLIE)
        Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
        Call REPLA(Y$, ECOARCH$("GRUCOCLI", MKI$(CUECOX%)), 79, 32)
        Call REPLA(Y$, CUEMADRE$(NCLIE), 111, 16)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NCLIE), 1, 2)
        Call REPLA(Y$, DOMICLI$(NCLIE), 3, 32)
        'Call REPLA(Y$, FAXCLI$(NCLIE), 35, 16)
        LPRX% = LIPRCLI%(NCLIE)
        Call REPLA(Y$, MKI$(LPRX%), 77, 2)
        DDLX$ = TRIM$(DELIPRE$(LPRX%))
        DCCL$ = TRIM$(DESCUCLI$(NCLIE))
        If DCCL$ <> "" Then
           If InStr(DCCL$, "%") < 1 Then
              DCCL$ = DCCL$ + " %"
           End If
           DDLX$ = Left$(DDLX$, 29 - Len(DCCL$))
           DDLX$ = DDLX$ + " (" + DCCL$ + ")"
        End If
        Call REPLA(Y$, DDLX$, 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NCLIE), 1, 2)
        DOMI$ = TRIM$(CPOSCLI$(NCLIE))
        If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
        DOMI$ = DOMI$ + LOCACLI$(NCLIE)
        Call REPLA(Y$, DOMI$, 3, 32)
        'Call REPLA(Y$, EMAICLI$(NCLIE), 35, 16)
        Call REPLA(Y$, NIBRCLI$(NCLIE), 51, 16)
        CTRX$ = TRIM$(CATIBRU$(NCLIE))
        CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                             If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                             If CTRX$ = "2" Then CATIBX$ = "C.Multil."
        Call REPLA(Y$, CATIBX$, 67, 10)
        CPAG% = CPAGCLI%(NCLIE)
        Call REPLA(Y$, MKI$(CPAG%), 77, 2)
        Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        Call REPLA(Y$, MKI$(NCLIE), 1, 2)
        VDX% = VENDCLI%(NCLIE)
        Call REPLA(Y$, MKI$(VDX%), 77, 2)
        Call REPLA(Y$, DEVENDE$(CInt(NCLIE)), 79, 32)
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
        Y$ = REGBLAN$("LICLIEN")
        ULI& = ULI& + 1
        Call GRAREG("LICLIEN", Y$, ULI&)
        '
99  Next U&
    '
    Call SETULTREG("LICLIEN", ULI&)
    Call CIERRARCH("LICLIEN")
    Screen.MousePointer = 1
    Call FINAL("*LICLIEN")
    '
End Sub
'

Sub SELRAPRO()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    FIN& = ULTREG&("PROVED1")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("PROVED1", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
    Next U&
    FORLIPROVE.Text10.TEXT = TRIM$(Str$(NCMIN%))
    FORLIPROVE.Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    FORLIPROVE.Text8.TEXT = TRIM$(Str$(NCMAX%))
    FORLIPROVE.Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub GELISPRO()
    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    DETICUE$(0) = "Cuenta Corriente"
    DETICUE$(1) = "Gastos"
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            TICUE%(NCX%) = 0
        End If
    Next U&
    '
    NCMIN% = Val(FORLIPROVE.Text10.TEXT)
    NCMAX% = Val(FORLIPROVE.Text8.TEXT)
    '
    TICUENTA% = 2
    If FORLIPROVE.Option2(0).Value = True Then
         TICUENTA% = 0
       ElseIf FORLIPROVE.Option2(1).Value = True Then
         TICUENTA% = 1
    End If
    '
    MODORDE% = 1
    If FORLIPROVE.Option1(2) = True Then
        MODORDE% = 2
      ElseIf FORLIPROVE.Option1(3) = True Then
        MODORDE% = 3
    End If
    '
    FIN& = ULTREG&("PROVED1")
    UJ& = 0
    FORLIPROVE.List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NCLIE = CVI(Left$(X$, 2))
        '
        If NCLIE >= NCMIN% Then
          If NCLIE <= NCMAX% Then
             If TICUENTA% = 2 Or TICUE%(NCLIE) = TICUENTA% Then
                '
                TLISTA$ = Space$(36)
                If MODORDE% = 1 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(NCLIE), 1, 6, 0, 2), 1, 6)
                  ElseIf MODORDE% = 2 Then
                     Call REPLA(TLISTA$, RASOCLI$((-1) * NCLIE), 1, 30)
                  ElseIf MODORDE% = 3 Then
                     Call REPLA(TLISTA$, "A" + AJUSTD$(CPOSCLI$((-1) * NCLIE), 8), 1, 9)
                End If
                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                FORLIPROVE.List1.AddItem TLISTA$
                '
             End If
          End If
        End If
        '
    Next U&
    '
    ULI& = 0
    FIN& = FORLIPROVE.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        FORLIPROVE.List1.ListIndex = U& - 1
        REG& = Val(Mid$(FORLIPROVE.List1.TEXT, 31, 8))
        X$ = REGLEIDO$("PROVED1", REG&)
        NP% = CVI(Left$(X$, 2))
        NCLIE = (-1) * NP%
        '
        Y$ = REGBLAN$("LIPROVE")
        Call REPLA(Y$, MKI$(NP%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NCLIE), 3, 32)
        Call REPLA(Y$, TEL1CLI$(NCLIE), 35, 16)
        Call REPLA(Y$, CUITCLI$(NCLIE), 51, 16)
        Call REPLA(Y$, POSIVAC$(NCLIE), 67, 10)
        CUECOX% = TICUE%(NP%)
        Call REPLA(Y$, MKI$(CUECOX%), 77, 2)
        Call REPLA(Y$, DETICUE$(CUECOX%), 79, 32)
        Call REPLA(Y$, CUEMADRE$(NCLIE), 111, 16)
        ULI& = ULI& + 1
        Call GRAREG("LIPROVE", Y$, ULI&)
        '
        If TICUE%(NP%) = 0 Then
           Y$ = REGBLAN$("LIPROVE")
           Call REPLA(Y$, MKI$(NP%), 1, 2)
           Call REPLA(Y$, DOMICLI$(NCLIE), 3, 32)
           'Call REPLA(Y$, FAXCLI$(NCLIE), 35, 16)
           Call REPLA(Y$, NIBRCLI$(NCLIE), 51, 16)
           CTRX$ = TRIM$(CATIBRU$(NCLIE))
           CATIBX$ = "No Def.": If CTRX$ = "0" Then CATIBX$ = "No Resp."
                                If CTRX$ = "1" Then CATIBX$ = "Cont.Puro"
                                If CTRX$ = "2" Then CATIBX$ = "C.Multil."
           Call REPLA(Y$, CATIBX$, 67, 10)
           CPAG% = CPAGCLI%(NCLIE)
           Call REPLA(Y$, MKI$(CPAG%), 77, 2)
           Call REPLA(Y$, DECONPA$(CPAG%), 79, 32)
           ULI& = ULI& + 1
           Call GRAREG("LIPROVE", Y$, ULI&)
           '
           Y$ = REGBLAN$("LIPROVE")
           Call REPLA(Y$, MKI$(NP%), 1, 2)
           DOMI$ = TRIM$(CPOSCLI$(NCLIE))
           If DOMI$ <> "" Then DOMI$ = DOMI$ + " - "
           DOMI$ = DOMI$ + LOCACLI$(NCLIE)
           Call REPLA(Y$, DOMI$, 3, 32)
           'Call REPLA(Y$, EMAICLI$(NCLIE), 35, 16)
           ULI& = ULI& + 1
           Call GRAREG("LIPROVE", Y$, ULI&)
        End If
        '
        Y$ = REGBLAN$("LIPROVE")
        ULI& = ULI& + 1
        Call GRAREG("LIPROVE", Y$, ULI&)
        '
    Next U&
    '
    Call SETULTREG("LIPROVE", ULI&)
    Call CIERRARCH("LIPROVE")
    Screen.MousePointer = 1
    Call FINAL("*LIPROVE")
    '
End Sub
'

Sub GELIGEPRO()
    '
    Dim TICUE%(32767)
    Dim DETICUE$(2)
    Screen.MousePointer = 11
    '
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    DETICUE$(0) = "Cuenta Corriente"
    DETICUE$(1) = "Gastos"
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            TICUE%(NCX%) = 0
        End If
    Next U&
    '
    FIN& = ULTREG&("PROVED1")
    ULI& = 0
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NP% = CVI(Left$(X$, 2))
        NCLIE = (-1) * NP%
        '
        Y$ = REGBLAN$("LIGEPRO")
        Call REPLA(Y$, MKI$(NP%), 1, 2)
        CUECOX% = TICUE%(NP%)
        Call REPLA(Y$, DETICUE$(CUECOX%), 3, 30)
        ULI& = ULI& + 1
        Call GRAREG("LIGEPRO", Y$, ULI&)
        '
    Next U&
    '
    Call SETULTREG("LIGEPRO", ULI&)
    Call CIERRARCH("LIGEPRO")
    Screen.MousePointer = 1
    Call FINAL("*LIGEPRO")
    '
End Sub
'

Sub ACTUDESC()
    '
    LDT$ = LUDAT$
    If EXISTE%(LDT$ + "DESCUCLI.DAT") < 1 Then
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("DEUDOR1")
            X$ = REGLEIDO$("DEUDOR1", U&)
            X$ = X$ + REGLEIDO$("DEUDOR2", U&)
            PDX1$ = TRIM$(Str$(CVS(Mid$(X$, 153, 4))))
            If Val(PDX1$) >= 1 And Val(PDX1$) <= 72 Then MODO% = 1: GoTo 90
            PDX0$ = TRIM$(Mid$(X$, 122, 14))
            PDX1$ = REPLACAR$(PDX0$, "%", " ")
            If InStr(PDX1$, "+") > 0 Then MODO% = 2: GoTo 90
            If Len(PDX1$) < 6 And Val(PDX1$) >= 1 Then MODO% = 2: GoTo 90
            GoTo 99
            '
90          NCXII% = CVI(Left$(X$, 2))
            Y$ = REGBLAN$("DESCUCLI")
            Call REPLA(Y$, MKI$(NCXII%), 1, 2)
            Call REPLA(Y$, PDX1$, 3, 14)
            Call REGAPP("DESCUCLI", Y$)
            '
            X1$ = REGLEIDO$("DEUDOR2", U&)
            If MODO% = 2 Then Call REPLA(X1$, Space$(14), 90, 14)
            If MODO% = 1 Then Call REPLA(X1$, MKS$(0), 121, 4)
            Call GRAREG("DEUDOR2", X1$, U&)
            '
99      Next U&
        Call CIERRARCH("DEUDOR1")
        Call CIERRARCH("DEUDOR2")
        Call REGAPP("DESCUCLI", String$(16, 0))
        Call CIERRARCH("DESCUCLI")
        Screen.MousePointer = 1
        '
    End If
    '
End Sub
'

Sub VALDATCLI(NCXI%, OK2%)
   '
   OK2% = 0
   Call BLOQARCH("DEUDOR1")
   Call BLOQARCH("DEUDOR2")
   REG& = REGICLI&(NCXI%)
   If REG& < 1 Or REG& > ULTREG&("DEUDOR1") Then
      Exit Sub
   End If
   XX0$ = REGLEIDO$("DEUDOR1", REG&) + REGLEIDO$("DEUDOR2", REG&)
   OBX$ = OBJPLA$("CARDATCLIEN", XX0$)
   If OBX$ <> "" Then
      XX1$ = Left$(OBX$, 32)
      XX2$ = Mid$(OBX$, 33)
      Call GRAREG("DEUDOR1", XX1$, REG&)
      Call GRAREG("DEUDOR2", XX2$, REG&)
      Call CIERRARCH("DEUDOR1")
      Call CIERRARCH("DEUDOR2")
      OK2% = 1
   End If
   Call LIBARCH("DEUDOR1")
   Call LIBARCH("DEUDOR2")
   '
End Sub
'
Function CUEIBRUPRO$(A$)
   '
   Static CTX%, COPROVI$, CUEPROVI$(64)
   If CTX% = 0 Then
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVI$(Len(COPROVI$)) = Mid$(ECOARCH$("CUECON", Mid$(X$, 27, 2)), 31, 12)
          End If
       End If
     Next I&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQQ$ = Space$(12)
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQQ$ = CUEPROVI$(PPXX%)
   End If
   '
   CUEIBRUPRO$ = AJUSTI$(QQQ$, 12)
   '
End Function

Function CUEIPROVIN%(A$)
   '
   Static CTX%, COPROVI$, CUEPROVII%(64)
   If CTX% = 0 Then
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVII%(Len(COPROVI$)) = CVI(Mid$(X$, 27, 2))
          End If
       End If
     Next I&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQI% = 0
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQI% = CUEPROVII%(PPXX%)
   End If
   '
   CUEIPROVIN% = QQI%
   '
End Function

Function CUEMECOB%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(CONTROL$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MECOBRA", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(CUCU$)
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMECOB% = CMC%
   '
End Function

Function CUEMEPAG%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(CONTROL$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MEPAGO", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       XX$ = REGLEIDO$("MEPAGO", REME&)
       CMC% = CVI(Mid$(XX$, 61, 2))
       If CMC% < 1 Then
         CUCU$ = Mid$(XX$, 45, 12)
         CMC% = CUECOINT%(CUCU$)
       End If
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMEPAG% = CMC%
   '
End Function

Function CUECOINT%(CUEX$)
   '
   CCIT% = 0
   RFF$ = RECFILT$("ECUECON", 1, CUEX$)
   If Len(RFF$) >= 4 Then
     RECUE& = CVS(Left$(RFF$, 4))
     XXY$ = REGLEIDO$("ECUECON", RECUE&)
     CCIT% = CVI(Mid$(XXY$, 41, 2))
   End If
   '
   If CCIT% = 8224 Then
     RFF$ = RECFILT$("INVECUE", 1, CUEX$)
     If Len(RFF$) >= 4 Then
       RECUE& = CVS(Left$(RFF$, 4))
       XXY$ = REGLEIDO$("INVECUE", RECUE&)
       CCIT% = CVI(Mid$(XXY$, 13, 2))
     End If
   End If
   '
   CUECOINT% = CCIT%
   '
End Function

Function CUECARCHE$()
   '
   CUECARCHE$ = CONTROL$("", "CUECHECA")
   CUEICHEQ% = CUECOINT%(CUECARCHE$)
   If TRIM$(CUECARCHE$) = "" Or CUEICHEQ% < 1 Then
     CUECARCHE$ = ""
     Exit Function
   End If
   '
End Function

Function CUEICARCH%()
   '
   If NUCUEN% < 1 Then NUCUEN% = ULTREG&("CUENTAS")
   CUECHECA$ = CONTROL$("", "CUECHECA")
   CUEICARCH% = CUECOINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICARCH% < 1 Or CUEICARCH% > NUCUEN% Then
     CUEICARCH% = 0
     Exit Function
   End If
   '
End Function

Function MONEMICO%(NPROX%)
   '
   MNCLIE = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNCLIE = Asc(Mid$(XX$, 31, 1))
   End If
   '
   MONEMICO% = MNCLIE
   '
End Function

Function MONEVACO%(NPROX%)
   '
   MNCLIE = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNCLIE = Asc(Mid$(XX$, 32, 1))
   End If
   '
   MONEVACO% = MNCLIE
   '
End Function

Function ALIRETBRU#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
    Static PROVANT%, ALIRETBRUANT#, MINOIMPANT#, MEPAGOANT%
    
    If NPROVE% > 0 And NPROVE% = PROVANT% Then GoTo 99
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    For U& = 1 To ULTREG&("TACATIBP")
       TCX$ = REGLEIDO$("TACATIBP", U&)
       TCY% = Asc(Mid$(TCX$, 1, 1))
       If TCY% > 0 Then
          If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             If ALIBRU(TCY%) > 0 Then CTCALCU% = 1
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
          End If
       End If
    Next U&
    '
    REBLA$ = REGBLAN$("DARIBPRO")
    RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(Abs(NPROVE%)))
    CATIB1% = 1: PORCONVE = 100   ' VALORES POR DEFAULT
    If RERET$ <> REBLA$ Then
        CATIB1% = Asc(Mid$(RERET$, 3, 1))
        PORCONVE = CVS(Mid$(RERET$, 5, 4))
    End If
    RETENC# = ALIBRU(CATIB1%)
    MINOMPONIBLE# = MINIMPO#(CATIB1%)
    MEDIOPAGO% = MEPAGX%(CATIB1%)
    ' falta sacar el medio de pago
    '
    If EXISTE%(LUDAT$ + "PADIBRU.TXT") > 0 Then
       CUITPROV$ = CUITCLI$((-1) * Abs(NPROVE%))
       CUITPROVEEDOR$ = REPLACAR$(CUITPROV$, "-", "")
       '
       NUMERODEARCHIVO% = FILEOPEN%(LUDAT$ + "PADIBRU.TXT", 1, 128)
       Do While Not EOF(NUMERODEARCHIVO%)
          Line Input #NUMERODEARCHIVO%, XX$
          CUITPADRON$ = Mid$(XX$, 28, 11)
          If CUITPADRON$ = CUITPROVEEDOR$ Then
             RETU$ = Mid$(XX$, 51, 4)
             RETENC# = XVALO(RETU$)
             Exit Do
          End If
       Loop
    End If
    '
    PROVANT% = NPROVE%
    ALIRETBRUANT# = RETENC#
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function

Function ALIPERBRU#(NCLIE%, MINIPER#, VARDO%)
    '
    Static CLIEANT%, ALIPERBRUANT#, MINIPERANT#
    Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPERCE
    
    If NCLIE% > 0 And NCLIE% = CLIEANT% Then GoTo 99
    If Abs(NCLIE%) = 0 Then
        ALIPERBRUANT# = 0: MINIPERANT# = 0: CLIEANT% = 0
        GoTo 99
    End If
    '
    If CTTXX% < 1 Then
       TASAA = XVALO(CONTROL$("PERIBRU", "TASA-A"))
       TASAB = XVALO(CONTROL$("PERIBRU", "TASA-B"))
       TASAEXP = XVALO(CONTROL$("PERIBRU", "TASA-EXP"))
       MINIPERCE = XVALO(CONTROL$("PERIBRU", "MINIPER"))
       CTTXX% = 1
    End If
    '
    TASAX = TASAA
    If VARDO% = 2 Then TASAX = TASAB
    If VARDO% = 3 Then TASAX = TASAEXP
    '
    REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NCLIE%))
    CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
    PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
    If PORCONVE > 100 Then PORCONVE = 100
    If UCase$(Mid$(REPERCI$, 9, 18)) = "EXENTO CERTIFICADO" Then
       TASAX = 0
       GoTo 98
    End If
    '
    If CATIB1% > 0 Then         ' NO ESTA EXENTO
      If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
           TASAX = TASAX * PORCONVE / 100
         Else
           TASAX = 0
      End If
    End If
    If CATIB1% < 1 Then TASAX = 0
    '
    If EXISTE%(LUDAT$ + "PADIBRU.TXT") > 0 Then
       CUITCLIEN$ = CUITCLI$(NCLIE%)
       If CUITCLIEN$ = "" Then GoTo 98
       '
       CUITCLIENTE$ = REPLACAR$(CUITCLIEN$, "-", "")
       '
       NUMERODEARCHIVO% = FILEOPEN%(LUDAT$ + "PADIBRU.TXT", 1, 128)
       Do While Not EOF(NUMERODEARCHIVO%)
          Line Input #NUMERODEARCHIVO%, XX$
          CUITPADRON$ = Mid$(XX$, 28, 11)
          If CUITPADRON$ = CUITCLIENTE$ Then
             PER$ = Mid$(XX$, 46, 4)
             TASAX = XVALO(PER$)
             Exit Do
          End If
       Loop
    End If
    '
98  CLIEANT% = NCLIE%
    ALIPERBRUANT# = TASAX
    MINIPERANT# = MINIPERCE
    '
99  ALIPERBRU# = ALIPERBRUANT#
    MINIPER# = MINIPERANT#
    '
End Function
Function CONTROL$(PROCE$, CLAVE$)
    '
    PROBUS$ = AJUSTI$(PROCE$, 8)
    If TRIM$(PROCE$) = "" Then PROBUS$ = ""
    CODBUS$ = AJUSTI$(CLAVE$, 8)
    '
    CONTROL$ = ""
    RFF$ = RECFILT$("FLEXCRL", 2, CODBUS$)
    For U& = 1 To Len(RFF$) Step 4
      REBU& = CVS(Mid$(RFF$, U&, 4))
      XI$ = REGLEIDO$("FLEXCRL", REBU&)
      If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
        CONTROL$ = TRIM$(Mid$(XI$, 17, 80))
        Exit For
      End If
    Next U&
    '
End Function
'
Sub GRACONTROL(PROCE$, CLAVE$, VALCON$)
    '
    PROBUS$ = TRIM$(UCase$(PROCE$))
    CODBUS$ = TRIM$(UCase$(CLAVE$))
    '
    For I& = 1 To ULTREG&("FLEXCRL")
       XI$ = REGLEIDO$("FLEXCRL", I&)
       If TRIM$(UCase$(Left$(XI$, 8))) = PROBUS$ Then
           If TRIM$(UCase$(Mid$(XI$, 9, 8))) = CODBUS$ Then
               Call REPLA(XI$, VALCON$, 17, 80)
               Call GRAREG("FLEXCRL", XI$, I&)
               Call CIERRARCH("FLEXCRL")
               Exit Sub
           End If
       End If
    Next I&
    '
    XI$ = REGBLAN$("FLEXCRL")
    Call REPLA(XI$, PROBUS$, 1, 8)
    Call REPLA(XI$, CODBUS$, 9, 8)
    Call REPLA(XI$, VALCON$, 17, 80)
    Call REGAPP("FLEXCRL", XI$)
    Call CIERRARCH("FLEXCRL")
    '
End Sub

Sub MUESTRADOC(DOCUNU&, Optional MODOX$)
    '
    XXX$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
    EDITFORM.Label98 = TRIM$(Mid$(XXX$, 5, 60))
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    
    Dim TD0 As String * 512
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
      Close #NX0%
      NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    End If
    '
    URE& = Abs(LOF(NX0%) / 512)
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   tpsp$ = ""
    Screen.MousePointer = 11
    While L& > 1
      Get #NX0%, L&, TD0$
      DCC& = CVS(Left$(TD0$, 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      Get #NX0%, KKL&, TD0$
      TBUF$ = TD0$
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        tpsp$ = tpsp$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  Close #NX0%
    SPOOLSTRING$ = tpsp$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(tpsp$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(tpsp$, PPXX% - 1)
      tpsp$ = Mid$(tpsp$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(tpsp$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command1.Enabled = False
    '
    On Error Resume Next
    If UCase$(MODOX$) = "PRINT" Then
        Call EDITFORM.Command4_Click
      Else
        EDITFORM.Show 1
    End If
    On Error GoTo 0

99  End Sub
Sub SAVESPOOL(DESCRIDOC$, OKX%)
   '
   Static CACOMM%, PRICOMM$(128), COMMINDEX$
   OKX% = 0
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
      COMMINDEX$ = ""
      For KKII% = 1 To CACOMM%
        COMMINDEX$ = COMMINDEX$ + AJUSTI$(PRICOMM$(KKII%), 16)
      Next KKII%
   End If
   '
   'If FORMALTER.PSPOOL.ListCount < 1 Then
'MsgBox "Entré en SAVESPOOL con CAREVEC = " & CAREVEC&
   If CAREVEC& < 1 Then
      Exit Sub
   End If
   '
'10 FIN& = FORMALTER.PSPOOL.ListCount
10 FIN& = CAREVEC&
   tpsp$ = String$(4096, 0): PINSE& = 1
   For KKIL& = 1 To FIN&
      Call TRACE(20, KKIL&, FIN&)
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKIL& - 1))
      TTTT$ = TRIM$(VECPRICO$(KKIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      KKKK% = InStr(COMMINDEX$, PCOMM$)
      If (KKKK% Mod 16) = 1 Then
        KKJJ% = (KKKK% + 15) / 16
        TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
        If PINSE& + Len(TPSQ$) > Len(tpsp$) Then
          tpsp$ = tpsp$ + String$(4096, 0)
        End If
        Mid$(tpsp$, PINSE&, Len(TPSQ$)) = TPSQ$
        PINSE& = PINSE& + Len(TPSQ$)
        '+ Chr$(KKJJ%) + PATRI$ + Chr$(255)
        GoTo 19
      End If
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
           If PINSE& + Len(TPSQ$) > Len(tpsp$) Then
             tpsp$ = tpsp$ + String$(4096, 0)
           End If
           Mid$(tpsp$, PINSE&, Len(TPSQ$)) = TPSQ$
           PINSE& = PINSE& + Len(TPSQ$)
           'TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           'If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKIL&
   tpsp$ = Left$(tpsp$, PINSE& - 1)
   '
'MsgBox "Armado el Vector de Transferencia."
20 On Error Resume Next
   Call BLOQARCH("PDOCLST")
   If Err Then GoTo 21
   Call BLOQARCH("PDOCSPL")
   If Err Then
21    On Error GoTo 0
      Call COMUNI("Error al Bloquear Tablas 'PDOCLST' y 'PDOCSPL'.\  \Espere un Instante y Pulse 'Aceptar'\para Reanudar Ejecución.")
      GoTo 20
   End If
   On Error GoTo 0
   '
   UDOCU& = 0
   UREGI& = ULTREG&("PDOCLST")
   If UREGI& > 0 Then
      UDOCU& = CVS(Left$(REGLEIDO$("PDOCLST", UREGI&), 4))
   End If
   PDLS$ = String$(256, 0)
   DOCUNU& = 1 + UDOCU&
   HOII% = HOY(HOS$)
   Call REPLA(PDLS$, MKS$(DOCUNU&), 1, 4)
   Call REPLA(PDLS$, DESCRIDOC$, 5, 60)
   Call REPLA(PDLS$, MKI$(HOII%), 65, 2)
   Call REPLA(PDLS$, MKI$(Int(HORANUM(Time$))), 67, 2)
   Call REPLA(PDLS$, USERNAME$, 69, 30)
   Call REPLA(PDLS$, NOMTER$, 99, 30)
     On Error Resume Next
   Call REPLA(PDLS$, Printer.DeviceName, 129, 128)
     On Error GoTo 0
   Call REGAPP("PDOCLST", PDLS$)
   Call CIERRARCH("PDOCLST")
   '
'MsgBox "Guardado Encabezado de Documento"
   KKJJ% = 0
   FIN1& = Len(tpsp$): QCOUN& = 0
   While Len(tpsp$) > 0
     QCOUN& = QCOUN& + 1
     Call TRACE(20, QCOUN&, FIN1&)
     TBUF$ = MKS$(DOCUNU&) + Left$(tpsp$ + String$(508, 0), 508)
     Call REGAPP("PDOCSPL", TBUF$)
     tpsp$ = Mid$(tpsp$, 509)
   Wend
   Call CIERRARCH("PDOCSPL")
   '
'MsgBox "Guardado Cuerpo del Documento"
   Call CIERRARCH("PIMGLST")
   Call CIERRARCH("PIMGSPL")
   '
30 On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   On Error GoTo 0
   If MATRIXPUN% = 1 Then
      FORMALTER.List1.Clear
   End If
   '
   If CANCELPRINT% < 1 Then
     If HOJASUEL% = 1 Then
       Call COMUNI("Coloque Formulario Correspondiente a\" + TIPODOC$ + " " + NUMEDOC$ + "\y Pulse Botón 'Aceptar Mensaje'.")
     End If
   End If
   '
   If SENDBYMAIL$ <> "" Then
     IMPREDET$ = Printer.DeviceName
     Printer.TrackDefault = False
     For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then
         Printer.EndDoc
         On Error Resume Next
         Set Printer = PRX1
         If Err Then
           On Error GoTo 0
           GoTo 32
         End If
         On Error GoTo 0
         ARANEX$ = DEDOCOR$
         If ARANEX$ = "" Then ARANEX$ = "PDF"
         On Error Resume Next
         Kill LUCOM$ + "" + ARANEX$ + ".PDF"
         On Error GoTo 0
         EDITFORM.Label1 = ARANEX$
         EDITFORM.Timer1.Enabled = True
         GoTo 35
       End If
     Next
     '
32   Call MENSERR(24, "Imposible Direccionar a PDFWriter")
   End If
   '
35 On Error Resume Next
'MsgBox "Comienza la Impresión"
   Printer.EndDoc   ' INICIALIZA LA IMPRESORA
   Printer.ScaleMode = 1
        'MsgBox "SCALEMODE: " & Printer.ScaleMode
        'MsgBox "SCALEWIDTH: " & Printer.ScaleWidth
        'MsgBox "SCALELEFT: " & Printer.ScaleLeft
        'MsgBox "SCALEHEIGHT: " & Printer.ScaleHeight
        'MsgBox "SCALETOP: " & Printer.ScaleTop
        'MsgBox "PAPERSIZE: " & Printer.PaperSize
   If Err Then
     Resume
     On Error GoTo 0
     If COMALTER%("Error de Comunicacion con la Impresora\\Resolver y Reintentar\Cancelar Impresion") <> 2 Then
       GoTo 35
     End If
     GoTo 40
   End If
   '
   FUERADEAREA% = 0
   FUENODISPO% = 0
   'For KKII% = 1 To FORMALTER.PSPOOL.ListCount
   For KKIIL& = 1 To CAREVEC&
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      Call TRACE(20, KKIIL&, CAREVEC&)
      TTTT$ = TRIM$(VECPRICO$(KKIIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
               TTXX$ = TTTT$ + Chr$(10) + Chr$(13)
               TTXX$ = TTXX$ + "SCALEMODE: " & Printer.ScaleMode & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEWIDTH: " & Printer.ScaleWidth & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALELEFT: " & Printer.ScaleLeft & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEHEIGHT: " & Printer.ScaleHeight & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALETOP: " & Printer.ScaleTop & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "PAPERSIZE: " & Printer.PaperSize
               'MsgBox TTXX$
           GoTo 39
        End If
      Next KKJJ%
39 Next KKIIL&
   '
'MsgBox "Termina la Impresión"
   If FUERADEAREA% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: Se Han Encontrado Leyendas\Fuera de Area de Impresión Horizontal o Vertical.\   \FLEXOFT ha Suprimido estas Leyendas.\  \Se Recomienda Revisar y Corregir el Tamaño de Papel\Configurado en Esta Estación de Trabajo\para la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".")
   End If
   FUERADEAREA% = 0
   '
   If FUENODISPO% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: El Formulario Electrónico Actual\Utiliza Fuentes que no Están Disponibles\en la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".\   \Se Recomienda Revisar el Formulario Electrónico\o Utilizar Otra Impresora de la Red.")
   End If
   FUENODISPO% = 0
   '
'40 FORMALTER.PSPOOL.Clear
40 ReDim VECPRICO$(0)
   CAREVEC& = 0
   Call CLOSEFORMPRIN("")
   '
   If SENDBYMAIL$ <> "" Then
     TINI = Timer
     
50   If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then
       If Timer < TINI + 10 Then
         GoTo 50
       End If
       '
       Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
       Exit Sub
     End If
     Call ENVIAPORMAIL
   End If
   OKX% = 1
   '
End Sub


