Attribute VB_Name = "MIGRACLIEN_DAT"
Public PUCLI(1) As String
Public PUBAC(1) As String


Sub DATCLIEN(NC%, VDEV$, OP%)
'
   Static REBLA$(1)
   Static CAPOS%, PIVA$(), PIVA2$(), CTX%, REG&
   Static NCA%, TIBUA%, TANTU, XX$, XBA$, XBC$
   Static SIGANT%
   Static URECUEN&
   '
   VDEV$ = ""
   If CTX% = 0 Or Sgn(NC%) <> SIGANT% Then
       URECUEN& = ULTREG&("CUENTAS")
       ARX1$ = "DEUDOR1"
       If NC% < 0 Then
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
'       Call ACTUDESC           ' ACTUALIZAR TABLA DE DESCUENTOS
       CTX% = 1
   End If
   '
   SIGANT% = Sgn(NC%)
   If Abs(NC%) < 1 Then
       If OP% <> 0 Then
           Call MENSERR(24, "Numero de Cuenta Incorrecto Invocando Rutina 'DATCLIEN'.")
       End If
       Exit Sub
   End If
   '
   TIBU% = 0: If NC% < 0 Then TIBU% = 1
   If TIBU% = TIBUA% Then
     If NC% = NCA% Then
       If Timer < TANTU + 1 Then
         GoTo SELECTAR
       End If
     End If
   End If
   '
   ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2": ARX3$ = "BADACLIE": TIBU% = 0
   If NC% < 0 Then
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
3  ABUS$ = MKI$(Abs(NC%))
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
      '        If CVI(Left$(XXYY$, 2)) = NC% Then
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
23 NCA% = NC%: TIBUA% = TIBU%: TANTU = Timer
   '
   If OP% = 12 Then GoTo SELECTAR
   XBC$ = ""
   ABADA$ = "BADACLIE": If NC% < 0 Then ABADA$ = "BADAPROV"
   If Abs(NC%) > 0 Then
      RFF$ = RECFILT$(ABADA$, 1, MKI$(Abs(NC%)))
      REBACLI& = 0
      If Len(RFF$) >= 4 Then
         REBACLI& = CVS(Left$(RFF$, 4))
      End If
      If REBACLI& > 0 Then
         If REBACLI& <= ULTREG&(ABADA$) Then
            XBC$ = REGLEIDO$(ABADA$, REBACLI&)
            If CVI(Left$(XBC$, 2)) <> Abs(NC%) Then
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
         VDSS$ = Str$(NC%)
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
         If NC% < 0 Then VDSS$ = Str$(Asc(Mid$(XX$, 152, 1)))
      Case 18
         VDSS$ = "N/Def."
         VD% = Asc(Mid$(XX$, 136, 1))
         If NC% < 0 Then VD% = Asc(Mid$(XX$, 152, 1))
         If VD% >= 0 Then If VD% <= CAPOS% Then VDSS$ = PIVA$(VD%)
      Case 19
         VDSS$ = "Pos. IVA no Definida"
         VD% = Asc(Mid$(XX$, 136, 1))
         If NC% < 0 Then VD% = Asc(Mid$(XX$, 152, 1))
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
             CQXX$ = "AMACLI": If NC% < 0 Then CQXX$ = "AMAPRO"
             VD% = CUECOINT%(CONTROL$(CQXX$, "ECUECON"))
           End If
           VDSS$ = Str$(VD%)
      Case 25
           VDL& = CVI(Mid$(XX$, 159, 2))
           NCUECOI& = ULTREG&("CUENTAS")
           If VDL& > 0 And VDL& <= NCUECOI& Then
               VDSS$ = Left$(REGLEIDO$("CUENTAS", VDL&), 12)
             Else
               CQXX$ = "AMACLI": If NC% < 0 Then CQXX$ = "AMAPRO"
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
Function GRUCOCLI1%(NC%)
   Call DATCLIEN(NC%, VDEV$, 26)
   GRUCOCLI1% = Val(VDEV$)
End Function

Function ZONACLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 33)
   ZONACLI1$ = VDEV$
End Function

Function HORAENT1$(NC%)
   '
   HORAENT1$ = ""
   If NC% > 0 Then
     XX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(NC%))
     HORAENT1$ = TRIM$(Mid$(XX$, 101, 24))
   End If
   '
End Function
Function TEL2CLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 8)
   TEL2CLI1$ = VDEV$
End Function

Function DOMICLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 3)
   DOMICLI1$ = VDEV$
End Function

Function CPOSCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 5)
   CPOSCLI1$ = VDEV$
End Function
Function LOCACLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 6)
   LOCACLI1$ = VDEV$
End Function
Function LIMICRE1(NC%)
   Call DATCLIEN(NC%, VDEV$, 23)
   LIMICRE1 = Val(VDEV$)
End Function

Function PAISCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 27)
   PAISCLI1$ = VDEV$
End Function
Function TEL1CLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 7)
   TEL1CLI1$ = VDEV$
End Function

Function FAXCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 28)
   FAXCLI1$ = VDEV$
End Function
Function EMAILCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 29)
   EMAILCLI1$ = VDEV$
End Function
Function CONTACTCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 30)
   CONTACTCLI1$ = VDEV$
End Function
Function HORARCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 32)
   HORARCLI1$ = VDEV$
End Function
Function CUITCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 9)
   CUITCLI1$ = VDEV$
End Function

'

Function DESCUCLI1$(NC%)
   DSCX$ = ""
   For U& = 1 To ULTREG&("DESCUCLI")
       X$ = REGLEIDO$("DESCUCLI", U&)
       If CVI(Left$(X$, 2)) = NC% Then
           PDX0$ = TRIM$(Mid$(X$, 3, 14))
           DSCX$ = REPLACAR$(PDX0$, "%", " ")
           GoTo 99
       End If
   Next U&
   '
99 DESCUCLI1$ = DSCX$
   '
End Function
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

Function RASOCLI1$(NC%)
   Call DATCLIEN(NC%, VDEV$, 2)
   RASOCLI1$ = VDEV$
End Function
