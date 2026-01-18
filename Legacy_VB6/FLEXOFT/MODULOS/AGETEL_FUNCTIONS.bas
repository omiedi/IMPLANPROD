Attribute VB_Name = "AGETEL_FUNCTIONS"
Global CancelFlag


'--------------------------------------------------------
'   DIALER.FRM
'   Copyright (c) 1994 Crescent Software, Inc.
'   por Carl Franklin
'
'   Actualizado por Anton de Jong
'
'   Muestra cómo marcar números telefónicos con un módem.
'
'   Para que este programa funcione, el teléfono y el
'   módem deben estar conectados a la misma línea.
'--------------------------------------------------------
'Option Explicit

'DefInt A-Z

' Este indicador se establece cuando el usuario elige Cancelar.

Sub Dial(NULLAMA$, FORNAME$)
   '
   Dim DialString$, FromModem$, dummy
    ' AT es el comando compatible Hayes ATTENTION, y es necesario para enviar comandos al módem.
    ' DT significa "marcar con tonos". El comando Dial utiliza tonos, en lugar de pulsos (DP =
    ' marcar con pulsos). Number$ es el número de teléfono a marcar.
    ' El punto y coma indica al módem que vuelva al modo de comando después de marcar
    ' (importante). Es necesario un retorno de carro, vbCr, al enviar comandos al módem.
    '
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  Number$ = ""
   INTERN$ = ""
   TOMALIN$ = TRIM$(CONTROL("AGETELEF", "TOMALIN"))
   '
   PPXX% = InStr(NULLAMA$, "/")
   If PPXX% < 1 Then PPXX% = 1 + Len(NULLAMA$)
   INTERN$ = TRIM$(Mid$(NULLAMA$, PPXX% + 1))
   NULLAMA$ = TRIM$(Left$(NULLAMA$, PPXX% - 1))
   '
   Number$ = TOMALIN$
   For KKI% = 1 To Len(NULLAMA$)
     cara% = Asc(Mid$(NULLAMA$, KKI%, 1))
     If cara% = 44 Or cara% >= 48 Then
       If cara% <= 57 Then
         Number$ = Number$ + Chr$(cara%)
       End If
     End If
   Next KKI%
   If Number$ = "" Then Exit Sub
   '
10 Forms(KI% - 1).Label28 = Number$
   Forms(KI% - 1).Label29 = ""
   Forms(KI% - 1).Picture1.Visible = True
   Forms(KI% - 1).Picture1.Refresh
   CancelFlag = 0
   '
   DialString$ = "ATDT" + Number$ + ";" + vbCr
    '
    ' Configuración del puerto de comunicaciones.
    ' Suponiendo que hay conectado un mouse a COM1, se establece 2 en CommPort
   SETUPSTRI$ = CONTROL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   COPORT = Val(Left$(SETUPSTRI$, 6))
   On Error Resume Next
   Forms(KI% - 1).MSComm1.CommPort = COPORT
   On Error GoTo 0
   '
   SETTI$ = ""
   For KKI% = 1 To 4
     SETTI$ = SETTI$ + TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
     If KKI% < 4 Then SETTI$ = SETTI$ + ","
   Next KKI%
   On Error Resume Next
   Forms(KI% - 1).MSComm1.Settings = SETTI$
   If Err Then
      Call MENSERR(24, "Configuración de Puerto no Válida.\Corrija o Aplique Valores Predeterminados por Omisión.")
      Forms(KI% - 1).Picture1.Visible = False
      SETUPTEL.Show 1
      Exit Sub
   End If
    '
    ' Abre el puerto de comunicaciones.
   On Error Resume Next
   Forms(KI% - 1).MSComm1.PortOpen = True
   If Err Then
      Call MENSERR(24, "COM" + TRIM$(Str$(COPORT)) + " no está disponible.\Cambie la propiedad CommPort a otro puerto.")
      Forms(KI% - 1).Picture1.Visible = False
      Exit Sub
   End If
    '
    ' Vuelca el búfer de entrada.
   Forms(KI% - 1).MSComm1.InBufferCount = 0
    '
    ' Marca el número.
   Forms(KI% - 1).MSComm1.Output = DialString$
   TFIN = Timer + 10
   Forms(KI% - 1).Label30.Visible = False
    '
    ' Espera a que venga el "OK" del módem.
   Do
     dummy = DoEvents()
     CEVEN = Forms(KI% - 1).MSComm1.CommEvent
     Select Case CEVEN
        Case comEvReceive
          Forms(KI% - 1).Label29 = "Llamando ..."
          Forms(KI% - 1).Label29.Refresh
          Forms(KI% - 1).Label30.Visible = True
        Case comEvSend
          Forms(KI% - 1).Label29 = "Marcando ..."
        Case Else
     End Select
      '
      ' Si hay datos en el búfer, los lee.
     If Forms(KI% - 1).MSComm1.InBufferCount Then
       FromModem$ = FromModem$ + Forms(KI% - 1).MSComm1.Input
        ' Comprueba si hay "OK".
       If InStr(FromModem$, "OK") Then
             Forms(KI% - 1).Label29 = "Llamando ..."
             Forms(KI% - 1).Label29.Refresh
             Forms(KI% - 1).Label30.Visible = True
             TFIN = Timer + 3
             While Timer < TFIN
               dummy = DoEvents()
             Wend
             Exit Do
          End If
       End If
        
       ' ¿El usuario ha elegido Cancelar?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
       '
       If Timer > TFIN Then
         Forms(KI% - 1).Picture1.Visible = False
         MsgBox "Imposible Establecer Comunicación"
         Exit Do
       End If
       '
    Loop
    '
    ' Desconecta el módem.
    Forms(KI% - 1).MSComm1.Output = "ATH" + vbCr
    
    ' Cierra el puerto.
    Forms(KI% - 1).MSComm1.PortOpen = False
    Forms(KI% - 1).Picture1.Visible = False
    Call FRESHALL
    '
    If INTERN$ <> "" Then
      TFIN = Timer + 3
      While Timer < TFIN: Wend
      Number$ = INTERN$
      INTERN$ = ""
      vuelta% = 2
      GoTo 10
    End If
    '
End Sub

Sub ARMAMUESTRA(FORNAME$)
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  TT2$ = REPLACAR$(Forms(KI% - 1).Text2, Chr$(34), "")
   Forms(KI% - 1).Text2 = TT2$
   TT4$ = REPLACAR$(Forms(KI% - 1).Text4, Chr$(34), "")
   Forms(KI% - 1).Text4 = TT4$
   TTXX$ = TRIM$(Forms(KI% - 1).Text2)
   If TRIM$(Forms(KI% - 1).Text2) <> "" And TRIM$(Forms(KI% - 1).Text4) <> "" Then
        TTXX$ = TTXX$ + " - "
      Else
        TTXX$ = TTXX$ + " "
   End If
   Forms(KI% - 1).Label1 = TRIM$(TTXX$ + TRIM$(Forms(KI% - 1).Text4))
   '
End Sub

Sub GRAFICON(FORNAME$)
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  IFICHA% = Val(Forms(KI% - 1).Label26)
   REG& = REGBUS&("BCONTACT", 1, MKI$(IFICHA%))
   If REG& < 1 Then REG& = 1 + ULTREG&("BCONTACT")
   REGI% = REG&
   '
   If REG& <= ULTREG("BCONTACT") Then
        RELE$ = REGLEIDO$("BCONTACT", REG&)
      Else
        RELE$ = MKI$(REGI%) + Space$(1022)
   End If
   '
   Call REPLA(RELE$, Forms(KI% - 1).Label1, 3, 78)     ' nombre + empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text1, 81, 8)      ' apelativo Sr. / Ing. / Sres.
   Call REPLA(RELE$, Forms(KI% - 1).Text2, 89, 56)     ' nombre y apellido
   Call REPLA(RELE$, Forms(KI% - 1).Text4, 145, 64)    ' empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text5, 217, 64)    ' e-mail direccion 1
   Call REPLA(RELE$, Forms(KI% - 1).Text6, 281, 64)    ' e-mail direccion 2
   Call REPLA(RELE$, Forms(KI% - 1).Text7, 345, 64)    ' e-mail direccion 3
   '
   Call REPLA(RELE$, Forms(KI% - 1).Text16, 512 + 1, 48)  ' calle y nro part
   Call REPLA(RELE$, Forms(KI% - 1).Text15, 512 + 49, 16) ' c-post particular
   Call REPLA(RELE$, Forms(KI% - 1).Text13, 512 + 65, 48) ' localidad particular
   Call REPLA(RELE$, Forms(KI% - 1).Text8, 512 + 113, 32) ' provincia particular
   Call REPLA(RELE$, Forms(KI% - 1).Text12, 512 + 145, 32) ' pais particular
   Call REPLA(RELE$, Forms(KI% - 1).Text10, 512 + 177, 20) ' telef.partic.
   Call REPLA(RELE$, Forms(KI% - 1).Text25, 512 + 197, 20) ' telef2.partic
   Call REPLA(RELE$, Forms(KI% - 1).Text14, 512 + 217, 20) ' celular partic
   Call REPLA(RELE$, Forms(KI% - 1).Text11, 512 + 237, 20) ' fax partic.
   '
   Call REPLA(RELE$, Forms(KI% - 1).Text20, 768 + 1, 48)   ' calle y nro empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text21, 768 + 49, 16)  ' cpost empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text22, 768 + 65, 48)  ' localidad empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text19, 768 + 113, 32) ' provincia empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text23, 768 + 145, 32) ' pais empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text26, 768 + 177, 20) ' telef.empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text17, 768 + 197, 20) ' telef2 empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text24, 768 + 217, 20) ' fax empresa
   Call REPLA(RELE$, Forms(KI% - 1).Text18, 768 + 237, 20) ' cargo
   '
   Call GRAREG("BCONTACT", RELE$, REG&)
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   '
   Screen.MousePointer = 11
   NCON% = Val(Forms(KI% - 1).Label26)
   URN& = ULTREG&("NOTACONT")
   RNC$ = RECFILT$("NOTACONT", 1, MKI$(NCON%))
   KK& = (-3)
   '
   TTYZ$ = Forms(KI% - 1).Text9.Text
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&)
   While Len(TEXTO$) > 0
      TTYY$ = MKI$(NCON%) + TEXTO$ + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("NOTACONT", TTYY$, RENOTA&)
      TEXTO$ = Mid$(TEXTO$, 63)
   Wend
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("NOTACONT", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("NOTACONT")
   Screen.MousePointer = 1

End Sub

Sub CARFICHA(FORNAME$)
   '
   For KI% = 1 To Forms.Count
     If Forms(KI% - 1).Name = FORNAME$ Then GoTo 5
   Next KI%
   Exit Sub
   '
5  IFICHA% = Val(Forms(KI% - 1).Label26)
   REG& = REGBUS&("BCONTACT", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BCONTACT") Then
       Forms(KI% - 1).Text1 = ""
       Forms(KI% - 1).Text2 = ""
       Forms(KI% - 1).Text4 = ""
       Forms(KI% - 1).Text5 = ""
       Forms(KI% - 1).Text6 = ""
       Forms(KI% - 1).Text7 = ""
       Forms(KI% - 1).Text8 = ""
       Forms(KI% - 1).Text9 = ""
       Forms(KI% - 1).Text16 = ""
       Forms(KI% - 1).Text15 = ""
       Forms(KI% - 1).Text13 = ""
       Forms(KI% - 1).Text8 = ""
       Forms(KI% - 1).Text12 = ""
       Forms(KI% - 1).Text10 = ""
       Forms(KI% - 1).Text25 = ""
       Forms(KI% - 1).Text14 = ""
       Forms(KI% - 1).Text11 = ""
       Forms(KI% - 1).Text20 = ""
       Forms(KI% - 1).Text21 = ""
       Forms(KI% - 1).Text22 = ""
       Forms(KI% - 1).Text19 = ""
       Forms(KI% - 1).Text23 = ""
       Forms(KI% - 1).Text26 = ""
       Forms(KI% - 1).Text17 = ""
       Forms(KI% - 1).Text24 = ""
       Forms(KI% - 1).Text18 = ""
     Else
       RELE$ = REGLEIDO$("BCONTACT", REG&)
       Forms(KI% - 1).Text1 = TRIM$(Mid$(RELE$, 81, 8))
       Forms(KI% - 1).Text2 = TRIM$(Mid$(RELE$, 89, 56))
       Forms(KI% - 1).Text4 = TRIM$(Mid$(RELE$, 145, 64))
       Forms(KI% - 1).Text5 = TRIM$(Mid$(RELE$, 217, 64))
       Forms(KI% - 1).Text6 = TRIM$(Mid$(RELE$, 281, 64))
       Forms(KI% - 1).Text7 = TRIM$(Mid$(RELE$, 345, 64))
       '
       Forms(KI% - 1).Text16 = TRIM$(Mid$(RELE$, 512 + 1, 48))
       Forms(KI% - 1).Text15 = TRIM$(Mid$(RELE$, 512 + 49, 16))
       Forms(KI% - 1).Text13 = TRIM$(Mid$(RELE$, 512 + 65, 48))
       Forms(KI% - 1).Text8 = TRIM$(Mid$(RELE$, 512 + 113, 32))
       Forms(KI% - 1).Text12 = TRIM$(Mid$(RELE$, 512 + 145, 32))
       Forms(KI% - 1).Text10 = TRIM$(Mid$(RELE$, 512 + 177, 20))
       Forms(KI% - 1).Text25 = TRIM$(Mid$(RELE$, 512 + 197, 20))
       Forms(KI% - 1).Text14 = TRIM$(Mid$(RELE$, 512 + 217, 20))
       Forms(KI% - 1).Text11 = TRIM$(Mid$(RELE$, 512 + 237, 20))
       '
       Forms(KI% - 1).Text20 = TRIM$(Mid$(RELE$, 768 + 1, 48))
       Forms(KI% - 1).Text21 = TRIM$(Mid$(RELE$, 768 + 49, 16))
       Forms(KI% - 1).Text22 = TRIM$(Mid$(RELE$, 768 + 65, 48))
       Forms(KI% - 1).Text19 = TRIM$(Mid$(RELE$, 768 + 113, 32))
       Forms(KI% - 1).Text23 = TRIM$(Mid$(RELE$, 768 + 145, 32))
       Forms(KI% - 1).Text26 = TRIM$(Mid$(RELE$, 768 + 177, 20))
       Forms(KI% - 1).Text17 = TRIM$(Mid$(RELE$, 768 + 197, 20))
       Forms(KI% - 1).Text24 = TRIM$(Mid$(RELE$, 768 + 217, 20))
       Forms(KI% - 1).Text18 = TRIM$(Mid$(RELE$, 768 + 237, 20))
       '
       Forms(KI% - 1).Text9 = "": TTYZ$ = ""
       If IFICHA% > 0 Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTACONT", 1, MKI$(IFICHA%))
         URN& = ULTREG&("NOTACONT")
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URN& Then
               TTYZ$ = TTYZ$ + Mid$(REGLEIDO$("NOTACONT", RENOTA&), 3)
             End If
           End If
         Next U&
         Screen.MousePointer = 1
       End If
       '
       LOCADE& = Len(TTYZ$)
       For U& = LOCADE& To 1 Step -1
         If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
       Next U&
       U& = 0
30     Forms(KI% - 1).Text9 = Left$(TTYZ$, U&)
       Forms(KI% - 1).Text9.Refresh
       '
   End If
   '
End Sub

Sub SORCONTAC()
   '
   FICONTAC.List1.Clear
   Screen.MousePointer = 11
   '
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     TTXX$ = TRIM$(TTXX$ + TRIM$(TX4$))
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     TTXX$ = AJUSTI$(TTXX$, 64) + Str$(U&)
   Next U&
   '
   Call FILESORT("BCONTACT", "", 2, 2, "ASC")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub EXPORAGE()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:EXPORLIB.CSV" For Output As #N1%
   Print #N1%, "Nombre,Dirección electrónica"
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     '
     TTYY$ = TRIM$(TTXX$ + TRIM$(TX4$))
     TTZZ$ = REPLACAR$(TTYY$, Chr$(34), "")
     TTXX$ = REPLACAR$(TTZZ$, ",", " ")
     '
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     EMA$ = TRIM$(Mid$(RELE$, 217, 64))
     If EMA$ <> "" Then
       If Asc(Left$(EMA$, 1)) <> 39 Then
         AX$ = TTXX$ + "," + EMA$
         Print #N1%, AX$
         JJ& = JJ& + 1
       End If
     End If
   Next U&
   '
   Close #N1%
   Call CIERRARCH("*.*")
   MsgBox Str$(JJ&) + " Direcciones Exportadas"
   Screen.MousePointer = 1
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Sub IMPORAGE()
   '
   Screen.MousePointer = 11
   Dim RAEMP$()
   MAXEMP& = 4096
   ReDim Preserve RAEMP$(MAXEMP&)
   '
   For U& = 1 To ULTREG&("BCONTACT")
     X$ = REGLEIDO$("BCONTACT", U&)
     If U& > MAXEMP& Then
       MAXEMP& = MAXEMP& + 1024
       ReDim Preserve RAEMP$(MAXEMP&)
     End If
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
     CAEMP& = U&
   Next U&
   '
   LUCOMA$ = LUCOM$
   LUCOM$ = "C:\FLEXOFT\"
   '
   FIN& = ULTREG&("!BCONTACT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!BCONTACT", U&)
     RASOE$ = UCase$(TRIM$(Mid$(X$, 145, 64)))
     EMAI$ = LCase$(TRIM$(Mid$(X$, 217, 64)))
     For JJ& = 1 To CAEMP&
       If RAEMP$(JJ&) = RASOE$ Then
         Y$ = REGLEIDO$("BCONTACT", JJ&)
         If InStr(Y$, EMAI$) > 0 Then GoTo 19            ' YA ESTA
         If TRIM$(Mid$(Y$, 217, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 217, 64)
           ElseIf TRIM$(Mid$(Y$, 281, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 281, 64)
           ElseIf TRIM$(Mid$(Y$, 345, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 345, 64)
         End If
         Call GRAREG("BCONTACT", Y$, JJ&)
         GoTo 19
       End If
    Next JJ&
    '
    CAEMP& = CAEMP& + 1
    If CAEMP& > MAXEMP& Then
      MAXEMP& = MAXEMP& + 1024
      ReDim Preserve RAEMP$(MAXEMP&)
    End If
    RAEMP$(CAEMP&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
    KKI% = CAEMP&
    Call REPLA(X$, MKI$(KKI%), 1, 2)
    Call REGAPP("BCONTACT", X$)
19 Next U&
   Call CIERRARCH("*.*")
   LUCOM$ = LUCOMA$
   Screen.MousePointer = 1
   '
End Sub

Sub IMPOROUT()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:\OUTLOOK.CSV" For Input As #N1%
   Line Input #N1%, A$
   On Error GoTo 0
   '
   REG& = 0
   While Not EOF(N1%)
     REG& = REG& + 1
     REGI% = REG&
     RELE$ = MKI$(REGI%) + Space$(1022)
     '
     Line Input #N1%, TTXX$
     PPXX% = InStr(TTXX$, ",")
     If PPXX% > 0 Then
        EMPRE$ = Left$(TTXX$, PPXX% - 1)
        TTXX$ = Mid$(TTXX$, PPXX% + 1)
        PPXX% = InStr(TTXX$, ",")
        If PPXX% > 0 Then
          EMAI$ = Left$(TTXX$, PPXX% - 1)
          '
          Call REPLA(RELE$, EMPRE$, 3, 78)
          Call REPLA(RELE$, "Sr.", 81, 8)
          Call REPLA(RELE$, EMPRE$, 145, 64)
          Call REPLA(RELE$, EMAI$, 217, 64)
          Call GRAREG("BCONTACT", RELE$, REG&)
        End If
     End If
     '
   Wend
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub
