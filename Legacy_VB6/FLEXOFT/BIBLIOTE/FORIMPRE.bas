Attribute VB_Name = "FORIMPRE"
Public SETUPSTRI$
Public PRINTERPORT$
Dim FORMACT$
Public CODFOR$, DESFOR$
Dim CARE%, ATOT%
Public MARSUP%
Public MARIZW, MARSUW, EXTREDET, MUESTRAFORBLA%
Public DETDESMM, DETHASMM
Public INTERLI, ORIENTA%, OPCIMPRE%
Public CACOPI%
'
Public YAC
'
Dim CODVAR$(128), REGFOR%(128)
Dim CAFUEN As Integer
Dim FUENTE$(256)
Public ABOPRIN%
Public CANCELPRINT%
Public NOCONFIRM%
Public HOJASUEL%
Public MATRIXPUN%
'
Public NUMEDOC$, FECHDOC$, TIPODOC$, DOCUORIG$
Public CAPARDOC%, DOCPARAM$(256), CODPARAM$(256)
Public CAIMAGEN%, CODIMAGEN$(16), ARCHIMAGEN$(16)
Public CACOLTAB1%, CAITAB1%, TETABU1$(32, 1024), CODTABU1$(32)
Public CACOLTAB2%, CAITAB2%, TETABU2$(32, 1024), CODTABU2$(32)
Public DESTIDOC%   ' Destinatario del Documento
Public CACOPDOC%   ' CANTIDAD DE COPIAS
'
Public ALTUPAGINA, TOPEPAGINA
Public PAGINACTIVA%, CAPAGINAS%

Dim PORIE&, PSCAL&, PHEIG&, PWIDT&
Dim DRAMO&, DRAWI&, DRAST&, FILST&
Dim FCOLO&, FNAME$, FITAL&, FSIZE, FBOLD&
Dim CURRX1&, CURRY1&, CURRX2&, CURRY2&
Dim PTEXT$
'
Public REDUFORMU
'
Sub SETSPOOL(COMANDO$, ATRIBU$)
   '
   COPRIN$ = Space$(256)
   Call REPLA(COPRIN$, COMANDO$, 1, 16)
   Call REPLA(COPRIN$, ATRIBU$, 17, 240)
   FORMALTER.PSPOOL.AddItem COPRIN$
   '
End Sub
'
Sub SPOOLINE(H1, V1, H2, V2, CCOLO, BBXX$)
   ATRIBU$ = Space$(112)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(H1 + 0.5), 4, 12, 0, 1), 1, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(V1 + 0.5), 4, 12, 0, 1), 17, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(H2 + 0.5), 4, 12, 0, 1), 33, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(V2 + 0.5), 4, 12, 0, 1), 49, 16)
   Call REPLA(ATRIBU$, Str$(Int(CCOLO)), 65, 16)
   Call REPLA(ATRIBU$, BBXX$, 81, 16)
   Call SETSPOOL("LINE", ATRIBU$)
End Sub
'
Sub SPOOLIMAGE(IMAGEFILE$, PHX1, PVX1, ACHX, ALTX)
   '
   IMAGFI$ = TRIM$(IMAGEFILE$)
   If EXISTE%(IMAGFI$) < 1 Then Exit Sub
   '
   Dim IMAGFIDATE As Date
   IMAGFIDATE = FileDateTime(IMAGFI$)
   LARGIMAG& = FileLen(IMAGFI$)
   FECHIMAG# = CDbl(IMAGFIDATE)
   '
   NOMBIMAG$ = IMAGFI$
10 PPXX% = InStr(NOMBIMAG$, "\")
   If PPXX% > 0 Then
      NOMBIMAG$ = Mid$(NOMBIMAG$, PPXX% + 1)
      GoTo 10
   End If
   If TRIM$(NOMBIMAG$) = "" Then Exit Sub
   If Len(NOMBIMAG$) > 38 Then
      Call MENSERR(24, "Imposible Imprimir Imagen Contenida en Archivo\  \'" + NOMBIMAG$ + "'\  \Nombre de Archivo Demasiado Extenso (max = 38 car)")
      Exit Sub
   End If
   '
   ABUS$ = AJUSTI$(NOMBIMAG$, 38) + MKS$(LARGIMAG&) + MKD$(FECHIMAG#)
   BBUS% = BINVAL%(ABUS$)
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELE& = CVS(Mid$(RFF$, KKI%, 4))
     If RELE& > 0 Then
       XX$ = REGLEIDO$("PIMGLST", RELE&)
       If Mid$(XX$, 3, 50) = ABUS$ Then
         IMAGENUMBER& = CVS(Mid$(XX$, 53, 4))
         GoTo 50
       End If
     End If
   Next KKI%
   '
   Call BLOQARCH("PIMGLST")
   Call BLOQARCH("PIMGSPL")
   '
   LASTIMAGENUMBER& = 0
   URELIS& = ULTREG&("PIMGLST")
   If URELIS& > 0 Then
     XX$ = REGLEIDO$("PIMGLST", URELIS&)
     LASTIMAGENUMBER& = CVS(Mid$(XX$, 51, 4))
     If LASTIMAGENUMBER& < URELIS& Then LASTIMAGENUMBER& = URELIS&
   End If
   IMAGENUMBER& = 1 + LASTIMAGENUMBER&
   FIRSTRECORD& = 1 + ULTREG&("PIMGSPL")
   '
   IMGHEADER$ = ABUS$
   BBUS% = BINVAL%(ABUS$)
   '
   PPXX% = InStr(IMAGFI$, ".")
   EXTE$ = ".BMP"
   If PPXX% > 0 Then EXTE$ = Mid$(IMAGFI$, PPXX%)
   FileCopy IMAGFI$, "C:\FLEXOFT\IMAGEN" + EXTE$
   'NX% = FILEOPEN%(IMAGFI$, 0, 508)
   NX% = FILEOPEN%("C:\FLEXOFT\IMAGEN" + EXTE$, 0, 508)
   Dim BUFI As String * 508
   CARELE& = 1 + LOF(NX%) / 508
   JJ& = FIRSTRECORD& - 1
   For RELE& = 1 To CARELE&
     Get #NX%, RELE&, BUFI$
     ZZ$ = String$(512, 0)
     Call REPLA(ZZ$, MKS$(IMAGENUMBER&), 1, 4)
     Call REPLA(ZZ$, BUFI$, 5, 508)
     JJ& = JJ& + 1
     LASTRECORD& = JJ&
     Call GRAREG("PIMGSPL", ZZ$, JJ&)
   Next RELE&
   Call CIERRARCH("PIMGSPL")
   Close #NX%: NX% = 0
   '
   ZZ$ = String$(128, 0)
   Call REPLA(ZZ$, MKI$(BBUS%), 1, 2)
   Call REPLA(ZZ$, IMGHEADER$, 3, 50)
   Call REPLA(ZZ$, MKS$(IMAGENUMBER&), 53, 4)
   Call REPLA(ZZ$, MKS$(FIRSTRECORD&), 57, 4)
   Call REPLA(ZZ$, MKS$(LASTRECORD&), 61, 4)
   Call REGAPP("PIMGLST", ZZ$)
   Call CIERRARCH("PIMGLST")
   '
50 ATRIBU$ = Space$(112)
   Call REPLA(ATRIBU$, Str$(BBUS%), 1, 16)
   Call REPLA(ATRIBU$, Str$(IMAGENUMBER&), 17, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(PHX1 + 0.5), 4, 12, 0, 1), 33, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(PVX1 + 0.5), 4, 12, 0, 1), 49, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(ACHX + 0.5), 4, 12, 0, 1), 65, 16)
   Call REPLA(ATRIBU$, CSTRING$(MKS$(ALTX + 0.5), 4, 12, 0, 1), 81, 16)
   Call SETSPOOL("PAINTPICTURE", ATRIBU$)
   '
End Sub
'
Sub SAVESPOOL(DESCRIDOC$, OKX%)
   '
   Static CACOMM%, PRICOMM$(128)
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
   End If
   '
   If FORMALTER.PSPOOL.ListCount < 1 Then
      Exit Sub
   End If
   '
   Call BLOQARCH("PDOCLST")
   Call BLOQARCH("PDOCSPL")
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
   Call REPLA(PDLS$, Printer.DeviceName, 129, 128)
   Call REGAPP("PDOCLST", PDLS$)
   Call CIERRARCH("PDOCLST")
   '
10 For KKII% = 1 To FORMALTER.PSPOOL.ListCount
      TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           'If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKII%
   '
   KKJJ% = 0
   While Len(TPSP$) > 0
     TBUF$ = MKS$(DOCUNU&) + Left$(TPSP$ + String$(508, 0), 508)
     Call REGAPP("PDOCSPL", TBUF$)
     TPSP$ = Mid$(TPSP$, 509)
   Wend
   Call CIERRARCH("PDOCSPL")
   '
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
       Call COMUNI("Coloque Formulario y Pulse Aceptar")
     End If
   End If
   '
   Printer.EndDoc   ' INICIALIZA LA IMPRESORA
   '
   'If OPCIMPRE% > 0 Then
   '  FORMALTER.CommonDialog1.ShowPrinter
   'End If
   '
   For KKII% = 1 To FORMALTER.PSPOOL.ListCount
      TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
           GoTo 39
        End If
      Next KKJJ%
39 Next KKII%
   '
   FORMALTER.PSPOOL.Clear
   Call CLOSEFORMPRIN("")
   OKX% = 1
   '
End Sub
'
Sub SETPRINPORT(NEWPORT$)
   '
   Static IMPREDETER$, CALIPREDET
   '
   NP$ = TRIM$(NEWPORT$)
   Dim PRX As Object
   '
   If IMPREDETER$ = "" Then
      On Error Resume Next
      IMPREDETER$ = Printer.DeviceName
      On Error GoTo 0
   End If
   '
   If NP$ = "RESTORE" Then
      'On Error Resume Next
      'Printer.PrintQuality = CALIPREDET
      'On Error GoTo 0
      If TRIM$(IMPREDETER$) <> "" Then
        For Each PRX In Printers
          If PRX.DeviceName = IMPREDETER$ Then
            On Error Resume Next
            Set Printer = PRX
            On Error GoTo 0
            Exit Sub
          End If
        Next PRX
      End If
      Exit Sub
   End If
   '
   If NP$ <> "" Then
      For Each PRX In Printers
        If PRX.Port = NP$ Then
          On Error Resume Next
          Set Printer = PRX
          CALIPREDET = Printer.PrintQuality
          Printer.PrintQuality = (-1)         ' draft
          On Error GoTo 0
          Exit Sub
        End If
      Next PRX
      '
      NP1$ = REPLACAR$(NP$, "\", "/")
      Call COMUNI("Puerto de Impresión '" + NP1$ + "' no Encontrado.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub PRINTCOMMAND(COPRI%, ATRIP$)
   '
   Static LASTX, LASTY, LASTFONTNAME$
   '
   Select Case COPRI%
   Case 1
     PORIE& = Val(TRIM$(ATRIP$))
'MsgBox "ORIENTATION " + ATRIP$
'     On Error Resume Next
'     Printer.Orientation = Val(TRIM$(ATRIP$))
'     On Error GoTo 0
   Case 2
     Call VUELCAIMPRE(1)
'MsgBox "NEWPAGE "
'     If MATRIXPUN% = 1 Then
'        Call VUELCAIMPRE(1)
'        If HOJASUEL% = 1 Then
'          Call COMUNI("Coloque Formulario y Pulse Aceptar")
'        End If
'        Exit Sub
'     End If
'     If HOJASUEL% = 1 Then
'       Printer.EndDoc
'       Call COMUNI("Coloque Formulario y Pulse Aceptar")
'       Exit Sub
'     End If
'     Printer.NewPage
   Case 3
     Call VUELCAIMPRE(2)
'MsgBox "ENDDOC"
'     If MATRIXPUN% = 1 Then
'        Call VUELCAIMPRE(2)
'        Exit Sub
'     End If
'     Printer.EndDoc
   Case 4
     PHEIG& = Val(TRIM$(ATRIP$))
'MsgBox "HEIGHT " + ATRIP$
'     On Error Resume Next
'     Printer.Height = Val(TRIM$(ATRIP$))
'     On Error GoTo 0
   Case 5
     PWIDT& = Val(TRIM$(ATRIP$))
'MsgBox "WIDTH " + ATRIP$
'     On Error Resume Next
'     Printer.Width = Val(TRIM$(ATRIP$))
'     On Error GoTo 0
   Case 6
     PSCAL& = Val(TRIM$(ATRIP$))
'MsgBox "SCALEMODE " + ATRIP$
'     Printer.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     FCOLO& = Val(TRIM$(ATRIP$))
'MsgBox "FORECOLOR " + ATRIP$
'     Printer.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     DRAMO& = Val(TRIM$(ATRIP$))
'MsgBox "DRAWMODE " + ATRIP$
'     Printer.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
'MsgBox "DRAWWIDTH " + ATRIP$
     DRAWI& = Val(TRIM$(ATRIP$))
'     Printer.DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     DRAST& = Val(TRIM$(ATRIP$))
'MsgBox "DRAWSTYLE " + ATRIP$
'     Printer.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
'MsgBox "FILLSTYLE " + ATRIP$
     FILST& = Val(TRIM$(ATRIP$))
'     Printer.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
'MsgBox "LINE " + ATRIP$
     CURRX1& = Val(Mid$(ATRIP$, 1, 16))
     CURRY1& = Val(Mid$(ATRIP$, 17, 16))
     CURRX2& = Val(Mid$(ATRIP$, 33, 16))
     CURRY2& = Val(Mid$(ATRIP$, 49, 16))
     FCOLO& = Val(Mid$(ATRIP$, 65, 16))
     PTEXT$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     Call CARLIMPRE(2)
'     If BBLO$ = "B" Then
'        Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&, B
'       Else
'        Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&
'     End If
   Case 13
'MsgBox "FONTNAME " + ATRIP$
     FNAME$ = TRIM$(ATRIP$)
     'LASTFONTNAME$ = TRIM$(ATRIP$)
   Case 14
'MsgBox "FONTITALIC " + ATRIP$
     FITAL& = Val(TRIM$(ATRIP$))
     'Printer.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
'MsgBox "FONTSIZE " + ATRIP$
     FSIZE = Val(TRIM$(ATRIP$))
     'Printer.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     FBOLD& = Val(TRIM$(ATRIP$))
'MsgBox "FONTBOLD " + ATRIP$
'     Printer.FontBold = Val(TRIM$(ATRIP$))
   Case 17
'MsgBox "CURRENTX " + ATRIP$
     CURRX1& = Int(Val(TRIM$(ATRIP$)) + 0.5)
     'LASTX = Int(Val(TRIM$(ATRIP$)) + 0.5)
   Case 18
'MsgBox "CURRENTY " + ATRIP$
     CURRY1& = Int(Val(TRIM$(ATRIP$)) + 0.5)
     'LASTY = Int(Val(TRIM$(ATRIP$)) + 0.5)
   Case 19
'MsgBox "PRINT " + ATRIP$
'     ATRIP$ = RTrim$(ATRIP$)
'     If TRIM$(ATRIP$) = "" Or TRIM$(ATRIP$) = ";" Then Exit Sub
     '
'     PRUVUELTA% = 0
191 '  If Printer.CurrentX <> LASTX Or Printer.CurrentY <> LASTY Then
'       If PRUVUELTA% < 3 Then
'         PRUVUELTA% = PRUVUELTA% + 1
'         Printer.CurrentX = LASTX
'         Printer.CurrentY = LASTY
'         GoTo 191
'       End If
'       MsgBox "Imposible Imprimir '" + TRIM$(ATRIP$) + "' en Coordenadas " + TRIM$(Str$(LASTX)) + "," + TRIM$(Str$(LASTY))
'       Exit Sub
'     End If
     '
'     PRUVUELTA% = 0
192 ' If Printer.FontName <> LASTFONTNAME$ Then
'       If PRUVUELTA% < 3 Then
'         PRUVUELTA% = PRUVUELTA% + 1
'         Printer.FontName = LASTFONTNAME$
'         GoTo 192
'       End If
'       MsgBox "Fuente '" + TRIM$(LASTFONTNAME$) + "' no Disponible en el Presente Modelo de Impresora para Imprimir '" + TRIM$(ATRIP$) + "'."
'     End If
     '
     'If MATRIXPUN% = 1 Then
     PTEXT$ = RTrim$(ATRIP$)
     Call CARLIMPRE(3)    'Printer.CurrentX, Printer.CurrentY, ATRIP$)
     '   Exit Sub
     'End If
     '
195  'If Right$(ATRIP$, 1) = ";" Then
     '    ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
     '    Printer.Print RTrim$(ATRIP$);
     '   Else
     '    Printer.Print RTrim$(ATRIP$)
     'End If
     '
   Case 20
'MsgBox "UNDERLINE " + ATRIP$
'     Printer.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
'MsgBox "IMAGE " + ATRIP$
     'Call INCLUDEIMAGE(ATRIP$, "PRINT")
     PTEXT$ = ATRIP$
     Call CARLIMPRE(1)
   End Select
   GoTo 99
   '
99 End Sub
'

Sub CARLIMPRE(MODO%)
   TTXX$ = Space$(256)
   Call REPLA(TTXX$, CSTRING$(MKI$(MODO%), 1, 4, 0, 2), 1, 4)
   Select Case MODO%
     Case 1         ' IMAGEN
       Call REPLA(TTXX$, PTEXT$, 65, 192)
       GoTo 20
     Case 2         ' LINEA
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRY1&), 3, 8, 0, 2), 5, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRX1&), 3, 8, 0, 2), 13, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRY2&), 3, 8, 0, 2), 21, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRX2&), 3, 8, 0, 2), 29, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(FCOLO&), 3, 4, 0, 2), 37, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(DRAMO&), 3, 4, 0, 2), 41, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(DRAWI&), 3, 4, 0, 2), 45, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(DRAST&), 3, 4, 0, 2), 49, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(FILST&), 3, 4, 0, 2), 53, 4)
       Call REPLA(TTXX$, PTEXT$, 65, 192)
       GoTo 20
     Case 3         ' TEXTO
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRY1&), 3, 8, 0, 2), 5, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(CURRX1&), 3, 8, 0, 2), 13, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(FCOLO&), 3, 4, 0, 2), 37, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(FSIZE), 3, 8, 1, 2), 41, 8)
       Call REPLA(TTXX$, CSTRING$(MKS$(FBOLD&), 3, 4, 0, 2), 49, 4)
       Call REPLA(TTXX$, CSTRING$(MKS$(FITAL&), 3, 4, 0, 2), 53, 4)
       Call REPLA(TTXX$, FNAME$, 65, 64)
       Call REPLA(TTXX$, PTEXT$, 129, 128)
       GoTo 20
   End Select
   Exit Sub
   '
20 For KKL& = 1 To FORMALTER.List1.ListCount
     If FORMALTER.List1.List(KKL& - 1) > TTXX$ Then
       FORMALTER.List1.AddItem TTXX$, KKL& - 1
       Exit Sub
     End If
   Next KKL&
   FORMALTER.List1.AddItem TTXX$
   '
End Sub

Sub VUELCAIMPRE(MODOVUEL%)
   '
   ' MODOVUEL%=1 --> NEWPAGE
   ' MODOVUEL%=2 --> ENDDOC
   '
   Static FONODIS$(128), CAFONODI%
   '
   On Error Resume Next
   '
   Printer.EndDoc
   Printer.Orientation = PORIE&
   Printer.ScaleMode = PSCAL&
   '
   If PHEIG& < 1 Then
     PHEIG& = 56.7 * Val(ATRIFORM$("", "LARGOJMM"))
   End If
   '
   If PHEIG& > 8000 Then
       'Printer.Height = PHEIG&
     Else
       PHEIG& = Printer.Height
   End If
   If PWIDT& > 12000 Then
       'Printer.Width = PWIDT&
     Else
       PWIDT& = Printer.Width
   End If
   Printer.ForeColor = QBColor(0)
   Printer.FontTransparent = True
   '
   For KKL& = 1 To FORMALTER.List1.ListCount
     TTXX$ = FORMALTER.List1.List(KKL& - 1)
     MODO% = Val(Left$(TTXX$, 4))
     '
     Select Case MODO%
         Case 1
           ATRIP$ = Mid$(TTXX$, 65)
           Call INCLUDEIMAGE(ATRIP$, "PRINT")
           GoTo 19
         Case 2
           'Printer.ForeColor = Val(Mid$(TTXX$, 37, 4))
           'Printer.DrawMode = Val(Mid$(TTXX$, 41, 4))
           Printer.DrawWidth = Val(Mid$(TTXX$, 45, 4))
           'Printer.DrawStyle = Val(Mid$(TTXX$, 49, 4))
           Printer.FillStyle = Val(Mid$(TTXX$, 53, 4))
           AA11& = Val(Mid$(TTXX$, 13, 8)): If AA11& > PWIDT& - 10 Then AA11& = PWIDT& - 10
           BB11& = Val(Mid$(TTXX$, 5, 8)): If BB11& > PHEIG& - 10 Then BB11& = PHEIG& - 10
           AA22& = Val(Mid$(TTXX$, 29, 8)): If AA22& > PWIDT& - 10 Then AA22& = PWIDT& - 10
           BB22& = Val(Mid$(TTXX$, 21, 8)): If BB22& > PHEIG& - 10 Then BB22& = PHEIG& - 10
           BBLO$ = TRIM$(Mid$(TTXX$, 65))
           If BBLO$ = "B" Then
               Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&, B
            Else
               Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&
           End If
           GoTo 19
         Case 3
           PTXX$ = RTrim$(Mid$(TTXX$, 129))
           If TRIM$(PTXX$) <> ";" Then
             Printer.ForeColor = Val(Mid$(TTXX$, 37, 4))
             '
             FONAME$ = TRIM$(Mid$(TTXX$, 65, 64))
             Printer.FontName = FONAME$
             FOAPLI$ = Printer.FontName
             If FOAPLI$ <> FONAME$ Then
               For KKI% = 1 To CAFONODI%
                 If FONODIS$(KKI%) = FONAME$ Then GoTo 9
               Next KKI%
               CAFONODI% = CAFONODI% + 1
               FONODIS$(CAFONODI%) = FONAME$
               MERRO$ = "Fuente '" + FONAME$ + "' no Disponible"
               MERRO$ = MERRO$ + "\en el Presente Modelo de Impresora"
               MERRO$ = MERRO$ + "\para Leyenda '" + TRIM$(PTXX$) + "'."
               MERRO$ = MERRO$ + "\   \FLEXOFT ha Reemplazado esta Fuente"
               MERRO$ = MERRO$ + "\no Disponible por '" + FOAPLI$ + "'."
               Call COMUNI(MERRO$)
             End If
             '
9            Printer.FontSize = Val(Mid$(TTXX$, 41, 8))
             Printer.FontBold = Val(Mid$(TTXX$, 49, 4))
             Printer.FontItalic = Val(Mid$(TTXX$, 53, 4))
             AA11& = Val(Mid$(TTXX$, 13, 8)): If AA11& > PWIDT& - 100 Then AA11& = PWIDT& - 100
             BB11& = Val(Mid$(TTXX$, 5, 8)): If BB11& > PHEIG& - 100 Then BB11& = PHEIG& - 100
 If AA11& + Printer.TextWidth(RTrim$(PTXX$)) >= Printer.ScaleWidth Then
    MsgBox "Texto '" + TRIM$(PTXX$) + "' Fuera de Area de Impresion Horizontal"
    GoTo 19
 End If
 If BB11& + Printer.TextHeight(RTrim$(PTXX$)) >= Printer.ScaleHeight Then
    MsgBox "Texto '" + TRIM$(PTXX$) + "' Fuera de Area de Impresion Vertical"
    GoTo 19
 End If
             Printer.CurrentX = AA11&
             Printer.CurrentY = BB11&
             If Right$(PTXX$, 1) = ";" Then
                 PTXX$ = Left$(PTXX$, Len(PTXX$) - 1)
                 Printer.Print RTrim$(PTXX$);
               Else
                 Printer.Print RTrim$(PTXX$)
             End If
           End If
           GoTo 19
     End Select
19 Next KKL&
   '
   If MODOVUEL% = 1 Then
        Printer.NewPage
      Else
        Printer.EndDoc
   End If
   FORMALTER.List1.Clear
   On Error GoTo 0
   '
End Sub
'
Sub SHOWCOMMAND(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     'Printer.NewPage
     'If PAGINACTIVA% = 0 Then
     '   EDIFORMU.Picture1(0).Refresh
     '   EDIFORMU.VScroll1.Min = 0
     '   EDIFORMU.VScroll1.Value = 0
     '   EDIFORMU.VScroll1.Max = 1
     '   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
     '      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
     '   End If
     '   EDIFORMU.Picture1(0).Visible = True
     '   EDIFORMU.Picture1(0).Refresh
     'End If
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     'Printer.EndDoc
     'If PAGINACTIVA% = 0 Then
     '   EDIFORMU.Picture1(0).Refresh
     '   EDIFORMU.VScroll1.Min = 0
     '   EDIFORMU.VScroll1.Value = 0
     '   EDIFORMU.VScroll1.Max = 1
     '   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
     '      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
     '   End If
     '   EDIFORMU.Picture1(0).Visible = True
     '   EDIFORMU.Picture1(0).Refresh
     'End If
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'BB11 = Val(TRIM$(ATRIP$))
     'If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     'For KKI% = 0 To CAPAGINAS%
     '  If EDIFORMU.Picture1(KKI%).Height < ALTUPAGINA Then
     '    EDIFORMU.Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
   Case 5
     EDIFORMU.Picture1(PAGINACTIVA%).Width = Val(TRIM$(ATRIP$))
   Case 6
     EDIFORMU.Picture1(PAGINACTIVA%).ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     EDIFORMU.Picture1(PAGINACTIVA%).ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     EDIFORMU.Picture1(PAGINACTIVA%).DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     EDIFORMU.Picture1(PAGINACTIVA%).DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     EDIFORMU.Picture1(PAGINACTIVA%).DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     EDIFORMU.Picture1(PAGINACTIVA%).FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16))
     BB11 = Val(Mid$(ATRIP$, 17, 16))
     AA22 = Val(Mid$(ATRIP$, 33, 16))
     BB22 = Val(Mid$(ATRIP$, 49, 16))
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
     For KKI% = 0 To CAPAGINAS%
       If EDIFORMU.Picture1(KKI%).Height < ALTUPAGINA - 1 Then
         EDIFORMU.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     If AA11 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
        EDIFORMU.Picture1(PAGINACTIVA%).Width = AA11
     End If
     If AA22 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
        EDIFORMU.Picture1(PAGINACTIVA%).Width = AA22
     End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        EDIFORMU.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        EDIFORMU.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     EDIFORMU.Picture1(PAGINACTIVA%).FontName = TRIM$(ATRIP$)
   Case 14
     EDIFORMU.Picture1(PAGINACTIVA%).FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     EDIFORMU.Picture1(PAGINACTIVA%).FontSize = Val(TRIM$(ATRIP$))
   Case 16
     EDIFORMU.Picture1(PAGINACTIVA%).FontBold = Val(TRIM$(ATRIP$))
   Case 17
     If Val(TRIM$(ATRIP$)) > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
        EDIFORMU.Picture1(PAGINACTIVA%).Width = Val(TRIM$(ATRIP$))
     End If
     EDIFORMU.Picture1(PAGINACTIVA%).CurrentX = Val(TRIM$(ATRIP$))
   Case 18
     If Val(TRIM$(ATRIP$)) + 200 > ALTUPAGINA Then ALTUPAGINA = Val(TRIM$(ATRIP$)) + 1000
     For KKI% = 0 To CAPAGINAS%
       If EDIFORMU.Picture1(KKI%).Height < ALTUPAGINA - 1 Then
         EDIFORMU.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     EDIFORMU.Picture1(PAGINACTIVA%).CurrentY = Val(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         EDIFORMU.Picture1(PAGINACTIVA%).Print ATRIP$;
        Else
         EDIFORMU.Picture1(PAGINACTIVA%).Print ATRIP$
     End If
   Case 20
     EDIFORMU.Picture1(PAGINACTIVA%).FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     REDUFORMU = 1
     Call INCLUDEIMAGE(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'
Sub INCLUDEIMAGE(ATRIP$, MODOPRE$)
   '
   BBUS% = Val(Mid$(ATRIP$, 1, 16))
   IMAGENUMBER& = Val(Mid$(ATRIP$, 17, 16))
   PHX1 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
   PVX1 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
   ACHX = Val(Mid$(ATRIP$, 65, 16)) * REDUFORMU
   ALTX = Val(Mid$(ATRIP$, 81, 16)) * REDUFORMU
   '
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELI& = CVS(Mid$(RFF$, KKI%, 4))
     If RELI& > 0 Then
       ZZ$ = REGLEIDO$("PIMGLST", RELI&)
       If CVS(Mid$(ZZ$, 53, 4)) = IMAGENUMBER& Then GoTo 10
     End If
   Next KKI%
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
10 NOMBIMAG$ = TRIM$(Mid$(ZZ$, 3, 38))
   LARGIMAG& = CVS(Mid$(ZZ$, 41, 4))
   FECHIMAG# = CVD(Mid$(ZZ$, 45, 8))
   FIRSTRECORD& = CVS(Mid$(ZZ$, 57, 4))
   LASTRECORD& = CVS(Mid$(ZZ$, 61, 4))
   '
   If FIRSTRECORD& < 1 Or LASTRECORD& < FIRSTRECORD& Or LASTRECORD& > ULTREG&("PIMGSPL") Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
   If FIRSTRECORD& > 1 Then
      XX$ = REGLEIDO$("PIMGSPL", FIRSTRECORD& - 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   If LASTRECORD& < ULTREG&("PIMGSPL") Then
      XX$ = REGLEIDO$("PIMGSPL", LASTRECORD& + 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   IMAGSTRI$ = String$(508 * (1 + LASTRECORD& - FIRSTRECORD&), 0)
   PINSERT& = 1
   For U& = FIRSTRECORD& To LASTRECORD&
      XX$ = REGLEIDO$("PIMGSPL", U&)
      If CVS(Left$(XX$, 4)) <> IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
      Mid$(IMAGSTRI$, PINSERT&, 508) = Mid$(XX$, 5, 508)
      PINSERT& = PINSERT& + 508
   Next U&
   '
   If Len(IMAGSTRI$) < LARGIMAG& Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   IMAGSTRI$ = Left$(IMAGSTRI$, LARGIMAG&)
   '
   NX% = FreeFile
   Open LUCOM$ + NOMBIMAG$ For Binary As #NX%
   Put #NX%, , IMAGSTRI$
   Close #NX%
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(LUCOM$ + NOMBIMAG$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   If MODOPRE$ = "PRINT" Then
       Printer.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1, ACHX, ALTX
     ElseIf MODOPRE$ = "SHOW" Then
       '
       AA11 = PHX1
       BB11 = PVX1 + TOPEPAGINA * REDUFORMU
       AA22 = AA11 + ACHX
       BB22 = BB11 + ALTX
       If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
       If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
       For KKI% = 0 To CAPAGINAS%
         If EDIFORMU.Picture1(KKI%).Height < ALTUPAGINA - 1 Then
           EDIFORMU.Picture1(KKI%).Height = ALTUPAGINA
         End If
       Next KKI%
       If AA11 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
          EDIFORMU.Picture1(PAGINACTIVA%).Width = AA11
       End If
       If AA22 > EDIFORMU.Picture1(PAGINACTIVA%).Width Then
          EDIFORMU.Picture1(PAGINACTIVA%).Width = AA22
       End If
       '
       If REDUFORMU = 1 Then
          EDIFORMU.Picture1(PAGINACTIVA%).PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA, ACHX, ALTX
         Else
          EDIFORMU.Picture2(PAGINACTIVA%).PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA * REDUFORMU, ACHX, ALTX
       End If
   End If
   '
End Sub
'
Sub FORMMGR(Form$, RG%, CLII%, VLX$, CMD$)
   '
   ' FORM$  =  Nombre de formulario
   ' RG%    =  Renglon
   ' CLII%    =  Columna
   ' VLX$   =  Valor o texto a colocar
   ' CMD$   =  Comando a ejecutar
   '
   Static NX%, NY%, ATOT%
   Static FXXY As String * 128
   Static FYY As String * 256
   '
   If CMD$ = "CLOSE" Then
      '
      If NX% > 0 Then
         Close #NX%: NX% = 0
      End If
      If NY% > 0 Then
         Close #NY%: NY% = 0
      End If
      FORMACT$ = ""
      Exit Sub
      '
   End If
   '
   If Form$ <> "" Then
      If UCase$(TRIM$(Form$)) <> FORMACT$ Then
         FORMACT$ = UCase$(TRIM$(Form$))
         FORMUDRV$ = LUDAT$ + FORMACT$ + ".FRM"
         If NX% > 0 Then Close #NX%: NX% = 0
         NX% = FILEOPEN%(FORMUDRV$, 0, 128)
         '
         ATOT% = 256: CARE% = 128       ' ancho total y cantidad de renglones
         For K% = 1 To LOF(NX%) / 128
            Get #NX%, K%, FXXY$
            If Left$(FXXY$, 8) = "CAN-COL " Then
                 ATOT% = Val(TRIM$(Mid$(FXXY$, 89, 8)))
              ElseIf Left$(FXXY$, 8) = "CAN-REN " Then
                 CARE% = Val(TRIM$(Mid$(FXXY$, 89, 8)))
            End If
         Next K%
      End If
   End If
   '
   COMA$ = UCase$(TRIM$(CMD$))
   '
   If Left$(COMA$, 5) = "CLEAR" Then           ' genera formulario en blanco
      LUX$ = LUCOM$
      If NY% > 0 Then Close #NY%: NY% = 0
      NY% = FILEOPEN%(LUX$ + FORMACT$, 0, 256)
      '
      For I% = 1 To CARE%
         LSet FYY$ = Space$(256)
         Put #NY%, I%, FYY
      Next I%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "PRINT" Then
      Screen.MousePointer = 11
      UTI$ = LUBAS$
      If Mid$(UTI$, 2, 1) <> ":" Then
          UTI$ = Left$(CurDir, 2) + UTI$
      End If
      LUX$ = LUCOM$
      ATRIBU$ = "/" + LUX$ + FORMACT$
      ATRIBU$ = ATRIBU$ + "/CAREN=" + TRIM$(Str$(CARE%))
      ATRIBU$ = ATRIBU$ + "/SETUP=" + SETUPSTRI$
      ATRIBU$ = ATRIBU$ + "/MARSU=" + TRIM$(Str$(MARSUP%))
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
      Print #NZ%, "FINBAT"
      Close #NZ%: NZ% = 0
      '
      On Error Resume Next
      Kill "C:\FLEXOFT\FINBAT.$$$"
      On Error GoTo 0
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
      Print #NZ%, UTI$ + "PRINTFRM.EXE " + ATRIBU$
      Close #NZ%: NZ% = 0
      '
      Shell "C:\FLEXOFT\DOSCOMM.PIF"
      TINI = Timer
      MENSA$ = "Se ha producido un error no recuperable" + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "en la impresion de este documento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "Consulte a su Soporte de Sistemas FLEXOFT"
      '
149   If Timer > TINI + 30 Then
          Screen.MousePointer = 1
          MsgBox MENSA$, 4096 + 48, "Control de Impresion"
          On Error Resume Next
          Kill "C:\FLEXOFT\COMBAT.$$$"
          Kill "C:\FLEXOFT\FINBAT.$$$"
          On Error GoTo 0
          ABOPRIN% = 1
          Exit Sub
      End If
      '
      If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 149
      Screen.MousePointer = 1
      ABOPRIN% = 0
      Exit Sub
      '
   ElseIf Left$(COMA$, 4) = "FSET" Then
      If RG% > 0 Then
         If RG% <= CARE% Then
            If CLII% > 0 Then
               If CLII% <= ATOT% Then
                  Get #NY%, RG%, FYY$
                  TXTT$ = FYY$
                  Mid$(TXTT$, CLII%, Len(VLX$)) = VLX$
                  LSet FYY$ = TXTT$
                  Put #NY%, RG%, FYY$
               End If
            End If
         End If
      End If
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "LEYEN" Then
      For K% = 1 To LOF(NX%) / 128
        Get #NX%, K%, FXXY$
        If Left$(FXXY$, 4) = "LEYE" Then
          VLX$ = TRIM$(Mid$(FXXY$, 9, 80))
          If VLX$ <> "" Then
            For T% = 97 To 128 Step 8
               RG% = Asc(Mid$(FXXY$, T%, 1))
               If RG% > 0 Then
                  If RG% <= CARE% Then
                     CLII% = Asc(Mid$(FXXY$, T% + 1, 1))
                     If CLII% > 0 Then
                        If CLII% <= ATOT% Then
                           FORMATO$ = TRIM$(Mid$(FXXY$, T% + 2, 6))
                           If FORMATO$ <> "" Then
                              TX$ = TEXTFORM$(VLX$, FORMATO$)
                              Get #NY%, RG%, FYY$
                              TXTT$ = FYY$
                              Mid$(TXTT$, CLII%, Len(TX$)) = TX$
                              FYY$ = TXTT$
                              Put #NY%, RG%, FYY$
                           End If
                        End If
                     End If
                  End If
               End If
            Next T%
          End If
        End If
      Next K%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "@" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      For K% = 1 To LOF(NX%) / 128
         Get #NX%, K%, FXXY$
         If Left$(FXXY$, 8) = CD$ Then
            RG0% = RG%
            UREGUT% = 0
            For T% = 97 To 128 Step 8
               RG% = RG0%
               RGX% = Asc(Mid$(FXXY$, T%, 1))
               If RG% + RGX% > 0 Then
                  If RG% + RGX% <= CARE% Then
                     CLII% = Asc(Mid$(FXXY$, T% + 1, 1))
                     If CLII% > 0 Then
                        If CLII% <= ATOT% Then
                           FORMATO$ = TRIM$(Mid$(FXXY$, T% + 2, 6))
                           If FORMATO$ <> "" Then
                              ILETRA% = 0
                              '
                              If CD$ = "IMP-LET1" Then
                                 ILETRA% = 1
                                 VLX1$ = VLX$
                                 LOTEX% = Val(Mid$(FORMATO$, 2, 4))
                                 Call FRATEX(VLX$, NRX1$, NRX2$, NRX3$, LOTEX%)
                                 VLX$ = NRX1$ + String$(LOTEX%, "*")
                                 RL% = 1
                              End If
                              '
                              If CD$ = "ORD-COPI" Then
                                 ORDECOP% = ORDECOP% + 1
                                 VLX$ = ""
                                 If ORDECOP% <= 10 Then
                                    VLX$ = ORCOP$(ORDECOP%)
                                 End If
                              End If
                              '
1000                          TX$ = TEXTFORM$(VLX$, FORMATO$)
                              RG% = RG% + RGX%
                              Get #NY%, RG%, FYY$
                              TXTT$ = FYY$
                              Mid$(TXTT$, CLII%, Len(TX$)) = TX$
                              LSet FYY$ = TXTT$
                              Put #NY%, RG%, FYY$
                              If RG% > UREGUT% Then UREGUT% = RG%
                              '
                              If ILETRA% = 1 Then
                                 If RL% = 1 Then
                                     RL% = 2
                                     VLX$ = NRX2$ + String$(LOTEX%, "*")
                                     RGX% = 1
                                     GoTo 1000
                                   ElseIf RL% = 2 Then
                                     RL% = 3
                                     VLX$ = NRX3$ + String$(LOTEX%, "*")
                                     RGX% = 1
                                     GoTo 1000
                                 End If
                                 VLX$ = VLX1$
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            Next T%
            If UREGUT% > RG% Then RG% = UREGUT%
            Exit Sub
         End If
      Next K%
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "?" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      If NX% > 0 Then
         For K% = 1 To LOF(NX%) / 128
            Get #NX%, K%, FXXY$
            If Left$(FXXY$, 8) = CD$ Then
               VLX$ = TRIM$(Mid$(FXXY$, 89, 8))
               Exit Sub
            End If
         Next K%
      End If
      VLX$ = ""
      Exit Sub
      '
   Else
      '
      Call COMUNI("Comando '" + COMA$ + "' no Procesable\en Rutina 'FORMMGR'\para Formulario '" + FORMACT$ + "'.")
      '
   End If
   '
999 End Sub
   '
'*****************************************************************************
   '
Sub SETFORM(Form$, VLX$, CTRX$)
   '
   CMD$ = "@" + TRIM$(CTRX$)
   RG% = 0: CLII% = 0
   Call FORMMGR(Form$, RG%, CLII%, VLX$, CMD$)
   ' sobre RG% y CLII% devuelve renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CARLEYEN(Form$)
   '
   CMD$ = "LEYEN"
   RG% = 0: CLII% = 0
   Call FORMMGR(Form$, RG%, CLII%, VLX$, CMD$)
   '
End Sub
'
'*****************************************************************************
'
Sub APPFORM(VLX$, CTRX$)
   '
   TEXT$ = VLX$
   CODVARI$ = CTRX$
   X0 = 0
   Y0 = YAC
   If TRIM$(TEXT$) <> "" Then
       Call PRINTFORWIN(TEXT$, CODVARI$, X0, Y0)
       YACNUE = Printer.CurrentY / 56.7 - MARSUW
       If YACNUE > YAC Then YAC = YACNUE
   End If
   '
End Sub
'
'*****************************************************************************
'
Sub CLEARFORMPRIN(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "CLEAR")
   '
End Sub
'
'*****************************************************************************
'
Sub CLOSEFORMPRIN(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "CLOSE")
   '
End Sub
'
'*****************************************************************************
'
Sub PRINTFORM(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "PRINT")
   '
End Sub
'
'
'*****************************************************************************
'
Sub CARSETUP(AAA$)
   '
   If AAA$ = "NEW" Then
        SETUPSTRI$ = ""
      Else
        For UU% = 1 To Len(AAA$)
           UV$ = TRIM$(Str$(Asc(Mid$(AAA$, UU%, 1))))
           While Len(UV$) < 3
              UV$ = "0" + UV$
           Wend
           SETUPSTRI$ = SETUPSTRI$ + UV$
        Next UU%
   End If
   '
End Sub
'
'*****************************************************************************
'
Function ATRIFORM$(Form$, CTRX$)
   '
   CMD$ = "?" + TRIM$(CTRX$)
   Call FORMMGR(Form$, RG%, CLII%, VLX$, CMD$)
   ATRIFORM$ = VLX$
   '
End Function
'
'*****************************************************************************
'
Sub PRINTDOC(DOCTYPE$)
   '
   Dim ORCOP$(6)
   ORCOP$(1) = "ORIGINAL"
   ORCOP$(2) = "DUPLICADO"
   ORCOP$(3) = "TRIPLICADO"
   ORCOP$(4) = "CUADRUPLICADO"
   ORCOP$(5) = "QUINTUPLICADO"
   ORCOP$(6) = "SEXTUPLICADO"
   '
   If DOCTYPE$ = "" Then
      Call MENSERR(24, "Error de Configuración:\  \Falta Definir Tipo de Documento 'DOCTYPE'.")
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   If Left$(DOCTYPE$, 1) = "@" Then
        FORIPRE$ = TRIM$(Mid$(DOCTYPE$, 2))
      Else
        FORIPRE$ = TRIM$(CONTROL$("", DOCTYPE$))
   End If
   '
   If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario de Impresión\Correspondiente a Documento " + TRIM$(DOCTYPE$) + ".")
        Screen.MousePointer = 1
        Exit Sub
      Else
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LUDAT$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\Correspondiente a " + TRIM$(DOCTYPE$) + ":\  \'" + RMCC$ + "'.")
           Screen.MousePointer = 1
           Exit Sub
        End If
   End If
   '
   If App.EXEName = "FACTUMAI" Then
      Screen.MousePointer = 1
      COLOFORM.Label1.Caption = NUMEDOC$
      COLOFORM.Label2.Caption = TIPODOC$
      COLOFORM.Show 1
   End If
   '
   Screen.MousePointer = 11   ' Generando Documento
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT(PRINTERPORT$)
   End If
   '
   CODFOR$ = TRIM$(UCase$(FORIPRE$))
   MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
   MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   LARHOW = Val(ATRIFORM$(CODFOR$, "LARGOJMM"))
   If LARHOW < 80 Then LARHOW = 25.4 * 12
   '
   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   MULTIREN% = 1
   PPXX% = InStr(INTERX$, "/")
   If PPXX% > 0 Then
      MULTIREN% = Val(Mid$(INTERX$, PPXX% + 1))
      INTERX$ = Left$(INTERX$, PPXX% - 1)
   End If
   If MULTIREN% < 1 Then MULTIREN% = 1
   INTERLI = Val(INTERX$)
   If INTERLI < 1 Then
      Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
      INTERLI = 4
   End If
   '
   DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
   DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   '
   'CANHOJ& = 1
   'If CAITAB1% > 0 Then
   '   On Error Resume Next
   '   CANHOJ& = 1 + Int(MULTIREN% * CAITAB1% / ((DETHASMM - DETDESMM) / INTERLI))
   '   On Error GoTo 0
   'End If
   'If CANHOJ& < 1 Then CANHOJ& = 1
   '
   CACOPX$ = TRIM$(ATRIFORM$(CODFOR$, "CAN-COP"))
   If Left$(CACOPX$, 1) = "#" Then
5      CACOPI% = Val(InputBox$("Cantidad de Copias", "Ingrese Cantidad de Copias", Mid$(CACOPX$, 2)))
       If CACOPI% < 1 Or CACOPI% > 6 Then GoTo 5
     Else
       CACOPI% = Val(CACOPX$)
       If CACOPDOC% > 0 Then CACOPI% = CACOPDOC%
       If CACOPI% < 1 Or CACOPI% > 6 Then
         Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
         CACOPI% = 1
       End If
   End If
   '
   ORIENTA% = 1
   PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
   If PORI$ = "PAISAJE" Or PORI$ = "2" Then
      ORIENTA% = 2
   End If
   '
   PORIAN = Printer.Orientation  ' retiene orientacion
   'Printer.Orientation = 1
   Call SETSPOOL("ORIENTATION", "1")
   On Error Resume Next
     Printer.Orientation = 1
     LARHOAN = Printer.Height
     ANCHOAN = Printer.Width
   On Error GoTo 0
   '
10 'Printer.Orientation = 1
   Call SETSPOOL("ORIENTATION", "1")
   If ORIENTA% = 2 Then
      'Printer.Orientation = 2
      Call SETSPOOL("ORIENTATION", "2")
   End If
   '
   'Printer.Height = LARHOW * 56.7
   Call SETSPOOL("HEIGHT", Str$(LARHOW * 56.7))
   'Printer.Width = ANCHOAN
   Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   '
   For COP% = 1 To CACOPI%
      '
      INIDET& = 1
      NUHOJ% = 0
      CANHOJ& = 1
      '
20    Call CARFORWIN         ' Carga Formulario en Blanco
      Call PRINTFORWIN(TIPODOC$, "TIPO-COM", 0, 0)
      Call PRINTFORWIN(NUMEDOC$, "NUME-COM", 0, 0)
      Call PRINTFORWIN(DOCUORIG$, "DOCUORIG", 0, 0)
      Call PRINTFORWIN(FECHDOC$, "FECH-EMI", 0, 0)
      Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
      Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
      '
      Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
      '
      If Abs(DESTIDOC%) > 0 Then
         Call PRINTFORWIN(RASOCLI$(DESTIDOC%), "RASO-CLI", 0, 0)
         Call PRINTFORWIN(DOMICLI$(DESTIDOC%), "DOMI-CLI", 0, 0)
         Call PRINTFORWIN(CPOSCLI$(DESTIDOC%), "CPOS-CLI", 0, 0)
         Call PRINTFORWIN(LOCACLI$(DESTIDOC%), "LOCA-CLI", 0, 0)
         Call PRINTFORWIN(DEPROVI$(PROVCLI$(DESTIDOC%)), "PCIA-CLI", 0, 0)
         Call PRINTFORWIN(TEL1CLI$(DESTIDOC%), "TELE-CLI", 0, 0)
         Call PRINTFORWIN(FAXCLI$(DESTIDOC%), "FAX-CLI", 0, 0)
         Call PRINTFORWIN(CUITCLI$(DESTIDOC%), "CUIT-CLI", 0, 0)
         Call PRINTFORWIN(POSIVAL$(DESTIDOC%), "PIVA-CLI", 0, 0)
         Call PRINTFORWIN(NPROCLI$(DESTIDOC%), "PROV-CLI", 0, 0)
         Call PRINTFORWIN(Str$(Abs(DESTIDOC%)), "CUEN-CLI", 0, 0)
      End If
      '
      For KKI% = 1 To CAPARDOC%          '
         Call PRINTFORWIN(DOCPARAM$(KKI%), CODPARAM$(KKI%), 0, 0)
      Next KKI%
      '
      For KKI% = 1 To CAIMAGEN%          '
         Call PRINTIMAGE(CODIMAGEN$(KKI%), ARCHIMAGEN$(KKI%))
      Next KKI%
      '
      YAC = 0
      For YX& = 1 To CAITAB2%
         For KKI% = 1 To CACOLTAB2%
            TTTXXX$ = TETABU2$(KKI%, YX&)
            Call PRINTFORWIN(TTTXXX$, CODTABU2$(KKI%), 0, YAC)
         Next KKI%
         YAC = YAC + INTERLI
      Next YX&
      '
      YAC = 0: INTEROBS = INTERLI
      If INTERLI > 6 Then INTEROBS = INTERLI / 2
      For U& = 1 To ULTREG&("!OBSERVOF")
         TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
         Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
         YAC = YAC + INTEROBS
      Next U&
      Call CIERRARCH("!OBSERVOF")
      '
      YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
      '
      For YX& = INIDET& To CAITAB1%
         '
         If YAC > DETHASMM Then
            '
            NUHOJ% = NUHOJ% + 1
            If NUHOJ% = 1 Then
               On Error Resume Next
               CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.05)
               On Error GoTo 0
            End If
            Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
            Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
            '
            INIDET& = YX&
            'Printer.NewPage
            Call SETSPOOL("NEWPAGE", "")
            GoTo 20
         End If
         '
         YACMAX = 0
         For KKI% = 1 To CACOLTAB1%
            TTTXXX$ = TETABU1$(KKI%, YX&)
            If Left$(TTTXXX$, 1) = "@" Then
               YAC = YAC - INTERLI
               TTTXXX$ = Mid$(TTTXXX$, 2)
            End If
            Call PRINTFORWIN(TTTXXX$, CODTABU1$(KKI%), 0, YAC)
            If Printer.CurrentY / 56.7 - MARSUW + INTERLI > YACMAX Then
               YACMAX = Printer.CurrentY / 56.7 - MARSUW + INTERLI
            End If
         Next KKI%
         '
         If YACMAX < YAC Then YACMAX = YAC + INTERLI
         YAC = YACMAX
         '
      Next YX&
      '
      NUHOJ% = NUHOJ% + 1
      If NUHOJ% = 1 Then
         On Error Resume Next
         CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.05)
         On Error GoTo 0
      End If
      Call PRINTFORWIN(TRIM$(Str$(NUHOJ%)), "HOJA-NUM", 0, 0)
      Call PRINTFORWIN(TRIM$(Str$(CANHOJ&)), "HOJA-TOT", 0, 0)
      '
      If COP% < CACOPI% Then
         'Printer.NewPage
         Call SETSPOOL("NEWPAGE", "")
      End If
      '
   Next COP%
   '
   'Printer.EndDoc
99 Call SETSPOOL("ENDDOC", "")
   '
   DESCRIDOC$ = TRIM$(TIPODOC$) + " Nro." + NUMEDOC$
   If Len(TIPODOC$) > 8 Then
     DESCRIDOC$ = TRIM$(TIPODOC$) + " " + NUMEDOC$
   End If
   If Abs(DESTIDOC%) > 0 Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + TRIM$(RASOCLI$(DESTIDOC%))
   End If
   If TRIM$(FECHDOC$) <> "" Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + FECHDOC$
   End If
   Call SAVESPOOL(DESCRIDOC$, OKX%)
   '
   Call CIERRARCH("PDOCLST")
   Call CIERRARCH("PDOCSPL")
   Call CIERRARCH("PIMGLST")
   Call CIERRARCH("PIMGSPL")
   '
   Screen.MousePointer = 1
   If NOCONFIRM% < 1 Then
     If CANCELPRINT% < 1 Then
       If TRIM$(NUMEDOC$) <> "" Then
         TTXYZ$ = "Se ha Emitido por Impresora"
         If MATRIXPUN% = 1 Then TTXYZ$ = TTXYZ$ + " (MP)"
         TTXYZ$ = TTXYZ$ + "\" + TIPODOC$ + " Numero '" + NUMEDOC$ + "'"
         If COMALTER%(TTXYZ$ + "\\Repetir Impresion\Continuar") = 1 Then GoTo 10
       End If
     End If
   End If
   '
   'Printer.Orientation = 1
   Call SETSPOOL("ORIENTATION", "1")
   'Printer.Height = LARHOAN
   Call SETSPOOL("HEIGHT", Str$(LARHOAN))
   'Printer.Width = ANCHOAN
   Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   'Printer.Orientation = PORIAN
   Call SETSPOOL("ORIENTATION", Str$(PORIAN))
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT("RESTORE")
   End If
   '
   Call SETULTREG("!OBSERVOF", 0)
   Call CIERRARCH("!OBSERVOF")
   '
End Sub
'
Sub PRINTIMAGE(IMAGECOD$, IMAGEFILE$)
   '
   Static INIBUS%
   '
   If TRIM$(CODFOR$) = "" Then Exit Sub
   If EXISTE%(IMAGEFILE$) < 1 Then Exit Sub
   '
   CODVARI$ = TRIM$(UCase$(IMAGECOD$))
   If CODVARI$ = "" Then Exit Sub
   '
   Dim RFP As String * 128
   LU$ = LUDAT$
   NXP% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
   If INIBUS% < 1 Then INIBUS% = 1
   '
10 For I% = INIBUS% To LOF(NXP%) / 128
      Get #NXP%, I%, RFP$
      If Left$(RFP$, 8) = "IMAGEWIN" Then
         If TRIM$(Mid$(RFP$, 9, 16)) = CODVARI$ Then
            INIBUS% = I% + 1
            GoTo 20
         End If
      End If
   Next I%
   If INIBUS% > 1 Then
     INIBUS% = 1
     GoTo 10
   End If
   Exit Sub
   '
20 PH1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
   PV1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
   '
   PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
   PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   HCUADRO = PH2 - PH1
   VCUADRO = PV2 - PV1
   '
   REDUH = HIMAGEN / HCUADRO
   REDUV = VIMAGEN / VCUADRO
   REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
   '
   ALTUCUA = VCUADRO * REDUV / REDUC
   ANCHCUA = HCUADRO * REDUH / REDUC
   BORIZQ = (HCUADRO - ANCHCUA) / 2
   BORSUP = (VCUADRO - ALTUCUA) / 2
   '
   'Printer.PaintPicture FORMALTER.Image1.Picture, PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
   Call SPOOLIMAGE(IMAGEFILE$, PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA)
   '
End Sub
'
'*****************************************************************************
'
Sub FRATEX(TX$, NRX1$, NRX2$, NRX3$, LOTEX%)
   '
   Dim PAL$(24), NUMERO$(3)
   DE$ = TX$
   KK% = 1
   '
11550 For KM% = 1 To Len(DE$)
11560 PAL$(KK%) = PAL$(KK%) + Mid$(DE$, KM%, 1)
11570 If Mid$(DE$, KM%, 1) = " " Then KK% = KK% + 1
      If KK% > 24 Then GoTo 11590
11580 Next KM%
      '
11590 KN% = 1
11600 LON% = 0
11610 For KL% = 1 To KK%
11620    If LON% + Len(PAL$(KL%)) > LOTEX% Then
            If KN% < 3 Then
               KN% = KN% + 1
               LON% = 0
            End If
         End If
11630    NUMERO$(KN%) = NUMERO$(KN%) + PAL$(KL%)
         LON% = LON% + Len(PAL$(KL%))
11655 Next KL%
      '
11660 NRX1$ = NUMERO$(1)
      NRX2$ = NUMERO$(2)
      NRX3$ = NUMERO$(3)
      '
End Sub
'

Sub GRATITUFORM()
    '
    Screen.MousePointer = 11
    If TRIM$(PRINFORM.Text1.TEXT) = TRIM$(CODFOR$) Then
        If TRIM$(CODFOR$) <> "" Then
            ARXYZ$ = TRIM$(CODFOR$)
            TITUFOR$ = PRINFORM.Text2.TEXT
            NX% = FILEOPEN%(LU$ + ARXYZ$ + ".FRM", 0, 128)
            Dim RF As String * 128
            '
            For I% = 1 To LOF(NX%) / 128
                Get #NX%, I%, RF$
                If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                    RFY$ = RF$
                    Call REPLA(RFY$, TITUFOR$, 9, 48)
                    LSet RF$ = RFY$
                    Put #NX%, I%, RF$
                    GoTo 10
                End If
            Next I%
            '
            I% = 1 + LOF(NX%) / 128
            RFY$ = Space$(56) + String$(72, 0)
            Call REPLA(RFY$, "TITUFORM", 1, 8)
            Call REPLA(RFY$, TITUFOR$, 9, 48)
            LSet RF$ = RFY$
            Put #NX%, I%, RF$
            '
10          Close #NX%: NX% = 0
            '
            For V& = 1 To ULTREG&("INDIFRM")
                   INF$ = REGLEIDO$("INDIFRM", V&)
                   If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
                       If TITUFOR$ <> "" Then
                           Call REPLA(INF$, TITUFOR$, 9, 48)
                           Call GRAREG("INDIFRM", INF$, V&)
                       End If
                       GoTo 19
                   End If
            Next V&
            '
            INF$ = REGBLAN$("INDIFRM")
            Call REPLA(INF$, ARXYZ$, 1, 8)
            Call REPLA(INF$, TITUFOR$, 9, 48)
            Call REGAPP("INDIFRM", INF$)
            '
        End If
    End If
    '
19  Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'

Sub ALTANUEFORM()
        '
1000    CCC$ = OBJPLA$("CODEFORMU", CCC$)
        CODFOR$ = AJUSTI$(Left$(CCC$, 8), 8)
        If TRIM$(CODFOR$) = "" Then Exit Sub
        For UU% = 1 To Len(TRIM$(CODFOR$))
           VV% = Asc(Mid$(CODFOR$, UU%))
           If VV% < 48 Or VV% > 90 Or (VV% > 57 And VV% < 65) Then
              If VV% <> 45 Then
                 Call COMUNI("Código no Válido")
                 GoTo 1000
              End If
           End If
        Next UU%
        '
        DESFOR$ = Mid$(CCC$, 9, 48)
        If TRIM$(DESFOR$) = "" Then
            Beep
            GoTo 1000
        End If
        '
        If EXISTE%(LU$ + TRIM$(CODFOR$) + ".FRM") > 0 Then
              X% = COMALTER%("Ya Existe Formulario\'" + TRIM$(CODFOR$) + ".FRM'\\Otro Código\Sobre-Escribir")
              If X% <> 2 Then GoTo 1000
              Kill LU$ + TRIM$(CODFOR$) + ".FRM"
            Else
              Call REGAPP("INDIFRM", CCC$)
              Call CIERRARCH("INDIFRM")
        End If
        '
        If ULTREG&("INDIFRM") > 1 Then
           If ALTERNA%("Copia Formato Existente") = 1 Then
1060          Call SELECHO("INDIFRM")
              CCC$ = VALACT1$("INDIFRM")
              If Len(CCC$) <= 4 Then Exit Sub
              '
              CODFOR1$ = AJUSTI$(Left$(CCC$, 8), 8)
              If CODFOR1$ = CODFOR$ Then
                  Beep
                  GoTo 1060
              End If
              '
              FileCopy LU$ + TRIM$(CODFOR1$) + ".FRM ", LU$ + TRIM$(CODFOR$) + ".FRM"
           End If
        End If
        '
End Sub
'
Sub DEFLAYOUTW()

    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMU")
    Call BLANARCH("!CARFORMU")
    '
    'Call REGAPP("!CARFORMU", "TI-LETRATipo de letra (PICA - ELITE - COMPRIM)")
    'Call REGAPP("!CARFORMU", "LIN-PULGLineas por Pulgada (6 - 8)")
    'Call REGAPP("!CARFORMU", "CAN-COL Cantidad de Columnas sin Margen (h.256)")
    'Call REGAPP("!CARFORMU", "CAN-REN Cantidad de Renglones (inc.Margen Sup)")
    Call REGAPP("!CARFORMU", "MARSUPMMMargen Superior (milimetros)")
    Call REGAPP("!CARFORMU", "MARIZQMMMargen Izquierdo (milimetros)")
    Call REGAPP("!CARFORMU", "MARGENINMargen Intermedio (espacios)")
    Call REGAPP("!CARFORMU", "LARGOJMMLargo de Hoja (milimetros)")
    Call REGAPP("!CARFORMU", "CAN-COP Cantidad de Copias (hasta 6)")
    Call REGAPP("!CARFORMU", "IMP-DES Primera Copia con Importes (Factura)")
    Call REGAPP("!CARFORMU", "IMP-HAS Ultima Copia con Importes (Factura)")
    Call REGAPP("!CARFORMU", "DETDESMMBorde Superior de Detalle (milimetros)")
    Call REGAPP("!CARFORMU", "DETHASMMBorde Inferior de Detalle (milimetros)")
    Call REGAPP("!CARFORMU", "INTERLINInterlinea de Detalle (milimetros)")
    Call REGAPP("!CARFORMU", "ORIENTA Orientacion (Retrato / Paisaje)")
    Call REGAPP("!CARFORMU", "OPCIMPREPresentar Opciones Impresión (SI / NO)")
    Call REGAPP("!CARFORMU", ".")
    Call REGAPP("!CARFORMU", "MON-PRI Moneda Principal (PESOS - DOLARES)")
    Call REGAPP("!CARFORMU", "MON-CEN Moneda Secundaria (CENTAVOS)")
    Call REGAPP("!CARFORMU", "LOIMPLETLongitud Importe en Letras")
    Call REGAPP("!CARFORMU", "PREFIFACPrefijo de Facturación")
    Call REGAPP("!CARFORMU", "MAX-ITEMMaximo de Items por Documento")
    Call REGAPP("!CARFORMU", "LIN-SEP Linea de Separacion (Desde - Hasta)")
    Call REGAPP("!CARFORMU", ".")
    Call REGAPP("!CARFORMU", "CAR-OBSECampo Memo Observaciones (SI/NO)")
    Call REGAPP("!CARFORMU", ".")
    Call REGAPP("!CARFORMU", "FUEN-DEFFuente (Letra) por Omisión")
    '
    For U& = 1 To 128
        CODVAR$(U&) = ""
        REGFOR%(U&) = 0
    Next U&
    '
    CAREL& = ULTREG&("!CARFORMU")
    For U& = 1 To CAREL&
        CODVAR$(U&) = Left$(REGLEIDO$("!CARFORMU", U&), 8)
    Next U&
    '
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        CODVARI$ = Left$(RF$, 8)
        For J& = 1 To CAREL&
            If CODVAR$(J&) = CODVARI$ Then
                REGFOR%(J&) = I%
                NYY$ = RF$
                Call GRAREG("!CARFORMU", NYY$, J&)
            End If
        Next J&
    Next I%
    '
    VTB% = VENTABU%("CONFORMU/TITUPAN=Configuracion General: " + DESFOR$)
    If VTB% >= 0 Then
        For U& = 1 To ULTREG&("!CARFORMU")
           XXX$ = REGLEIDO$("!CARFORMU", U&)
           If REGFOR%(U&) < 1 Then
               REGFOR%(U&) = 1 + LOF(NX%) / 128
           End If
           RF$ = XXX$
           Put #NX%, REGFOR%(U&), RF$
        Next U&
    End If
    '
    Close #NX%: NX% = 0
    '
End Sub
'

Sub DEFLEYENW()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next I%
    '
    VTB% = VENTABU%("CARLEYEW")
    If VTB% >= 0 Then
        J% = 0
        For I% = 1 To LOF(NX%) / 128
            Get #NX%, I%, RF$
            If TRIM$(Left$(RF$, 8)) <> "TEXTWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next I%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "TEXTWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For I% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next I%
        '
    End If
    Close #NX%: NX% = 0
    '
End Sub
'

Sub DEFCAJASW()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next I%
    '
    VTB% = VENTABU%("CARCAJAW")
    If VTB% >= 0 Then
        J% = 0
        For I% = 1 To LOF(NX%) / 128
            Get #NX%, I%, RF$
            If TRIM$(Left$(RF$, 8)) <> "CAJAWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next I%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "CAJAWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For I% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next I%
        '
    End If
    Close #NX%: NX% = 0
    '
End Sub
'

Sub DEFLINEAW()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next I%
    '
    VTB% = VENTABU%("CARLINEW")
    If VTB% >= 0 Then
        J% = 0
        For I% = 1 To LOF(NX%) / 128
            Get #NX%, I%, RF$
            If TRIM$(Left$(RF$, 8)) <> "LINEWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next I%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "LINEWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For I% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next I%
        '
    End If
    Close #NX%: NX% = 0
    '
End Sub
'
Sub DEFIMAGEW()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "IMAGEWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next I%
    '
    VTB% = VENTABU%("CARIMAGW")
    If VTB% >= 0 Then
        J% = 0
        For I% = 1 To LOF(NX%) / 128
            Get #NX%, I%, RF$
            If TRIM$(Left$(RF$, 8)) <> "IMAGEWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next I%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "IMAGEWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For I% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next I%
        '
    End If
    Close #NX%: NX% = 0
    '
End Sub
'

Sub DEFPARAMW()
    '
    Dim RF As String * 128
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    Call BLANARCH("ECOPARAM")
    FORMALTER.LINOSORT.Clear
    FORMALTER.LINOSORT.AddItem "RASO-EMPRazon Social de la Empresa"
    FORMALTER.LINOSORT.AddItem "NOM-FANTNombre de Fantasia"
    FORMALTER.LINOSORT.AddItem "DOMI-EMPDomicilio"
    FORMALTER.LINOSORT.AddItem "CPOS-EMPCodigo Postal"
    FORMALTER.LINOSORT.AddItem "LOCA-EMPLocalidad"
    FORMALTER.LINOSORT.AddItem "TELE-EMPTelefonos"
    FORMALTER.LINOSORT.AddItem "CUIT-EMPNumero de C.U.I.T."
    FORMALTER.LINOSORT.AddItem "IBRU-EMPNumero de Ing.Brutos"
    FORMALTER.LINOSORT.AddItem "CAJU-EMPNumero de Caja Jubilacion"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "TIPO-COMTipo de Documento"
    FORMALTER.LINOSORT.AddItem "NUME-COMNumero de Documento"
    FORMALTER.LINOSORT.AddItem "DOCUORIGDocumento de origen"
    FORMALTER.LINOSORT.AddItem "FECH-EMIFecha de Emision"
    FORMALTER.LINOSORT.AddItem "HORA-EMIHora de Emision"
    FORMALTER.LINOSORT.AddItem "USUARIO Nombre Usuario FLEXOFT"
    FORMALTER.LINOSORT.AddItem "FECH-LANFecha de Lanzamiento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "FECH-VENFecha de Vencimiento / Entrega"
    FORMALTER.LINOSORT.AddItem "HOJA-NUMNumero de Orden de Hoja Actual"
    FORMALTER.LINOSORT.AddItem "HOJA-TOTTotal de Hojas del Documento"
    FORMALTER.LINOSORT.AddItem "ORD-COPIOrden de Copia (Original / Duplicado / etc.)"
    FORMALTER.LINOSORT.AddItem "DESTINO Destino (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "LIST-PRETexto Lista de Precios"
    FORMALTER.LINOSORT.AddItem "COND-PAGTexto Condicion de Pago"
    FORMALTER.LINOSORT.AddItem "VENDEDORTexto Vendedor"
    FORMALTER.LINOSORT.AddItem "ZONA-VENTexto Zona de Ventas"
    FORMALTER.LINOSORT.AddItem "DEPOSITODeposito de E/S - Lugar de Entrega"
    FORMALTER.LINOSORT.AddItem "OBSERVA Observaciones - Comentarios"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "COD-CONJCodigo de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "DES-CONJDescripcion de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "REF-CONJReferencia de Complementos (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "UNI-CONJUnidad del Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "COL-CONJColor de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TAL-CONJTalle de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "CAN-CONJCantidad de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "PES-CONJPeso en Kilos de Conjunto (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "COD-MPRICodigo de Materia Prima (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "DES-MPRIDescripcion M.Prima (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "UNI-MPRIUnidad de Mat. Prima (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "USO-MPRIUso Unitario de M.Prima (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "CAN-MPRICantidad de M.Prima (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TIP-MPRITipo de Material (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TIT-MPRITitulo del Hilado (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "COL-MPRIColor de Material (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "PRO-MPRIProveedor del Material (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "PAR-MPRILote-Partida del Material (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "UBI-MPRIUbicación del Material (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "OPE-PUEPOperacion (Puesta a Punto)"
    FORMALTER.LINOSORT.AddItem "MAQ-PUEPMaquina (Puesta a Punto)"
    FORMALTER.LINOSORT.AddItem "PAR-PUEPParámetro (Puesta a Punto)"
    FORMALTER.LINOSORT.AddItem "HER-PUEPHerramental (Puesta a Punto)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "COT-PLCOCota (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "REF-PLCOReferencia (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "MED-PLCOMedio de Control (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "CAN-PLCOCantidad (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "FRE-PLCOFrecuencia (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "REA-PLCOPlan de Reacción (Plan de Control)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "LAR-CUTELargo Cuerpo Tejeduria (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "ACH-CUTEAncho Cuerpo Tejeduria (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "LAR-MATELargo Manga Tejeduria (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "ACH-MATEAncho Manga Tejeduria (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "LAR-CUCOLargo Cuerpo Corte (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "ACH-CUCOAncho Cuerpo Corte (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "LAR-MACOLargo Manga Corte (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "ACH-MACOAncho Manga Corte (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "PIE-COMPPieza Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "DIB-COMPDibujo Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "MAQ-COMPMáquina Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "LAR-COMPLargo(cm) Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "ACH-COMPAncho(cm) Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "UNI-COMPUnidad Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "CAN-COMPCantidad Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "PES-COMPPeso(kg) Complemento (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "NOP-OPERNumero de Operacion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "DES-OPERDescripcion de Operacion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "OBS-OPERDescripcion Larga de Operacion (O.Fabricacion)"
    FORMALTER.LINOSORT.AddItem "EQUIPO1 Maquina / Equipo Nro.1 (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "EQUIPO2 Maquina / Equipo Nro.2 (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "EQUIPO3 Maquina / Equipo Nro.3 (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TIE-PREPHoras de Preparacion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TIE-PRODHoras de Produccion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TIE-TOTATotal de Horas (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "NUOPERS Numero de Operarios (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "TOT-PREPSuma Horas Preparacion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TOT-PRODSuma Horas Produccion (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "TOT-TIEMSuma Total de Horas (Orden de Fabricacion)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "NRO-REMINumero de Remito o Resumen de Entregas"
    FORMALTER.LINOSORT.AddItem "NRO-OCOMNumero de Orden de Compra"
    FORMALTER.LINOSORT.AddItem "NRO-AENTNumero de Aut.Entregas o Pedido"
    FORMALTER.LINOSORT.AddItem "IMP-LET1Importe en Letras - Primer Renglon"
    FORMALTER.LINOSORT.AddItem "IMP-LET2.                 - Segundo Renglon"
    FORMALTER.LINOSORT.AddItem "IMP-LET3.                 - Tercer Renglon"
    FORMALTER.LINOSORT.AddItem "."
    FORMALTER.LINOSORT.AddItem "TRANSPORTransportista"
    FORMALTER.LINOSORT.AddItem "DOMI-TRADomicilio del Transportista"
    FORMALTER.LINOSORT.AddItem "LOCA-TRALocalidad del Transportista"
    FORMALTER.LINOSORT.AddItem "CUIT-TRANumero de C.U.I.T."
    FORMALTER.LINOSORT.AddItem "PATE-TRANumero de Patente"
    FORMALTER.LINOSORT.AddItem "COND-TRAConductor"
    FORMALTER.LINOSORT.AddItem "DOCUCONDDocumento del Conductor"
    FORMALTER.LINOSORT.AddItem "."
    FORMALTER.LINOSORT.AddItem "VALO-DECValor Declarado (Remito)"
    FORMALTER.LINOSORT.AddItem "NUBULTOSNumero de Bultos (Remito)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "RASO-CLIRazon Social de Cliente / Proveedor"
    FORMALTER.LINOSORT.AddItem "DOMI-CLIDomicilio de Cliente / Proveedor"
    FORMALTER.LINOSORT.AddItem "CPOS-CLICodigo Postal Destinatario"
    FORMALTER.LINOSORT.AddItem "LOCA-CLILocalidad Destinatario"
    FORMALTER.LINOSORT.AddItem "PCIA-CLIProvincia Destinatario"
    FORMALTER.LINOSORT.AddItem "PAIS-CLIPaís Destinatario"
    FORMALTER.LINOSORT.AddItem "CTAC-CLIContacto de Cliente / Proveedor"
    FORMALTER.LINOSORT.AddItem "TELE-CLINúmero de Telefono Destinatario"
    FORMALTER.LINOSORT.AddItem "FAX-CLI Número de FAX Destinatario"
    FORMALTER.LINOSORT.AddItem "MAIL-CLIDirección de Mail Destinatario"
    FORMALTER.LINOSORT.AddItem "CUIT-CLINumero de C.U.I.T. Destinatario"
    FORMALTER.LINOSORT.AddItem "IBRU-CLINumero de Ingresos Brutos"
    FORMALTER.LINOSORT.AddItem "CUEN-CLINumero de Cuenta Cliente / Proveedor"
    FORMALTER.LINOSORT.AddItem "PROV-CLINumero de Proveedor"
    FORMALTER.LINOSORT.AddItem "PIVA-CLIPosicion de IVA Cliente / Proveedor"
    FORMALTER.LINOSORT.AddItem "CIVA-CLICategoria de IVA ( 'A' - 'B' - 'C' )"
    '\\\OT-06-0550-RG-13/12/06///
    FORMALTER.LINOSORT.AddItem "NTRA-CLINombre del Transportista (Cliente)"
    FORMALTER.LINOSORT.AddItem "RTRA-CLIRazon Social del Transportista (Cliente)"
    FORMALTER.LINOSORT.AddItem "DTRA-CLIDomicilio del Transportista (Cliente)"
    FORMALTER.LINOSORT.AddItem "LTRA-CLILocalidad del Transportista (Cliente)"
    '\\\OT-06-0550-RG-FIN///
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "ORD-ITEMNumero de Orden del Item"
    FORMALTER.LINOSORT.AddItem "CANTIDADCantidad del Item"
    FORMALTER.LINOSORT.AddItem "CAKILOS Kilaje Total del Item"
    FORMALTER.LINOSORT.AddItem "COD-MAT C¢digo de Articulo"
    FORMALTER.LINOSORT.AddItem "DES-MAT Descripcion del Material"
    FORMALTER.LINOSORT.AddItem "ESP-MAT Especificacion del Material"
    FORMALTER.LINOSORT.AddItem "UNIMED  Unidad de medida"
    FORMALTER.LINOSORT.AddItem "REF-MAT1Referencia - Especificacion - Despacho a Plaza"
    FORMALTER.LINOSORT.AddItem "REF-MAT2Referencia - 2"
    FORMALTER.LINOSORT.AddItem "REF-MAT3Referencia - 3"
    FORMALTER.LINOSORT.AddItem "REF-MAT4Referencia - 4"
    FORMALTER.LINOSORT.AddItem "REF-MAT5Referencia - 5"
    FORMALTER.LINOSORT.AddItem "REF-MAT6Referencia - 6"
    FORMALTER.LINOSORT.AddItem "NRO-LOTEIdentificador de Lote"
    FORMALTER.LINOSORT.AddItem "TALLE   Talle"
    FORMALTER.LINOSORT.AddItem "COLOR   Color"
    FORMALTER.LINOSORT.AddItem "MEDIDA-1Medida 1 - Largo"
    FORMALTER.LINOSORT.AddItem "MEDIDA-2Medida 2 - Ancho"
    FORMALTER.LINOSORT.AddItem "MEDIDA-3Medida 3 - Alto"
    FORMALTER.LINOSORT.AddItem "N-PLANO Numero de Plano"
    FORMALTER.LINOSORT.AddItem "R-PLANO Revision de Plano"
    FORMALTER.LINOSORT.AddItem "F-PLANO Fecha de Revision de Plano"
    FORMALTER.LINOSORT.AddItem "Z-PLANO Plazo de Plano"
    FORMALTER.LINOSORT.AddItem "N-HPROF Numero de Hoja de Procesos de Fabricacion"
    FORMALTER.LINOSORT.AddItem "N-HINSP Numero de Hoja de Inspeccion"
    FORMALTER.LINOSORT.AddItem "N-ORIGINNumero Original de Pieza (s/lista de precios"
    FORMALTER.LINOSORT.AddItem "N-OCOMPRNumero de Orden de Compra"
    FORMALTER.LINOSORT.AddItem "N-AUTENTNumero de Autorizacion de Entregas"
    FORMALTER.LINOSORT.AddItem "N-REMITONumero de Remito"
    FORMALTER.LINOSORT.AddItem "N-BOLRECNumero de Comprobante de Recepcion"
    FORMALTER.LINOSORT.AddItem "FECH-RECFecha de Recepcion / Entrega"
    FORMALTER.LINOSORT.AddItem "PRE-LISTPrecio Unitario de Lista"
    FORMALTER.LINOSORT.AddItem "PORDESCUPorcentaje de Descuento"
    FORMALTER.LINOSORT.AddItem "PRE-NETOPrecio Unitario Neto"
    FORMALTER.LINOSORT.AddItem "NET-ITEMImporte Neto DEL ITEM"
    FORMALTER.LINOSORT.AddItem "ALICUIVAAlicuota - Porcentaje de IVA"
    FORMALTER.LINOSORT.AddItem "IVA-ITEMImporte de IVA del Item"
    FORMALTER.LINOSORT.AddItem "BRU-ITEMImporte Bruto del Item"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "ENCA-DETEncabezado de Detalle de Facturacion"
    FORMALTER.LINOSORT.AddItem "ESPE-DETEspecificacion Final de Detalle"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "MONEDEMIMoneda de Emisión (Texto)"
    FORMALTER.LINOSORT.AddItem "TICAMBIOTipo de Cambio ($/U$S)"
    FORMALTER.LINOSORT.AddItem "IMP-NETOTotal Neto Factura"
    FORMALTER.LINOSORT.AddItem "POR-DESCPorcentaje de Descuento"
    FORMALTER.LINOSORT.AddItem "IMP-DESCImporte de Descuento"
    FORMALTER.LINOSORT.AddItem "IMPNEDESImporte Neto de Descuento"
    FORMALTER.LINOSORT.AddItem "ITO-GRA1Importe Gravado - Tasa Basica"
    FORMALTER.LINOSORT.AddItem "ITO-GRA2Importe Gravado - Tasa Secundaria"
    FORMALTER.LINOSORT.AddItem "ITO-GRAVImporte Total Gravado"
    FORMALTER.LINOSORT.AddItem "ITO-NOGRImporte Total No Gravado o Exento"
    FORMALTER.LINOSORT.AddItem "POR-IVA Porcentaje de I.V.A. - Tasa Básica"
    FORMALTER.LINOSORT.AddItem "POR-IVA1 Porcentaje de I.V.A. - Tasa Básica"
    FORMALTER.LINOSORT.AddItem "IMP-IVA Importe de I.V.A. - Tasa Básica"
    FORMALTER.LINOSORT.AddItem "IMP-IVA1Importe de I.V.A. - Tasa Secundaria"
    FORMALTER.LINOSORT.AddItem "IMP-RNI Importe de I.V.A. Resp.No Inscripto"
    FORMALTER.LINOSORT.AddItem "ITO-IVA Importe Total de I.V.A."
    FORMALTER.LINOSORT.AddItem "IMP-VARSImpuestos Varios"
    FORMALTER.LINOSORT.AddItem "PERC-IVAImporte Percepciones I.V.A."
    FORMALTER.LINOSORT.AddItem "PERC-BRUImporte Percepcion I.Brutos."
    FORMALTER.LINOSORT.AddItem "SUB-TOTAImporte Sub-Total o Transporte"
    FORMALTER.LINOSORT.AddItem "IMP-TOTAImporte Total"
    FORMALTER.LINOSORT.AddItem "TOTA-DOLTotal en Dolares"
    FORMALTER.LINOSORT.AddItem "."
    FORMALTER.LINOSORT.AddItem "LIN-CIE Linea de Cierre de Detalle"
    FORMALTER.LINOSORT.AddItem "TOTA-ITSTotal de Items"
    FORMALTER.LINOSORT.AddItem "LETOT-ITLeyenda 'Total de Items'"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "TIPO-VALTipo de Valor (Recibo)"
    FORMALTER.LINOSORT.AddItem "NUME-CHENumero de Cheque"
    FORMALTER.LINOSORT.AddItem "BANC-CHEBanco Emisor del Cheque"
    FORMALTER.LINOSORT.AddItem "IMPO-CHEImporte del Cheque"
    FORMALTER.LINOSORT.AddItem "FECH-CHEFecha Vencimiento de Cheque"
    FORMALTER.LINOSORT.AddItem "DOCU-APLDocumento de Aplicacion"
    FORMALTER.LINOSORT.AddItem "IMPO-PENSaldo Pendiente Documento"
    FORMALTER.LINOSORT.AddItem "IMPO-CANImporte Parcial Aplicado"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "TOTA-VALTotal Valores Recibidos (Recibo)"
    FORMALTER.LINOSORT.AddItem "TOTA-CANTotal Cancelado (Recibo)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "REF-ASIEReferencia (Asiento Contable)"
    FORMALTER.LINOSORT.AddItem "NRO-PASENúmero de Pase (As.Contable)"
    FORMALTER.LINOSORT.AddItem "COD-CUENCódigo de Cuenta (As.Contable)"
    FORMALTER.LINOSORT.AddItem "DEN-CUENDenominacion de la Cuenta"
    FORMALTER.LINOSORT.AddItem "DET-PASEDetalle - Referencia del Pase"
    FORMALTER.LINOSORT.AddItem "IMP-DEBEImporte Debe (As.Contable)"
    FORMALTER.LINOSORT.AddItem "IMP-HABEImporte Haber (As.Contable)"
    FORMALTER.LINOSORT.AddItem "TOT-DEBETotal Debe (Asiento Contable)"
    FORMALTER.LINOSORT.AddItem "TOT-HABETotal Haber (Asiento Contable)"
    FORMALTER.LINOSORT.AddItem "."
    '
    FORMALTER.LINOSORT.AddItem "CANT-MINStock Mínimo (Situación de Stocks)"
    FORMALTER.LINOSORT.AddItem "CANT-RESStock Comprometido (Situación de Stocks)"
    FORMALTER.LINOSORT.AddItem "CANT-INGStock a Ingresar (Situación de Stocks)"
    FORMALTER.LINOSORT.AddItem "CANT-DISStock Disponible (Situación de Stocks)"
    FORMALTER.LINOSORT.AddItem "CANT-FALStock Faltante (Situación de Stocks)"
    FORMALTER.LINOSORT.AddItem "."
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    For I% = 1 To LOF(NX%) / 128
      Get #NX%, I%, RF$
      If TRIM$(Mid$(RF$, 75, 6) + Mid$(RF$, 81, 6)) <> "" Then
        CODVARI$ = Left$(RF$, 8)
        For UI% = 1 To FORMALTER.LINOSORT.ListCount
          If Left$(FORMALTER.LINOSORT.List(UI% - 1), 8) = CODVARI$ Then
            TTXX$ = AJUSTI$(FORMALTER.LINOSORT.List(UI% - 1), 56)
            Call REPLA(TTXX$, "  *", 54, 3)
            FORMALTER.LINOSORT.RemoveItem (UI% - 1)
            FORMALTER.LINOSORT.AddItem TTXX$, UI% - 1
          End If
        Next UI%
      End If
    Next I%
    Close #NX%: NX% = 0
    '
    JJ& = 0
    For UI% = 1 To FORMALTER.LINOSORT.ListCount
      TTXX$ = FORMALTER.LINOSORT.List(UI% - 1)
      JJ& = JJ& + 1
      Call GRAREG("ECOPARAM", TTXX$, JJ&)
    Next UI%
    Call SETULTREG("ECOPARAM", JJ&)
    Call CIERRARCH("ECOPARAM")
    '
100 Call SELECHO("ECOPARAM")
    CODVARI$ = TRIM$(VALACT1$("ECOPARAM"))
    If CODVARI$ = "." Then GoTo 100
    If CODVARI$ = "" Then
        Call CIERRARCH("ECOPARAM")
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    DEVARI$ = VALACT2$("ECOPARAM")
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = CODVARI$ Then
            REGCO& = I%
            RFF$ = RF$
            GoTo 200
        End If
    Next I%
    RFF$ = String$(128, 0)
    Call REPLA(RFF$, CODVARI$, 1, 8)
    REGCO& = 1 + LOF(NX%) / 128
    '
200 Call REPLA(RFF$, DEVARI$, 9, 48)
    Call REPLA(RFF$, "w", 126, 1)
    '
    TXTX$ = ""
    XX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRQ") 'ASIGNO A XX$ LA LECTURA DEL ARCHIVO
    For U% = 1 To Len(XX$) Step 256 'LO RECORRO (256 ES LA LONGITUD DEL REGISTRO)
      YY$ = Mid$(XX$, U%, 256)
      If TRIM$(Left(YY$, 16)) = TRIM$(CODVARI$) Then
        TXTX$ = Mid$(XX$, U%, 256)
        Exit For
      End If
    Next U%
    '
    RFF$ = RFF$ + Mid$(TXTX$, 17)
    '
    RFG$ = OBJPLA$("DEFPARAMW", RFF$)
    If RFG$ = "" Then
        Close #NX%: NX% = 0
        GoTo 100
    End If
    '
    TTXxx$ = AJUSTI$(CODVARI$, 16) & Mid$(RFG$, 129)
    '
    XX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRQ")
    '
    For U% = 1 To Len(XX$) Step 256
      YY$ = Mid$(XX$, U%, 256)
      If TRIM$(Left(YY$, 16)) = TRIM$(CODVARI$) Then
        Mid$(XX$, U%, 256) = TTXxx$
        GoTo 220
      End If
    Next U%
    XX$ = XX$ + TTXxx$
220 Call SAVEFILE(LUDAT$ + CODFOR$ + ".FRQ", XX$)
    '
    RF$ = RFG$
    Put #NX%, REGCO&, RF$
    Close #NX%: NX% = 0
    GoTo 100
    '
End Sub
'

Sub CARFORWIN()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    If CAFUEN% = 0 Then
      On Error Resume Next
      For I% = 1 To Printer.FontCount
        FUEX$ = TRIM$(Printer.Fonts(I% - 1))
        If FUEX$ <> "" Then
          If InStr(FUEX$, "BENGUIAT") < 1 Then
            If CAFUEN% < 256 Then
              CAFUEN% = CAFUEN% + 1
              FUENTE$(CAFUEN%) = FUEX$
            End If
          End If
        End If
      Next I%
      On Error GoTo 0
    End If
    '
    LU$ = LUDAT$
    'Printer.ScaleMode = 1
    ORIENTA% = 1
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
      ORIENTA% = 2
    End If
    Call SETSPOOL("ORIENTATION", TRIM$(Str$(ORIENTA%)))
    '
    OPCIMPRE% = 0
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OPCIMPRE")))
    If Left$(PORI$, 1) = "S" Then OPCIMPRE% = 1
    '
    Call SETSPOOL("SCALEMODE", "1")
    COEFH = 56.7
    COEFV = 56.7
    On Error Resume Next
    ANCHOMAX = Printer.ScaleWidth
    If ANCHOMAX < 100 Then ANCHOMAX = 10000
    'Printer.ColorMode = 2   ' blanco y negro
    'Printer.ForeColor = QBColor(0)
    Call SETSPOOL("FORECOLOR", Str$(QBColor(0)))
    'Printer.DrawMode = 1    ' linea negra
    Call SETSPOOL("DRAWMODE", "13")
    '
    PHEIG& = 56.7 * Val(ATRIFORM$("", "LARGOJMM"))
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    '
' imprime imagenes excepto isologo
    RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
    If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
    End If
    '
    For I% = 1 To LOF(NX%) / 128
      Get #NX%, I%, RF$
      If TRIM$(Left$(RF$, 8)) = "IMAGEWIN" Then
        ARIMA$ = TRIM$(Mid$(RF$, 9, 48))
        If InStr(ARIMA$, ".") > 0 Then
          ARCHIMAG$ = RUTARGRA$ + ARIMA$
          If EXISTE%(ARCHIMAG$) = 1 Then
            Call PRINTIMAGE(ARIMA$, ARCHIMAG$)
          End If
        End If
      End If
    Next I%
    '
' dibuja cajas *************************************
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
            '
            'Printer.DrawWidth = 3
            Call SETSPOOL("DRAWWIDTH", "3")
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                'Printer.DrawWidth = Asc(Mid$(RF$, 73, 1))
                Call SETSPOOL("DRAWWIDTH", Str$(Asc(Mid$(RF$, 73, 1))))
            End If
            '
            'Printer.DrawStyle = 0
            Call SETSPOOL("DRAWSTYLE", "0")
            If Asc(Mid$(RF$, 74, 1)) < 5 Then
                'Printer.DrawStyle = Asc(Mid$(RF$, 74, 1))
                Call SETSPOOL("DRAWSTYLE", Str$(Asc(Mid$(RF$, 74, 1))))
            End If
            '
            'Printer.FillStyle = 1
            Call SETSPOOL("FILLSTYLE", "1")
            '
            H1 = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            V1 = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            H2 = (CVS(Mid$(RF$, 65, 4)) + MARIZW) * COEFH
            V2 = (CVS(Mid$(RF$, 69, 4)) + MARSUW) * COEFV
            'Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
            Call SPOOLINE(H1, V1, H2, V2, QBColor(0), "B")
            If H1 > EXTREDET Then EXTREDET = H1
            If H2 > EXTREDET Then EXTREDET = H2
            '
        End If
    Next I%
    '
' dibuja lineas *************************************
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
            '
            'Printer.DrawWidth = 1
            Call SETSPOOL("DRAWWIDTH", "1")
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                'Printer.DrawWidth = Asc(Mid$(RF$, 73, 1))
                Call SETSPOOL("DRAWWIDTH", Str$(Asc(Mid$(RF$, 73, 1))))
            End If
            '
            'Printer.DrawStyle = 0
            Call SETSPOOL("DRAWSTYLE", "0")
            If Asc(Mid$(RF$, 74, 1)) < 5 Then
                'Printer.DrawStyle = Asc(Mid$(RF$, 74, 1))
                Call SETSPOOL("DRAWSTYLE", Str$(Asc(Mid$(RF$, 74, 1))))
            End If
            '
            H1 = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            V1 = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            H2 = (CVS(Mid$(RF$, 65, 4)) + MARIZW) * COEFH
            V2 = (CVS(Mid$(RF$, 69, 4)) + MARSUW) * COEFV
            'Printer.Line (H1, V1)-(H2, V2), QBColor(0)
            Call SPOOLINE(H1, V1, H2, V2, QBColor(0), "")
            If H1 = H2 Then
              If H1 > EXTREDET Then
                EXTREDET = H1
              End If
            End If
            '
        End If
    Next I%
    '
'Exit Sub
' escribe leyendas fijas ***************************
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
            If TRIM$(Mid$(RF$, 9, 48)) = "" Then
                GoTo 109
            End If
            '
            ' esto es para habilitar impresion fantasma
            IFANTA% = 0
            If UCase$(Mid$(RF$, 124, 1)) = "F" Then IFANTA% = 1
            If IFANTA% > 0 Then
               If SENDBYMAIL$ = "" Then
                  If MUESTRAFORBLA% < 1 Then
                     GoTo 109
                  End If
               End If
            End If
            '
            'Printer.DrawStyle = 0
            Call SETSPOOL("DRAWSTYLE", "0")
            'Printer.ForeColor = QBColor(0)
            Call SETSPOOL("FORECOLOR", Str$(QBColor(0)))
            FUENX$ = TRIM$(Mid$(RF$, 87, 32))
            If FUENX$ = "" Then
                Call COMUNI("Fuente No Definida para\Leyenda '" + TRIM$(Mid$(RF$, 9, 48)) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
                FUEREPLA$ = "Courier"
                GoTo 103
            End If
            '
            For I9% = 1 To CAFUEN%
                If FUENTE$(I9%) = FUENX$ Then GoTo 105
            Next I9%
            If FUEREPLA$ = "" Then
                FUEREPLA$ = "Courier"
                TTXX$ = "Fuente '" + FUENX$ + "' no Disponible\en el Presente Modelo de Impresora\para Leyenda '" + TRIM$(Mid$(RF$, 9, 48)) + "'\en Formulario '" + TRIM$(CODFOR$) + "'."
                TTXX$ = TTXX$ + "\   \FLEXOFT Reemplazará las Fuentes\no Disponibles por '" + FUEREPLA$ + "'."
                Call COMUNI(TTXX$)
            End If
103         FUENX$ = FUEREPLA$
            '
105         'Printer.FontName = FUENX$
            Call SETSPOOL("FONTNAME", FUENX$)
            'Printer.FontItalic = False
            Call SETSPOOL("FONTITALIC", "0")
            TAFUEN = CVS(Mid$(RF$, 119, 4))
            If TAFUEN > 2 Then
                'Printer.FontSize = TAFUEN
                Call SETSPOOL("FONTSIZE", Str$(TAFUEN))
            End If
            If Asc(Mid$(RF$, 123, 1)) > 32 Then
                'Printer.FontBold = True
                Call SETSPOOL("FONTBOLD", "1")
              Else
                'Printer.FontBold = False
                Call SETSPOOL("FONTBOLD", "0")
            End If
            'Printer.CurrentX = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            Call SETSPOOL("CURRENTX", CSTRING$(MKS$((CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH + 0.5), 4, 12, 0, 1))
            'Printer.CurrentY = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            Call SETSPOOL("CURRENTY", CSTRING$(MKS$((CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFH + 0.5), 4, 12, 0, 1))
            'Printer.Print TRIM$(Mid$(RF$, 9, 48))
            Call SETSPOOL("PRINT", TRIM$(Mid$(RF$, 9, 48)))
            '
        End If
109 Next I%
    '
    RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
    If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
    End If
    '
    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
       ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
       If EXISTE%(ARCHIMAG$) = 1 Then
         Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
       End If
    End If
    '
    Close #NX%: NX% = 0
    On Error GoTo 0
    '
End Sub
'

Sub PRINTFORWIN(TEXT$, CODVARI$, X0, Y0)
    '
    Static NXP%
    Static RFP As String * 128
    Static CODFORA$, LNXPA%
    Static OFSECOPIA%
    '
    On Error Resume Next
    LNXP% = LOF(NXP%)
    On Error GoTo 0
    If LNXP% < 1 Then NXP% = 0
    If LNXP% <> LNXPA% Then NXP% = 0
    LNXPA% = LNXP%
    '
    'Printer.CurrentX = 0
    Call SETSPOOL("CURRENTX", "0")
    'Printer.CurrentY = 0
    Call SETSPOOL("CURRENTY", "0")
    '
    If TRIM$(TEXT$) = "" Then
        GoTo 900
        ' Exit Sub
    End If
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    If CAFUEN% = 0 Then
      For I% = 1 To Printer.FontCount
        FUEX$ = TRIM$(Printer.Fonts(I% - 1))
        If FUEX$ <> "" Then
          If InStr(FUEX$, "BENGUIAT") < 1 Then
            If CAFUEN% < 256 Then
              CAFUEN% = CAFUEN% + 1
              FUENTE$(CAFUEN%) = FUEX$
            End If
          End If
        End If
      Next I%
    End If
    '
    If CODFOR$ <> CODFORA$ Then
        If NXP% > 0 Then
            Close #NXP%: NXP% = 0
            NXP% = 0
            OFSECOPIA% = 0
        End If
        CODFORA$ = CODFOR$
    End If
    '
    If NXP% = 0 Then
        NXP% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
        For I% = 1 To LOF(NXP%) / 128
           Get #NXP%, I%, RFP$
           If TRIM$(Left$(RFP$, 8)) = "MARGENIN" Then
              OFSECOPIA% = Val(Mid$(RFP$, 89, 8))
           End If
       Next I%
    End If
    '
    For I% = 1 To LOF(NXP%) / 128
        Get #NXP%, I%, RFP$
        If TRIM$(Left$(RFP$, 8)) = CODVARI$ Then
            GoTo 200
        End If
    Next I%
    Exit Sub
    '
200 FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    If TRIM$(FORMA1$ + FORMA2$) = "" Then
        GoTo 900
    End If
    '
    If FUENX$ = "" Then
        Call MENSERR(24, "Fuente No Definida para\Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
        FUEREPLA$ = "Courier"
        GoTo 203
    End If
    '
    For I% = 1 To CAFUEN%
        If FUENTE$(I%) = FUENX$ Then GoTo 205
    Next I%
    If FUEREPLA$ = "" Then
        FUEREPLA$ = "Courier"
        TTXX$ = "Fuente '" + FUENX$ + "' no Disponible\en el Presente Modelo de Impresora\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'."
        TTXX$ = TTXX$ + "\   \FLEXOFT Reemplazará las Fuentes\no Disponibles por '" + FUEREPLA$ + "'."
        Call COMUNI(TTXX$)
    End If
203 FUENX$ = FUEREPLA$
    '
205 On Error GoTo 999
    'Printer.FontName = FUENX$
    Call SETSPOOL("FONTNAME", FUENX$)

    If CVS(Mid$(RFP$, 119, 4)) < 1 Then
        Call MENSERR(24, "Tamaño de Fuente Inexistente\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
        Exit Sub
    End If
    'Printer.FontSize = CVS(Mid$(RFP$, 119, 4))
    Call SETSPOOL("FONTSIZE", Str$(CVS(Mid$(RFP$, 119, 4))))
    On Error GoTo 0
    '
    If Asc(Mid$(RFP$, 123, 1)) > 32 Then
       'Printer.FontBold = True
       Call SETSPOOL("FONTBOLD", "1")
    End If
    If Asc(Mid$(RFP$, 124, 1)) > 32 Then
       'Printer.FontItalic = True
       Call SETSPOOL("FONTITALIC", "1")
    End If
    If Asc(Mid$(RFP$, 125, 1)) > 32 Then
       'Printer.FontUnderline = True
       Call SETSPOOL("FONTUNDERLINE", "1")
    End If
    '
    PH1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
    PV1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    '
    PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
    PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    '
    If FORMA1$ <> "" Then
        TTXX$ = TEXTFORM$(TEXT$, FORMA1$)
        If Left$(TTXX$, 1) = "." Then TTXX$ = " " + Mid$(TTXX$, 2)
        If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
           For CAREI% = 1 To CARETEX%
              'Printer.CurrentX = PH1
              Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH1 + 0.5), 4, 12, 0, 1))
              'Printer.CurrentY = PV1
              Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV1 + 0.5), 4, 12, 0, 1))
              'Printer.Print RETEX$(CAREI%);
              Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
              PV1 = PV1 + INTERLI * 56.7
           Next CAREI%
           GoTo 880
        End If
        If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
              TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
        End If
        If TRIM$(TTXX$) <> "" Then
            'Printer.CurrentX = PH1 ' - Printer.TextHeight(TTXX$)
            Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH1 + 0.5), 4, 12, 0, 1))
            'Printer.CurrentY = PV1
            Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV1 + 0.5), 4, 12, 0, 1))
            'Printer.Print "";
            Call SETSPOOL("PRINT", "" + ";")
            'Printer.CurrentX = PH1 ' - Printer.TextHeight(TTXX$)
            Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH1 + 0.5), 4, 12, 0, 1))
            'Printer.CurrentY = PV1
            Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV1 + 0.5), 4, 12, 0, 1))
            If TRIM$(TTXX$) <> "" Then
               'Printer.Print TTXX$;
               Call SETSPOOL("PRINT", TTXX$ + ";")
            End If
        End If
    End If
    '
880 If OFSECOPIA% < 1 Then
       If FORMA2$ <> "" Then
          TTXX$ = TEXTFORM$(TEXT$, FORMA2$)
          If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
            For CAREI% = 1 To CARETEX%
              'Printer.CurrentX = PH1
              Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH1 + 0.5), 4, 12, 0, 1))
              'Printer.CurrentY = PV1
              Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV1 + 0.5), 4, 12, 0, 1))
              'Printer.Print RETEX$(CAREI%);
              Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
              PV1 = PV1 + INTERLI * 56.7
            Next CAREI%
            GoTo 890
          End If
          If TRIM$(TTXX$) <> "" Then
              'Printer.CurrentX = PH2 ' - Printer.TextHeight(TTXX$)
              Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH2 + 0.5), 4, 12, 0, 1))
              'Printer.CurrentY = PV2
              Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV2 + 0.5), 4, 12, 0, 1))
              'Printer.Print "";
              Call SETSPOOL("PRINT", "" + ";")
              'Printer.CurrentX = PH2 ' - Printer.TextHeight(TTXX$)
              Call SETSPOOL("CURRENTX", CSTRING$(MKS$(PH2 + 0.5), 4, 12, 0, 1))
              'Printer.CurrentY = PV2
              Call SETSPOOL("CURRENTY", CSTRING$(MKS$(PV2 + 0.5), 4, 12, 0, 1))
              If TRIM$(TTXX$) <> "" Then
                 'Printer.Print TTXX$;
                 Call SETSPOOL("PRINT", TTXX$ + ";")
              End If
          End If
       End If
890 End If
    '
900 'Printer.FontBold = False
    Call SETSPOOL("FONTBOLD", "0")
    'Printer.FontItalic = False
    Call SETSPOOL("FONTITALIC", "0")
    'Printer.FontUnderline = False
    Call SETSPOOL("FONTUNDERLINE", "0")
    '
999 On Error GoTo 0
    '
End Sub

Function NUMTEX$(VALOR#)
Dim UND$(19), DECA$(9), CENTO$(19): GoSub 11690

11000 NUMERO$ = "": VALOR# = (Int(100 * VALOR# + 0.5)) / 100
11010 ENT# = Int(VALOR#)
11020 DECI# = Int(100 * (VALOR# - ENT#) + 0.5)
11030 MILLO# = Int(ENT# / 1000000!)
11040 MILI# = Int((ENT# - MILLO# * 1000000!) / 1000)
11050 UND1# = Int(ENT# - MILLO# * 1000000! - MILI# * 1000)
11060 BLOQUE# = MILLO#
11070 If BLOQUE# = 0 Then GoTo 11130
11080 GoSub 11270
11090 If BLOQUE# > 1 Then GoTo 11120
11100 NUMERO$ = BLOQ$ + " MILLON "
11110 GoTo 11130
11120 NUMERO$ = BLOQ$ + " MILLONES "
11130 BLOQUE# = MILI#
11140 If BLOQUE# = 0 Then GoTo 11170
11150 GoSub 11270
11160 NUMERO$ = NUMERO$ + BLOQ$ + " MIL "
11170 BLOQUE# = UND1#
11180 If BLOQUE# = 0 Then GoTo 11210
11190 GoSub 11270
11200 NUMERO$ = NUMERO$ + BLOQ$
11205 If Right$(NUMERO$, 2) = "UN" Then NUMERO$ = NUMERO$ + "O"
11207 MPRI$ = ATRIFORM$("", "MON-PRI")
11210 If TRIM$(MPRI$) <> "" Then
          NUMERO$ = NUMERO$ + " " + MPRI$
      End If
      NUMERO$ = NUMERO$ + " CON "
11220 BLOQUE# = DECI#: GoSub 11270
11225 MCEN$ = ATRIFORM$("", "MON-CEN")
11230 NUMERO$ = NUMERO$ + BLOQ$
      If TRIM$(MCEN$) <> "" Then
          NUMERO$ = NUMERO$ + " " + MCEN$
      End If
11240 GoTo 40000
11260 '
11270 BLOQ$ = ""
11280 B1 = Int(BLOQUE# / 100)
11290 B2 = Int((BLOQUE# - B1 * 100) / 10)
11300 B3 = Int(BLOQUE# - B1 * 100 - B2 * 10)
11310 If B2 >= 2 Then GoTo 11340
11320 B3 = B2 * 10 + B3
11330 B2 = 0
11340 If B1 = 0 Then GoTo 11370
11345 '
11350 BLOQ$ = CENTO$(B1)
11360 If ((B2 <> 0) Or (B3 <> 0)) And (B1 = 1) Then BLOQ$ = BLOQ$ + "TO "
11370 If B2 = 0 Then GoTo 11420
11380 BLOQ$ = BLOQ$ + DECA$(B2)
11390 If (B3 = 0) And (B2 = 2) Then BLOQ$ = BLOQ$ + "E"
11400 If (B3 <> 0) And (B2 = 2) Then BLOQ$ = BLOQ$ + "I"
11410 If (B3 <> 0) And (B2 > 2) Then BLOQ$ = BLOQ$ + " Y "
11420 If B3 = 0 Then GoTo 11440
11430 BLOQ$ = BLOQ$ + UND$(B3)
11440 If ((B1 = 0) And (B2 = 0) And (B3 = 0)) Then If BLOQUE# = DECI# Then BLOQ$ = "CERO"
11450 Return
11460 '
11690 DTUNI$ = " UN, DOS, TRES, CUATRO, CINCO, SEIS, SIETE, OCHO, NUEVE, DIEZ, ONCE, DOCE, TRECE, CATORCE, QUINCE, DIECISEIS, DIECISIETE, DIECIOCHO, DIECINUEVE"
11700 DTDEC$ = ", VEINT, TREINTA, CUARENTA, CINCUENTA, SESENTA, SETENTA, OCHENTA, NOVENTA"
11710 DTCEN$ = " CIEN, DOSCIENTOS , TRESCIENTOS , CUATROCIENTOS , QUINIENTOS , SEISCIENTOS , SETECIENTOS , OCHOCIENTOS , NOVECIENTOS "
11720 'RESTORE 11690
11730 For I = 1 To 19
        PPXX% = InStr(DTUNI$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTUNI$)
        UND$(I) = TRIM$(Left$(DTUNI$, PPXX% - 1))
        DTUNI$ = Mid$(DTUNI$, PPXX% + 1)
      Next I
11740 For I = 1 To 9
        PPXX% = InStr(DTDEC$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTDEC$)
        DECA$(I) = TRIM$(Left$(DTDEC$, PPXX% - 1))
        DTDEC$ = Mid$(DTDEC$, PPXX% + 1)
      Next I
11750 For I = 1 To 9
        PPXX% = InStr(DTCEN$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTCEN$)
        CENTO$(I) = TRIM$(Left$(DTCEN$, PPXX% - 1))
        If I > 1 Then CENTO$(I) = CENTO$(I) + " "
        DTCEN$ = Mid$(DTCEN$, PPXX% + 1)
      Next I
11760 Return
11770 '
40000 NUMTEX$ = NUMERO$
40010 End Function



