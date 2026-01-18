Attribute VB_Name = "FISEGVEN_RUTS"
Sub CREAMAIL(NPX&)
   '
   IFICHA% = NPX&
   REG& = REGBUS&("BSEGCOM", 1, MKI$(IFICHA%))
   If REG& < 1 Or REG& > ULTREG&("BSEGCOM") Then Exit Sub
   '
   RELE$ = REGLEIDO$("BSEGCOM", REG&)
   '
   EMPRE$ = TRIM$(Mid$(RELE$, 3, 59))
   CALNUM$ = TRIM$(Mid$(RELE$, 129, 48))
   CPOST$ = TRIM$(Mid$(RELE$, 177, 8))
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + TRIM$(Mid$(RELE$, 185, 32))
   PROPA$ = TRIM$(Mid$(RELE$, 217, 32))
   If PROPA$ <> "" Then PROPA$ = PROPA$ + " - "
   PROPA$ = TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 249, 24)))
   '
   CADEST% = 0
   TTXX$ = TRIM$(Mid$(RELE$, 449, 32))
   If TTXX$ <> "" Then
      CADEST% = CADEST% + 1
      DESTIX$ = TTXX$
   End If
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PROPA$ <> "" Then TTXX$ = TTXX$ + PROPA$ + Chr$(13) + Chr$(10)
   '
   MMT4$ = TRIM$(DESTIX$)
   If MMT4$ <> "" Then MMT4$ = MMT4$ + " - "
   MAILMESSVE.Text4 = TRIM$(MMT4$ + EMPRE$)
     PPXX% = InStr(FISEGVEN.Text5, "/")
     If PPXX% < 1 Then PPXX% = 1 + Len(FISEGVEN.Text5)
   MAILMESSVE.Text5 = TRIM$(Left$(FISEGVEN.Text5, PPXX% - 1))
   MAILMESSVE.Text20 = CONTROL$("SEGUICOM", "ASUNTO")
   '
   hoii% = HOY(HOS$)
   MAILMESSVE.Label5 = "Buenos Aires, " + FETEXTO$(hoii%) + ".-"
   While MAILMESSVE.Label5.Width < MAILMESSVE.Text2.Width - 1500
     MAILMESSVE.Label5 = " " + MAILMESSVE.Label5
   Wend
   '
   TTYY$ = MAILMESSVE.Label5 + Chr$(13) + Chr$(10)
   If CADEST% < 1 Then
       TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
    Else
       TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
       TTYY$ = TTYY$ + DESTIX$ + Chr$(13) + Chr$(10)
   End If
   '
   TTYY$ = TTYY$ + TTXX$
   SST = 4 + Len(TTYY$)
   '
   For KKII% = 1 To 5: TTYY$ = TTYY$ + Chr$(13) + Chr$(10): Next KKII%
   TTYY$ = TTYY$ + LOADFILE$(LUDAT$ + "PIEDEMAI.DAT")
   'TTYY$ = TTYY$ + Space$(75) + "Ing. Ernesto P. Bauer" + Chr$(13) + Chr$(10)
   'TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
   'TTYY$ = TTYY$ + "FLEXOFT Soluciones Informáticas Integrales" + Chr$(13) + Chr$(10)
   'TTYY$ = TTYY$ + "Bauness 2170 - 5to. 'D' - 1431 Buenos Aires - Argentina" + Chr$(13) + Chr$(10)
   'TTYY$ = TTYY$ + "Tel (00-54-11) 4524-1284 / 4541-1515 - e-mail: info@flexoft.com.ar" + Chr$(13) + Chr$(10)
   'TTYY$ = TTYY$ + "===================================================================" + Chr$(13) + Chr$(10)
   '
   MAILMESSVE.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESSVE.Text2.SelStart = SST
   MAILMESSVE.Text2.SelLength = 0
   '
   RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
   MAILMESSVE.List1.Clear
   '
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
     For KU% = 5 To 92
       If Asc(Mid$(XX$, KU%)) < 32 Then
         Mid$(XX$, KU%, 1) = " "
       End If
     Next KU%
     MAILMESSVE.List1.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
   TPU& = MAILMESSVE.List1.ListCount - 7
   If TPU& < 0 Then TPU& = 0
   MAILMESSVE.List1.TopIndex = TPU&
   '
End Sub

Sub MUESTRADOC(DOCUACT$)
    '
    On Error Resume Next
    WORDDOCS.OLE1.SourceDoc = DOCUACT$
    TIDOCU$ = "WORD"
    If UCase$(Right$(DOCUACT$, 4)) = ".EML" Then
        TIDOCU$ = "MAIL"
      ElseIf UCase$(Right$(DOCUACT$, 4)) = ".XLS" Then
        TIDOCU$ = "EXCL"
      ElseIf UCase$(Right$(DOCUACT$, 4)) = ".ZIP" Then
        TIDOCU$ = "ZIPF"
    End If
    '
    If TIDOCU$ = "ZIPF" Then
      RETURNVALUE = Shell("C:\Archivos de programa\WinZip\WINZIP32.EXE", 4)
      AppActivate RETURNVALUE: AppActivate RETURNVALUE
      SendKeys "^O", True
      SendKeys DOCUACT$ + Chr$(13), True
      GoTo 99
    End If
      '
    UNILOG$ = "": DOCUNE$ = WORDDOCS.OLE1.SourceDoc
    If Mid$(DOCUNE$, 2, 1) = ":" Then
      UNILOG$ = Left$(DOCUNE$, 2)
      DOCUNE$ = Mid$(DOCUNE$, 3)
    End If
    '
    WORDDOCS.OLE1.Action = 1
    '
    WORDDOCS.OLE1.Action = 6
    WORDDOCS.VScroll1.Min = 0
    WORDDOCS.VScroll1.Value = 0
    ABARRA = WORDDOCS.OLE1.Height - WORDDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
    WORDDOCS.VScroll1.Max = ABARRA
    WORDDOCS.VScroll1.SmallChange = ABARRA / 100
    WORDDOCS.VScroll1.LargeChange = ABARRA / 10
    WORDDOCS.OLE1.Top = 0
    WORDDOCS.OLE1.DoVerb (-1)
    If TIDOCU$ = "WORD" Then
        SendKeys "%{F10}", True  ' esto para documentos en word
      ElseIf TIDOCU$ = "MAIL" Then
        SendKeys Chr$(13) + "%{F10}", True ' esto para mails
      ElseIf TIDOCU$ = "EXCL" Then
        SendKeys Chr$(13) & "%V" & (DOWN8) & "P", True ' ESTO PARA EXCEL
    End If
    '
    WORDDOCS.Show 1
    '
99  On Error GoTo 0
    '
End Sub


