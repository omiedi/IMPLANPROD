Attribute VB_Name = "FISECOM_RUTS"
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
   MAILMESS.Text4 = TRIM$(MMT4$ + EMPRE$)
     PPXX% = InStr(FISECOM.Text5, "/")
     If PPXX% < 1 Then PPXX% = 1 + Len(FISECOM.Text5)
   MAILMESS.Text5 = TRIM$(Left$(FISECOM.Text5, PPXX% - 1))
   MAILMESS.Text20 = CONTROL$("SEGUICOM", "ASUNTO")
   '
   HOII% = HOY(HOS$)
   MAILMESS.Label5 = "Buenos Aires, " + FETEXTO$(HOII%) + ".-"
   While MAILMESS.Label5.Width < MAILMESS.Text2.Width - 1500
     MAILMESS.Label5 = " " + MAILMESS.Label5
   Wend
   MAILMESS.Label5 = AJUSTD$(MAILMESS.Label5, 82)
   '
   TTYY$ = MAILMESS.Label5 + Chr$(13) + Chr$(10)
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
   '
   MAILMESS.Text2 = TTYY$ ' = TTYY$ + Text1
   MAILMESS.Text2.SelStart = SST
   MAILMESS.Text2.SelLength = 0
   '
   RFF$ = RECFILT$("ANOTASEG", 1, MKS$(NPX&))
   MAILMESS.List1.Clear
   '
   For U& = 1 To Len(RFF$) Step 4
     RENOTA& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("ANOTASEG", RENOTA&)
     For KU% = 5 To 92
       If Asc(Mid$(XX$, KU%)) < 32 Then
         Mid$(XX$, KU%, 1) = " "
       End If
     Next KU%
     MAILMESS.List1.AddItem Mid$(XX$, 5, 92)
   Next U&
   '
   TPU& = MAILMESS.List1.ListCount - 7
   If TPU& < 0 Then TPU& = 0
   MAILMESS.List1.TopIndex = TPU&
   '
End Sub

