Attribute VB_Name = "SELF_EXECUTE"
Sub AUTOEXEC()
    '
    LANGUA% = 0
    VERDEMO% = (-1)
    INTERPRE% = 0
    '
    IDENTER$ = "000000"
    NOMTER$ = "Estación Local"
    '
    ULOAC$ = Left$(CurDir, 2)
    LUBAS$ = CurDir + "\"
    LUCOM$ = CurDir + "\"
    LUDAT$ = CurDir + "\"
    LUBOOT$ = Left$(CurDir, 1)
    '
    If App.EXEName = "FLEXO00" Then
      LUBAS$ = "C:\FLEXOWIN\"
      LUCOM$ = "C:\FLEXOFT\"
      LUDAT$ = "C:\F_LIGHT\"
    End If
    '
    USNBR% = 101
    USERNAME$ = "Usuario General"
    MENUPRIN$ = ""
    MODINVO$ = ""
    DENOMOD$ = ""
    MODMENU$ = ""
    '
    EMPREAC$ = "Su Empresa - " + VERNUM$ + " - " + USERNAME$
    '
End Sub
'
Sub TRANSLABELS(FORNAM$)
   '
   If LANGUA% = 0 Then Exit Sub
   '
   For IFORMU% = 1 To Forms.Count
     If Forms(IFORMU% - 1).Name = FORNAM$ Then
       '
       CAPTI$ = ""
       On Error Resume Next
       CAPTI$ = TRIM$(Forms(IFORMU% - 1).Caption)
       Forms(IFORMU% - 1).Caption = TRANSLATE$(CAPTI$)
       On Error GoTo 0
       '
       For KKI% = 1 To Forms(IFORMU% - 1).Controls.Count
         '
10       CAPTI$ = ""
         On Error GoTo 91
         CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).Caption)
         Forms(IFORMU% - 1).Controls(KKI% - 1).Caption = TRANSLATE$(CAPTI$)
         For KKJ% = 1 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(KKJ%) = TRANSLATE$(CAPTI$)
         Next KKJ%
         On Error GoTo 0
         '
20       TOTIP$ = ""
         On Error GoTo 92
         TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText)
         Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText = TRANSLATE$(TOTIP$)
         For KKJ% = 1 To 99
           TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(KKJ%) = TRANSLATE$(TOTIP$)
         Next KKJ%
         On Error GoTo 0
         '
30       CAPTI$ = ""
         On Error GoTo 93
         For KKJ% = 0 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(KKJ%) = TRANSLATE$(CAPTI$)
         Next KKJ%
         On Error GoTo 0
         '
80       On Error Resume Next
         Forms(IFORMU% - 1).Controls(KKI% - 1).Refresh
         On Error GoTo 0
         '
       Next KKI%
     End If
   Next IFORMU%
   Exit Sub
   '
91 Resume 20
92 Resume 30
93 Resume 80
   '
End Sub

Function TRANSLATE$(TESPA$)
   '
   Static FITRA%
   Static BUFTRA As String * 1024
   Static URETRA&
   Static XINDI$
   '
   TRANSLATE$ = TESPA$
   If LANGUA% < 1 Or LANGUA% > 9 Then Exit Function
   If TRIM$(TESPA$) = "" Then Exit Function
   If LUBAS$ = "" Then Exit Function
   '
   If FITRA% < 1 Then
     Screen.MousePointer = 11
     FITRA% = FreeFile
     Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
     URETRA& = LOF(FITRA%) / 1024
     XINDI$ = String$(2 * URETRA&, 0)
     For U& = 1 To URETRA&
       Get #FITRA%, U&, BUFTRA$
       Mid$(XINDI$, 2 * U& - 1, 2) = Left$(BUFTRA$, 2)
     Next U&
     Screen.MousePointer = 1
   End If
   '
   BVT% = BINVAL%(TESPA$)
50 BVT% = BINVAL%(TESPA$)
   AZ$ = MKI$(BVT%)
   RFF$ = "": PINI& = 1
60 POSI& = InStr(PINI&, XINDI$, AZ$)
   If POSI& < 1 Then GoTo 70                           ' NEXT RGX&
   If (POSI& Mod 2) <> 1 Then
      PINI& = POSI& + 1
      GoTo 60
   End If
   PINI& = POSI& + 2
   RGL& = (POSI& + 1) / 2
   RFF$ = RFF$ + MKS$(RGL&)
   GoTo 60
   '
70 For KKI% = 1 To Len(RFF$) Step 4
     REIN& = CVS(Mid$(RFF$, KKI%, 4))
     If REIN& > 0 Then
       If REIN& <= URETRA& Then
         Get #FITRA%, REIN&, BUFTRA$
         TINDI$ = TRIM$(Mid$(BUFTRA$, 3, 126))
         If TINDI$ = TESPA$ Then
           TRALA$ = TRIM$(Mid$(BUFTRA$, 128 * LANGUA% + 1, 128))
           If TRALA$ <> "" Then TRANSLATE$ = TRALA$
           Exit Function
         End If
       End If
     End If
   Next KKI%
   '
End Function


