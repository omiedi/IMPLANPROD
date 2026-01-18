Attribute VB_Name = "DIRETRA1"
Public GRABA%
'
Dim DIAPLI As String
Dim LUCO1 As String
Dim LUDA1 As String
Dim LUBA1 As String
Dim LUDAX As String
Dim LUBAX As String
'

Sub DIRETRA()
    '
    DIRETRAB.Show
    DIRETRAB.MousePointer = 11
    '
    LCB$ = LUBOOT$ + ":"
    DITRA$ = ULODATOS$
    DIAPLI$ = UCase$(Mid$(CurDir, 3))
    LUCO1$ = UCase$(TRIM$(LUCOM$))
    LUDA1$ = UCase$(TRIM$(LUDAT$))
    If Mid$(LUDA1$, 2, 1) = ":" Then
       LUDA1$ = Mid$(LUDA1$, 3)
    End If
    '
    LUBA1$ = UCase$(TRIM$(LUBAS$))
    If Mid$(LUBA1$, 2, 1) = ":" Then
       LUBA1$ = Mid$(LUBA1$, 3)
    End If
    '
    If Len(LUCO1$) > 1 Then
       If Right$(LUCO1$, 1) = "\" Then
          LUCO1$ = Left$(LUCO1$, Len(LUCO1$) - 1)
       End If
    End If
    '
    If Len(LUDA1$) > 1 Then
       If Right$(LUDA1$, 1) = "\" Then
          LUDA1$ = Left$(LUDA1$, Len(LUDA1$) - 1)
       End If
    End If
    '
    If LUDA1$ = "" Then LUDA1$ = DIAPLI$
    On Error GoTo CAPERR1
    NX% = FreeFile
    Open LUDA1$ + "\MASTER.DAT" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    '
    If Len(LUBA1$) > 1 Then
       If Right$(LUBA1$, 1) = "\" Then
          LUBA1$ = Left$(LUBA1$, Len(LUBA1$) - 1)
       End If
    End If
    '
    If LUBA1$ = "" Then
       LUBA1$ = TRIM$(UCase$(Mid$(CurDir, 3)))
       LDDX$ = AJUSTI$(TRIM$(LUBA1$) + "\", 46)
    End If
    On Error GoTo CAPERR2
    NX% = FreeFile
    Open LUBA1$ + "\OBJNAME.DRV" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX%: NX% = 0
    '
    If EXISTE%(DIAPLI$ + "\LANGUAGE.DRV") > 0 Then
      On Error Resume Next
      LANGUA% = Val(LOADFILE$(DIAPLI$ + "\LANGUAGE.DRV"))
      On Error GoTo 0
    End If
    If LANGUA% < 0 Or LANGUA% > 5 Then LANGUA% = 0

    DIRETRAB.Text1.TEXT = LCB$
    DIRETRAB.Text2.TEXT = DIAPLI$
    DIRETRAB.Text3.TEXT = LUCO1$
    DIRETRAB.Text4.TEXT = LUDA1$
    DIRETRAB.Text5.TEXT = LUBA1$
    DIRETRAB.Text6.TEXT = DITRA$
    DIRETRAB.Option1(LANGUA%).Value = True
    '
    DIRETRAB.MousePointer = 1
    DIRETRAB.Text4.SetFocus
    GoTo FINSUB
    
CAPERR1:
    LUDA1$ = TRIM$(UCase$(Mid$(CurDir, 3)))
    Resume Next
    '
CAPERR2:
    LUBA1$ = Mid$(CurDir, 3)
    Resume Next
    '
FINSUB:
End Sub

Sub GRADIRE()
    '
    GRABA% = 0
    LUDAX$ = TRIM$(DIRETRAB.Text4.TEXT)
    On Error GoTo CAPERR3
    NX% = FreeFile
    Open LUDAX$ + "\MASTER.DAT" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    If LUDAX$ <> TRIM$(DIRETRAB.Text4.TEXT) Then
        Call MENSERR(24, "Carpeta no Válida")
        DIRETRAB.Text4.TEXT = LUDA1$
        DIRETRAB.Text4.SetFocus
        Exit Sub
      Else
        LUDA1$ = LUDAX$
    End If
    '
    LUBAX$ = TRIM$(DIRETRAB.Text5.TEXT)
    On Error GoTo CAPERR4
    NX% = FreeFile
    Open LUBAX$ + "\OBJNAME.DRV" For Random Access Read Write Shared As #NX% Len = 128
    On Error GoTo 0
    Close #NX% = NX% = 0
    If LUBAX$ <> TRIM$(DIRETRAB.Text5.TEXT) Then
        Call MENSERR(24, "Carpeta no Válida")
        DIRETRAB.Text5.TEXT = LUBA1$
        DIRETRAB.Text5.SetFocus
        Exit Sub
      Else
        LUBA1$ = LUBAX$
    End If
    '
grabadi:
    '
    NX% = FreeFile
    Open "FLEXCFG.DAT" For Random Access Read Write Shared As #NX% Len = 128
    Dim FF0 As String * 128
    '
    Get #NX%, 1, FF0$
    XXX$ = FF0$
    Call REPLA(XXX$, DIAPLI$, 3, 46)
    Call REPLA(XXX$, LUDA1$, 51, 46)
    FF0$ = XXX$
    Put #NX%, 1, FF0$
    '
    LDDX$ = AJUSTI$(TRIM$(LUDA1$) + "\", 46)
    AA11$ = TRIM$(DIRETRAB.Text6.TEXT)
    Call GRARAM(99, AA11$ + LDDX$)
    LUDAT$ = AA11$ + TRIM$(LDDX$)
    '
    Get #NX%, 3, FF0$
    XXX$ = FF0$
    Call REPLA(XXX$, LUBA1$, 1, 128)
    FF0$ = XXX$
    Put #NX%, 3, FF0$
    LDDX$ = AJUSTI$(TRIM$(LUBA1$) + "\", 46)
    Call GRARAM(145, LDDX$)
    LUBAS$ = TRIM$(LDDX$)
    If Mid$(LUBAS$, 2, 1) <> ":" Then
        ULOAC$ = Left$(CurDir, 2)
        LUBAS$ = ULOAC$ + LUBAS$
    End If
    '
    Close #NX%: NX% = 0
    '
    AA11$ = TRIM$(DIRETRAB.Text6.TEXT)
    On Error Resume Next
    NN11% = FreeFile
    Open "C:\FLEXOFT\ULOGDAT.DRV" For Output As #NN11%
    Print #NN11%, AA11$
    Close #NN11%
    On Error GoTo 0
    '
    IDTX$ = IDENTER$
    ULXX$ = LOADFILE$("ULODATOS.DRV")
    For KKI% = 1 To Len(ULXX$) Step 16
      ULYY$ = Mid$(ULXX$, KKI%, 16)
      If TRIM$(Left$(ULYY$, 14)) = IDTX$ Then
        Mid$(ULYY$, 15, 2) = AA11$
        Mid$(ULXX$, KKI%, 16) = ULYY$
        GoTo 100
      End If
    Next KKI%
    ULXX$ = ULXX$ + AJUSTI$(IDTX$, 14) + AA11$
100 Call SAVEFILE("ULODATOS.DRV", ULXX$)
    '
    '
    LANGUA% = 0
    For KKKI% = 0 To 5
      If DIRETRAB.Option1(KKKI%).Value = True Then
        LANGUA% = KKKI%
      End If
    Next KKKI%
    Call SAVEFILE(DIAPLI$ + "\LANGUAGE.DRV", TRIM$(Str$(LANGUA%)))
    Call GRARAM(305, TRIM$(Str$(LANGUA%)))
    Call BAJRAM
    '
    GRABA% = 1
  Exit Sub
  '
CAPERR3:
    LUDAX$ = LUDA1$
    Resume Next
    '
CAPERR4:
    LUBAX$ = LUBA1$
    Resume Next
    '
End Sub
'
