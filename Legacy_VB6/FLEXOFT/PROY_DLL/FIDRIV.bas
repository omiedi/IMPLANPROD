Attribute VB_Name = "FIDRIV"
Function EXISTE%(ARCHX$)
    '
    ARXX$ = Trim$(ARCHX$)
    If InStr(ARXX$, ".") < 1 Then ARXX$ = ARXX$ + ".DAT"
    If Right$(ARXX$, 1) = "." Then ARXX$ = Left$(ARXX$, Len(ARXX$) - 1)
    '
    EXI% = 0
    If Trim$(ARXX$) <> "" Then
      On Error Resume Next
      EXISZ$ = Trim$(Dir(ARXX$))
      If Err < 1 Then
        If EXISZ$ <> "" Then
          EXI% = 1
        End If
      End If
    End If
    '
    EXISTE% = EXI%
    On Error GoTo 0
    '
End Function
'

