Attribute VB_Name = "Module1"
Public Function EXISTEA(ARCHX$)
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
    EXISTEA = EXI%
    On Error GoTo 0
    '
End Function
'

