Attribute VB_Name = "NUM_CONVERT"

Function CVI(A$)
    '
    X = 0
    CVI = 0
    If Len(A$) >= 2 Then
       X = 256& * Asc(Mid$(A$, 2)) + Asc(Left$(A$, 1))
       If X > 32767 Then X = X - 65536
    End If
    If X <= 32767 Then
       If X >= (-32768) Then
           CVI = X
       End If
    End If
    '
End Function
'

Function MKI$(AAABA)
    '
    AA& = XVALO(AAABA)
    If AA& < 0 Then AA& = AA& + 65536
    MKI$ = Chr$(AA& Mod 256) + Chr$(AA& \ 256)
    '
End Function
'

Function CVS(A$)
    '
    Static A0 As String * 4
    Static CTX%
    '
    CVS = 0
    If CTX% = 0 Then
      CTX% = 1
      A0$ = String$(4, 0)
    End If
    '
    If Len(A$) < 4 Then Exit Function
    If A$ = A0$ Then
      Exit Function
    End If
    '
    Dim AA(4)
    Dim CCC As Single
    For k% = 1 To 4: AA(k%) = Asc(Mid$(A$, k%)): Next k%
    SIGNO = Sgn(128 - AA(3))
    If SIGNO = 0 Then SIGNO = (-1)
    If AA(4) = 0 Then SIGNO = 0
    '
    A41 = Int((AA(4) - 129) / 2): A42 = (AA(4) - 129) - A41
    N4 = 2 ^ A42
    N4 = N4 * 2 ^ A41
    N4 = N4 * SIGNO
    '
    A31 = AA(3) Mod 128
    N3 = A31 / 128: N3 = N3 * N4
    N2 = AA(2) / 128 / 256: N2 = N2 * N4
    n1 = AA(1) / 128 / 256 / 256: n1 = n1 * N4
    '
    CCC = N4 + N3 + N2 + n1
    CVS = CCC
    '
End Function
'

Function MKS$(n)
    '
    Dim AA(4)
    For k% = 1 To 4: AA(k%) = 0: Next k%
    On Error Resume Next
    SIGNO = Sgn(n)
    If Err Then
      n = 0: SIGNO = 0
    End If
    On Error GoTo 0
    '
    If SIGNO <> 0 Then
        '
        NX = Abs(n)
        If NX >= 1 Then
             AA(4) = 128
             While NX >= 1: NX = NX / 2: AA(4) = AA(4) + 1: Wend
           Else
             AA(4) = 129
             While NX < 1: NX = NX * 2: AA(4) = AA(4) - 1: Wend
        End If
        '
        A41 = Int((AA(4) - 129) / 2): A42 = (AA(4) - 129) - A41
        N4 = 2 ^ A42
        N4 = N4 * 2 ^ A41
        '
        DIFE = Abs(n) - N4
        AA(3) = Int(128 * DIFE / 2 ^ (AA(4) - 129))
        If SIGNO < 0 Then AA(3) = AA(3) + 128
        '
        A31 = AA(3) Mod 128
        N3 = A31 / 128: N3 = N3 * N4
        '
        DIFE = Abs(n) - N4 - N3
        AA(2) = 128 * DIFE / 2 ^ (AA(4) - 129)
        AA(2) = Int(256 * AA(2))
        '
        N2 = AA(2) / 128 / 256: N2 = N2 * N4
        DIFE = Abs(n) - N4 - N3 - N2
        AA(1) = 128 * DIFE / 2 ^ (AA(4) - 129)
        AA(1) = 256 * AA(1)
        AA(1) = Int(256 * AA(1))
        '
    End If
    '
    MKS$ = ""
    For k% = 1 To 4
       'If AA(K%) < 0 Then AA(K%) = 0
       'If AA(K%) > 255 Then AA(K%) = 255
       MKS$ = MKS$ + Chr$(AA(k%))
    Next k%
    '
End Function
'

Function CVD(A$)
    '
    CVD = 0
    If Len(A$) < 8 Then Exit Function
    '
    Dim AA(8)
    Dim CCC As Double
    For k% = 1 To 8: AA(k%) = Asc(Mid$(A$, k%)): Next k%
    SIGNO = Sgn(128 - AA(7))
    If SIGNO = 0 Then SIGNO = (-1)
    If AA(8) = 0 Then SIGNO = 0
    '
    A41 = Int((AA(8) - 129) / 2): A42 = (AA(8) - 129) - A41
    N8# = 2 ^ A42
    N8# = N8# * 2 ^ A41
    N8# = N8# * SIGNO
    '
    A31 = AA(7) Mod 128
    N7# = A31 / 128: N7# = N7# * N8#
    N6# = AA(6) / 128 / 256: N6# = N6# * N8#
    N5# = AA(5) / 128 / 256 / 256: N5# = N5# * N8#
    N4# = AA(4) / 128 / 256 / 256 / 256: N4# = N4# * N8#
    N3# = AA(3) / 128 / 256 / 256 / 256 / 256: N3# = N3# * N8#
    N2# = AA(2) / 128 / 256 / 256 / 256 / 256 / 256: N2# = N2# * N8#
    n1# = AA(1) / 128 / 256 / 256 / 256 / 256 / 256 / 256: n1# = n1# * N8#
    '
    CCC# = N8# + N7# + N6# + N5# + N4# + N3# + N2# + n1#
    CVD = CCC#
    '
End Function
'

Function MKD$(n#)
    '
    Dim AA(10)
    For k% = 1 To 10: AA(k%) = 0: Next k%
    '
    On Error Resume Next
    SIGNO = Sgn(n#)
    If Err Then
        n = 0: SIGNO = 0
    End If
    On Error GoTo 0
    If SIGNO = 0 Then
         MKD$ = String$(8, 0)
         Exit Function
    End If
    '
    NX# = Abs(n#)
    If NX# >= 1 Then
        AA(8) = 128
        While NX# >= 1: NX# = NX# / 2: AA(8) = AA(8) + 1: Wend
       Else
        AA(8) = 129
        While NX# < 1: NX# = NX# * 2: AA(8) = AA(8) - 1: Wend
    End If
    '
    A81 = Int((AA(8) - 129) / 2): A82 = (AA(8) - 129) - A81
    N8# = 2# ^ A82
    N8# = N8# * 2# ^ A81
    '
    DIFE# = Abs(n#) - N8#
    AA(7) = Int(128 * DIFE# / 2# ^ (AA(8) - 129))
    If SIGNO < 0 Then AA(7) = AA(7) + 128
    '
    A71 = AA(7) Mod 128
    N7# = A71 / 128: N7# = N7# * N8#
    '
    DIFE# = DIFE# - N7#    '  ABS(N#)-N8#-N7#
    AA(6) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(6) = Int(256 * AA(6))
    N6# = N8# * AA(6) / 128 / 256
    '
    DIFE# = DIFE# - N6#  '  ABS(N#)-N8#-N7#-N6#
    AA(5) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(5) = 256 * AA(5)
    AA(5) = Int(256 * AA(5))
    N5# = N8# * AA(5) / 128 / 256 / 256
    '
    DIFE# = DIFE# - N5#  '    ABS(N#)-N8#-N7#-N6#-N5#
    AA(4) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(4) = 256 * AA(4)
    AA(4) = 256 * AA(4)
    AA(4) = Int(256 * AA(4))
    N4# = N8# * AA(4) / 128 / 256 / 256 / 256
    '
    DIFE# = DIFE# - N4#        ' ABS(N#)-N8#-N7#-N6#-N5#-N4#
    AA(3) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(3) = 256 * AA(3)
    AA(3) = 256 * AA(3)
    AA(3) = 256 * AA(3)
    AA(3) = Int(256 * AA(3))
    N3# = N8# * AA(3) / 128 / 256 / 256 / 256 / 256
    '
    DIFE# = DIFE# - N3#        ' ABS(N#)-N8#-N7#-N6#-N5#-N4#
    AA(2) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(2) = 256 * AA(2)
    AA(2) = 256 * AA(2)
    AA(2) = 256 * AA(2)
    AA(2) = 256 * AA(2)
    AA(2) = Int(256 * AA(2))
    N2# = N8# * AA(2) / 128 / 256 / 256 / 256 / 256 / 256
    '
    DIFE# = DIFE# - N2#        ' ABS(N#)-N8#-N7#-N6#-N5#-N4#
    AA(1) = 128 * DIFE# / 2# ^ (AA(8) - 129)
    AA(1) = 256 * AA(1)
    AA(1) = 256 * AA(1)
    AA(1) = 256 * AA(1)
    AA(1) = 256 * AA(1)
    AA(1) = 256 * AA(1)
    AA(1) = Int(256 * AA(1))
    '
    MKD$ = ""
    For k% = 1 To 8
        While AA(k%) < 0: AA(k%) = AA(k%) + 256: AA(k% + 1) = AA(k% + 1) - 1: Wend
        'If AA(K%) > 255 Then AA(K%) = 255
        MKD$ = MKD$ + Chr$(AA(k%))
    Next k%
    '
End Function
'
Function FENUM%(FZ$)
    FNN = XVALO(Left$(FZ$, 2)) + 32 * XVALO(Mid$(FZ$, 4, 2)) + 32 * 13 * (100 + XVALO(Mid$(FZ$, 7, 2)) - 80)
    If FNN > 0 Then
      If FNN < 32767 Then
        FENUM% = FNN
      End If
    End If
End Function

Function XVALO(valor) As Double
   If IsNull(valor) = True Then
        XVALO = 0
      Else
        VTP% = VarType(valor)
        If VTP% = 8 Then
             VALOR2$ = valor
             VALOR2$ = REPLACAR$(VALOR2$, "%", "")
             PPX1% = InStr(VALOR2$, ".")
             PPX2% = InStr(VALOR2$, ",")
             If PPX1% > 0 And PPX2% > 0 And PPX1% < PPX2% Then
                VALOR2$ = REPLACAR$(VALOR2$, ".", "")
                AAA = BBB
             End If
             VALOR1 = REPLACAR$(VALOR2$, ",", ".")
             XVALO = Val(VALOR1)
           Else
             On Error Resume Next
             XVALO = valor
             If Err Then
                XVALO = 0
             End If
             On Error GoTo 0
        End If
   End If
End Function


