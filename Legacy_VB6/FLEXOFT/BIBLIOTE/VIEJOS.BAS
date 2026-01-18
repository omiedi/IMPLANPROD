Attribute VB_Name = "VIEJOS"
Function VERDEMOOLD%()
    '
    Static VDEMO%
    If VDEMO% = 0 Then
VDEMO% = (-1): GoTo 99                       ' despues sacar
        SN$ = SERIVOL$("C:")
        QQQ& = 143
        For U& = 1 To 9
            QQQ& = QQQ& + (7 * U&) ^ 2 * Asc(Mid$(SN$, U&, 1))
        Next U&
VDEMO% = (-1): GoTo 99                       ' despues sacar
        If QQQ& = 826136 Or QQQ& = 741268 Then
            VDEMO% = (-1)                       ' mi maquina
            GoTo 99
        End If
        
        If Len(SN$) <> 9 Then
            VDEMO% = 1
            GoTo 99
        End If
        SE$ = ENCRIPTA$(SN$)
        NX% = FreeFile
        Dim RFX As String * 113
        Open "FLEXWIN.DLL" For Random Access Read Shared As #NX% Len = 113
        For I% = 1 To LOF(NX%) / 113
            POSI% = I% Mod 91
            Get #NX%, I%, RFX$
            If Mid$(RFX$, POSI%, 9) = SE$ Then
                VDEMO% = (-1)
                Close #NX%
                GoTo 99
            End If
        Next I%
        Close #NX%
        VDEMO% = 1
        '
    End If
    '
99  VERDEMOOLD% = VDEMO%
    '
End Function
'
        Function EMPREACOLD$()
        '
        EC$ = LEERAM$(193, 30)
        If TRIM$(EC$) = "" Then
            EC$ = EMPRELI$
        End If
        EMPREACOLD$ = TRIM$(EC$)
        '
        End Function

' ***************************************************************************


