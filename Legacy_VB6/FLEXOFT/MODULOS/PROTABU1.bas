Attribute VB_Name = "PROTABU1"
Public COMMA$
Sub Main()
    '
    COMMA$ = TRIM$(UCase$(Command$))
    If COMMA$ = "BRSTOCK" Then Call BRSTOCK
    '
End Sub
'

Sub BRSTOCK()
    '
    FIN& = ULTREG&("BRSTOCK")
    If FIN& > 0 Then
        '
        PROBOLRE.ProgressBar1.Min = 0
        PROBOLRE.ProgressBar1.Max = FIN&
        '
        PROBOLRE.Show
        PROBOLRE.Refresh
        '
        J& = 0
        For I& = 1 To FIN&
            '
            PROBOLRE.ProgressBar1.Value = I&
            XX$ = REGLEIDO$("BRSTOCK", I&)
            CANOMI# = CVD(Mid$(XX$, 58, 8))
            CACEPT# = CVD(Mid$(XX$, 66, 8))
            CARECH# = CVD(Mid$(XX$, 66, 8))
            CAFALT# = CVD(Mid$(XX$, 66, 8))
            '
            If CANOMI# - CACEPT# - CARECH# - CAFALT# >= 0.05 Then
                '
                NBR = CVS(Mid$(XX$, 3, 4))
                FEX = CVI(Mid$(XX$, 1, 2))
                NP% = CVI(Mid$(XX$, 7, 2))
                CI% = CVI(Mid$(XX$, 9, 2))
                '
                NUBO$ = AJUSTI$("BR-" + TRIM$(Str$(NBR)), 12)
                FE$ = FECHATEX$(FEX)
                PRO$ = TRIM$(RASOCLI$((-1) * NP%))
                If TRIM$(PRO$) = "" Then
                    PRO$ = "Proveedor" + Str$(NP%)
                End If
                If CI% > 0 And CI% <= NUMAS% Then
                      COD$ = TRIM$(CODEXT$(CI%))
                      DE$ = DESCRIT$(CI%)
                    Else
                      COD$ = "S/Codigo"
                      DE$ = "Material Libre"
                End If
                REFE$ = FE$ + " - " + PRO$ + " - " + COD$ + " - " + DE$
                '
                YY$ = REGBLAN$("BOLREPE")
                Call REPLA(YY$, NUBO$, 1, 12)
                Call REPLA(YY$, REFE$, 13, 72)
                Call REPLA(YY$, "BRSTOCK", 85, 8)
                Call REPLA(YY$, MKS$(I&), 93, 4)
                Call REPLA(YY$, XX$, 97, 96)
                '
                J& = J& + 1
                Call GRAREG("BOLREPE", YY$, J&)
                '
            End If
            '
        Next I&
        '
        For K& = J& + 1 To ULTREG&("BOLREPE")
            Call GRAREG("BOLREPE", String$(512, 0), K&)
        Next K&
        '
    End If
    '
    Call CIERRARCH("*.*")
    PROBOLRE.Hide
    Close: End
    '
End Sub

