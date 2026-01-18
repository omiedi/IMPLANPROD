Attribute VB_Name = "CONEXI2"
Sub VERNOVTRA()
   '
   Dim SSXX As String * 256
   Call CARUTRA
   ULTIMPOR$ = Mid$(REGLEIDO$("STOREMOT", 1), 3, 64)
   For U& = 0 To 3
     ASTOREMO$ = TRIM$(ORITRA$(U&))
     If ASTOREMO$ <> "" Then
       FEGENER$ = ""
       On Error Resume Next
       FEGENER$ = TRIM$(FileDateTime(ASTOREMO$))
       On Error GoTo 0
       If FEGENER$ <> "" Then
         If FEGENER$ <> TRIM$(CONTROL$("", "FEGTRA-" + TRIM$(Str$(U& + 1)))) Then
           NK% = FILEOPEN%(ASTOREMO$, 0, 256)
           AAA = LOF(NK%)
           If LOF(NK%) >= 256 Then
             Get #NK%, 1, SSXX$
             If TRIM$(Mid$(SSXX$, 17, 16)) <> TRIM$(Mid$(ULTIMPOR$, 16 * U& + 1, 16)) Then
               Close #NK%: NK% = 0
               If COMALTER%("Hay Novedades en los Archivos\de Transferencia de Stocks Remotos.\  \FLEXOFT Recomienda Importar la Nueva Situación de Stocks\a Fin de Actualizar la Disponibilidad General.\\Importar Ahora\Continuar") = 1 Then
                 Call ITRASTOK
               End If
               Exit Sub
             End If
           End If
           Close #NK%: NK% = 0
         End If
       End If
     End If
   Next U&
   '
End Sub
'
Sub ITRASTOK()
   '
   Dim SSXX As String * 256
   Call BLOQARCH("STOREMOT")
   Screen.MousePointer = 11
   Z$ = MKI$(0) + Space$(126)
   Call GRAREG("STOREMOT", Z$, 1)
   JJ& = 1
   '
   For U& = 0 To 3
      ASTOREMO$ = TRIM$(ORITRA$(U&))
      If ASTOREMO$ <> "" Then
         NK% = FILEOPEN%(ASTOREMO$, 0, 256)
         FIN& = LOF(NK%) / 256
         Get #NK%, 1, SSXX$
         FEGENTRA$ = Mid$(SSXX$, 17, 16)
         For TU& = 2 To FIN&
            Call TRACE(20, TU&, FIN&)
            Get #NK%, TU&, SSXX$
            X$ = SSXX$
            SEDORIG$ = Left$(X$, 8)
            TIREGIS$ = UCase$(TRIM$(Mid$(X$, 9, 8)))
            If TIREGIS$ = "DISPOREM" Then
               CODD$ = Mid$(X$, 17, 16)
               TIMO$ = Mid$(X$, 33, 2)
               REFE$ = Mid$(X$, 35, 50)
               NUCO$ = Mid$(X$, 85, 12)
               CANTU# = CVD(Mid$(X$, 97, 8)) - CVD(Mid$(X$, 105, 8))
               FEPRO% = CVI(Mid$(X$, 113, 2))
               '
               CIIY% = CODINT%(CODD$)
               '
               Z$ = REGBLAN$("STOREMOT")
               Call REPLA(Z$, MKI$(CIIY%), 1, 2)
               Call REPLA(Z$, NUCO$, 3, 12)
               Call REPLA(Z$, TIMO$, 15, 2)
               Call REPLA(Z$, REFE$, 17, 50)
               Call REPLA(Z$, MKD$(CANTU#), 67, 8)
               Call REPLA(Z$, MKI$(FEPRO%), 83, 2)
               Call REPLA(Z$, SEDORIG$, 85, 8)
               JJ& = JJ& + 1
               Call GRAREG("STOREMOT", Z$, JJ&)
               '
            End If
         Next TU&
         '
         If FIN& > 0 Then
            Z$ = REGLEIDO$("STOREMOT", 1)
            Call REPLA(Z$, FEGENTRA$, 3 + 16 * U&, 16)
            Call GRAREG("STOREMOT", Z$, 1)
         End If
         Close #NK%
         If FIN& > 0 Then
            FEGENER$ = TRIM$(FileDateTime(ASTOREMO$))
            Call GRACONTROL("", "FEGTRA-" + TRIM$(Str$(U& + 1)), FEGENER$)
         End If
      End If
   Next U&
   '
   Call SETULTREG("STOREMOT", JJ&)
   Call CIERRARCH("STOREMOT")
   '
   Screen.MousePointer = 1
   '
End Sub

