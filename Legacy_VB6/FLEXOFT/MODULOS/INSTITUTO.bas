Attribute VB_Name = "INSTITUTO"
Sub ANOTAX_ALUMNO(LEGAJO&, TXTX$, Optional FF%, Optional MODO$)
    '
    
    If LEGAJO& > 0 Then
       '
       US% = USNBR% Mod 100
       ExLEGAJO& = CantRegistros&("DEUDOR12", "NUME_CLI = " & LEGAJO&) ' VALIDACION DE LEGAJO
       If ExLEGAJO& < 1 Then
          If MODO$ <> "NOMESS" Then
             Call MENSERR(24, "LEGAJO: " & LEGAJO& & " Inexistente")
          Else
             Exit Sub
          End If
       End If
       '
       HOII% = HOY(HOS$)
       fech$ = HOS$
       If FF% > 0 Then
          fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
       End If
       '
       Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
       '
       For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
           TTYY$ = Space$(96)
           If KKI% = 1 Then
              Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
              Call REPLA(TTYY$, USERNAM$(US%), 69, 24)
           End If
           Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
           ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
       Next KKI%
       '
       ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
       '
       CONDI$ = "LEGAJO = " & LEGAJO&
       '
       ANOTRIOR$ = VALOBADA("ANOTA_ALUMNO", "ANOTACION", CONDI$, "NOMESS")  'LA NOTA GUARDADA EN LA TABLA
       '
       ANOTA$ = ANOTRIOR$ & ANOTANTE$   ' CONCATENO LA ANTERIOR CON LA NUEVA
       '
       'GRABO LA NOTA
       SQLX$ = "SELECT * FROM ANOTA_ALUMNO "
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       '
       Set AUX = New ADODB.Recordset
       AUX.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With AUX
          If .EOF Then
           .AddNew
           !LEGAJO = LEGAJO&
          End If
          !ANOTACION = ANOTA$
          .Update
       End With
        
       AUX.Close
       Set AUX = Nothing
       
       End If
    
    '
End Sub

Sub ANOTAX_PROFESORES(LEGAJO&, TXTX$, Optional FF%, Optional MODO$)
    '
    
    If LEGAJO& > 0 Then
       '
       ExLEGAJO& = CantRegistros&("PROFESORES", "NUME_CLI = " & LEGAJO&) ' VALIDACION DE LEGAJO
       If ExLEGAJO& < 1 Then
          If MODO$ <> "NOMESS" Then
             Call MENSERR(24, "LEGAJO: " & LEGAJO& & " Inexistente")
          Else
             Exit Sub
          End If
       End If
       '
       HOII% = HOY(HOS$)
       fech$ = HOS$
       If FF% > 0 Then
          fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
       End If
       '
       Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
       '
       For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
           TTYY$ = Space$(96)
           If KKI% = 1 Then
              Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
              Call REPLA(TTYY$, NOMBRE_PC$, 69, 24)
           End If
           Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
           ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
       Next KKI%
       '
       ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
       '
       CONDI$ = "LEGAJO = " & LEGAJO&
       '
       ANOTRIOR$ = VALOBADA("ANOTA_PROFESOR", "ANOTACION", CONDI$, "NOMESS")  'LA NOTA GUARDADA EN LA TABLA
       '
       ANOTA$ = ANOTRIOR$ & ANOTANTE$   ' CONCATENO LA ANTERIOR CON LA NUEVA
       '
       'GRABO LA NOTA
       SQLX$ = "SELECT * FROM ANOTA_PROFESOR "
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       '
       Set AUX = New ADODB.Recordset
       AUX.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       With AUX
          If .EOF Then
           .AddNew
           !LEGAJO = LEGAJO&
          End If
          !ANOTACION = ANOTA$
          .Update
       End With
        
       AUX.Close
       Set AUX = Nothing
       
       End If
    
    '
End Sub



