Attribute VB_Name = "PADR_FUN_SQL"
Function LEGA_OPER&(ID$)        ' Trae el numero de legajo del empleado cuyo ID_SYS es igual a ID$
    '
    LEGA_OPER& = 0
    If ID$ <> "" Then
       CONDI$ = "Legajo = '" & ID$ & "'"
       LEGA_OPER = XVALO(VALOBADA("PADRON", "LEGAJO", CONDI$, "NOMESS"))
       If LEGA_OPER = "" Then
          CONDI$ = "ID_SYS = '" & ENCRIPTA$(ID$) & "'"
          LEGA_OPER& = XVALO(VALOBADA("PADRON", "LEGAJO", CONDI$, "NOMESS"))
       End If
    End If
    '
End Function

Function NOME_OPER$(ID$)        ' Trae el nombre completo  del empleado cuyo ID_SYS es igual a ID$
    '
    NOME_OPER$ = ""
    If ID$ <> "" Then
       CONDI$ = "Legajo = '" & ID$ & "'"
       NOME_OPER$ = VALOBADA("PADRON", "NOMBRE", CONDI$, "NOMESS")
       If NOME_OPER = "" Then
          CONDI$ = "ID_SYS = '" & ENCRIPTA$(ID$) & "'"
          NOME_OPER$ = VALOBADA("PADRON", "NOMBRE", CONDI$, "NOMESS")
       End If
       '
       If TRIM$(NOME_OPER) = "" Then        ' NO LO ENCONTRO EN PADRON - BUSCA EN USER.DAT
          ID1$ = TRIM$(ID$)
          VUSER$ = LOADFILE$("USER.DAT")
          CAUS% = Len(VUSER$) / 64
          For JJ% = 1 To CAUS%
             F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
             Ident$ = Mid$(F0$, 31, 12)
             If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
             If UCase$(LTrim$(RTrim$(Ident$))) = UCase$(ID1$) Then
                DENO$ = Left$(F0$, 30)
                If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
                NOME_OPER$ = TRIM$(DENO$)
                Exit Function
             End If
          Next JJ%
       End If
       '
    End If
    '
End Function

Function CCAT_OPER$(ID$)        ' Trae el codigo de la categoría del empleado cuyo ID_SYS es igual a ID$
    '
    VALIMCAT = 0
    ID1$ = ID$
    PPXX% = InStr(ID$, ">")
    If PPXX% > 0 Then
       ID1$ = Left$(ID$, PPXX% - 1)
       VALIMCAT = XVALO(Mid$(ID$, PPXX% + 1))
    End If
    '
    CCAT_OPER$ = ""
    If ID1$ <> "" Then
       CONDI$ = "Legajo = '" & ID1$ & "'"
       CCAT_OPER$ = VALOBADA("PADRON", "CATEGORIA", CONDI$, "NOMESS")
       If CCAT_OPER = "" Then
          CONDI$ = "ID_SYS = '" & ENCRIPTA$(ID1$) & "'"
          CCAT_OPER$ = VALOBADA("PADRON", "CATEGORIA", CONDI$, "NOMESS")
       End If
    End If
    '
    If VALIMCAT > 0 Then
       If XVALO(CCAT_OPER$) < VALIMCAT Then
          CATI$ = CCAT_OPER$
          If CATI$ <> "" Then
            Call MENSERR(24, "Categoría Informada (" & CATI$ & ") Menor a la Requerida (" & VALIMCAT & ").")
          End If
       End If
    End If
    '
End Function

Function DCAT_OPER$(ID$)        ' Trae LA DESCRIPCION de la categoría del empleado cuyo ID_SYS es igual a ID$
    '
    DCAT_OPER$ = ""
    If ID$ <> "" Then
       DCAT_OPER$ = ECOARCH$("CATEGO", CCAT_OPER$(ID$))
    End If
    '
End Function
Function NOMBRE_USUARIO$(US%)
   NOMBRE_USUARIO$ = ""
   If US% > 0 Then
    RFF$ = RECFILT$("FLEXUSR", 1, MKI$(US%)) ' SE DEBE CARGAR LA CUENTA DE EFECTIVO PRIMERA EN LOS  MEDIOS DE COBRANZA
    If Len(RFF$) >= 4 Then
      REME& = CVS(Mid$(RFF$, 1, 4))
      If REME& > 0 Then
        NOMBRE_USUARIO = TRIM$(Mid$(REGLEIDO$("FLEXUSR", REME&), 3, 30))
      End If
    End If
   End If
  

End Function
