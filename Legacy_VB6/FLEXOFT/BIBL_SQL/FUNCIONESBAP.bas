Attribute VB_Name = "FUNCIONESBAP"
Function CANTIRESERVA&(IDFUNCTION&)
  CANTIRESERVA& = XVALO(VALOBADA("BOOKINGS", "SUM(CANTIDAD)", "IDFUNCION = " & IDFUNCTION& & " AND ESTADO >= 0", "NOMESS"))
End Function
Function CANTILIBERADA&(IDFUNCTION&)
  CANTILIBERADA& = XVALO(VALOBADA("BOOKINGS", "SUM(LIBERADOS)", "IDFUNCION = " & IDFUNCTION& & " AND ESTADO >= 0", "NOMESS"))
End Function
Function RESERVAMASLIBERADA&(IDFUNCTION&)
  RESERVAMASLIBERADA& = CANTIRESERVA&(IDFUNCTION&) + CANTILIBERADA&(IDFUNCTION&)
End Function

Function DESCRI_INSTITU$(IDINSTITUCION&)
   DESCRI_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "SCHOOLS", "ID = " & IDINSTITUCION&, "NOMESS"))
End Function


Function DESCRI_PARTICULARES$(IDPARTICULAR&)
   DESCRI_PARTICULARES$ = TRIM$(VALOBADA("PARTICULARES", "SCHOOLS", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function

Function DESCRISTAT$(IDSTATUS%)
   DESCRISTAT$ = TRIM$(VALOBADA("ESTADOS", "ESTADO", "ID = " & IDSTATUS%, "NOMESS"))
End Function
Function NUMSTATUS%(ESTADOX$)
   CONDI$ = "substring(ESTADO,1,32) = '" & UCase(ESTADOX$) & "'"
   NUMSTATUS% = XVALO(VALOBADA("ESTADOS", "ID", CONDI$, "NOMESS"))
End Function


Function FECHAFUNCION(IDFUNCTION)
   FECHAFUNCION = TRIM$(VALOBADA("FUNCIONES", "FECHA", "IDFUNCION = " & IDFUNCTION, "NOMESS"))
End Function
'
Function EMAIL_PART$(IDPARTICULAR&)
   EMAIL_PART$ = TRIM$(VALOBADA("PARTICULARES", "EMAIL", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
'
Function EMAIL_PART2$(IDPARTICULAR&)
   EMAIL_PART2$ = TRIM$(VALOBADA("PARTICULARES", " [mail 2]", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
Function TELE_PART$(IDPARTICULAR&)
   TELE_PART$ = TRIM$(VALOBADA("PARTICULARES", "TELEPHONES", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
Function CIUDAD_PART$(IDPARTICULAR&)
   CIUDAD_PART$ = TRIM$(VALOBADA("PARTICULARES", "TOWN", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
Function PROVINCIA_PART$(IDPARTICULAR&)
   PROVINCIA_PART$ = TRIM$(VALOBADA("PARTICULARES", "PROVINCE", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function

Function CPOST_PART$(IDPARTICULAR&)
   CPOST_PART$ = TRIM$(VALOBADA("PARTICULARES", "[Zip Code]", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
Function DIREC_PART$(IDPARTICULAR&)
   DIREC_PART$ = TRIM$(VALOBADA("PARTICULARES", "DOMICILIO", "ID = " & IDPARTICULAR&, "NOMESS"))
End Function
Function EMAIL_INSTITU$(IDFUNCTION)
   EMAIL_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "EMAIL", "ID = " & IDFUNCTION, "NOMESS"))
End Function
Function EMAIL_INSTITU2$(IDFUNCTION)
   EMAIL_INSTITU2$ = TRIM$(VALOBADA("INSTITUCIONES", "[mail 2]", "ID = " & IDFUNCTION, "NOMESS"))
End Function
Function TELE_INSTITU$(IDFUNCTION)
   TELE_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "TELEPHONES", "ID = " & IDFUNCTION, "NOMESS"))
End Function

Function CIUDAD_INSTITU$(IDFUNCTION)
   CIUDAD_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "TOWN", "ID = " & IDFUNCTION, "NOMESS"))
End Function
    '
Function PROVINCIA_INSTITU$(IDFUNCTION)
   PROVINCIA_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "PROVINCE", "ID = " & IDFUNCTION, "NOMESS"))
End Function
    '
Function CPOST_INSTITU$(IDFUNCTION)
   CPOST_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "[Zip Code]", "ID = " & IDFUNCTION, "NOMESS"))
End Function
Function DIREC_INSTITU$(IDFUNCTION)
   DIREC_INSTITU$ = TRIM$(VALOBADA("INSTITUCIONES", "DOMICILIO", "ID = " & IDFUNCTION, "NOMESS"))
End Function

Function OBRA$(IDFUNCTION)
   OBRA$ = TRIM$(VALOBADA("FUNCIONES", "OBRA", "IDFUNCION = " & IDFUNCTION, "NOMESS"))
End Function
Function LUGAR$(IDFUNCTION)
   LUGAR$ = TRIM$(VALOBADA("FUNCIONES", "LUGAR", "IDFUNCION = " & IDFUNCTION, "NOMESS"))
End Function

Sub CARGACOMBOPROVINCIA(CMB As ComboBox)
    CMB.Clear
    SQLX$ = "SELECT NOMBRE, ID "
    SQLX$ = SQLX$ + " From PROVINCES WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY ID"
    Set RS = READSET(SQLX$)
    '
    With RS
      Do While Not .EOF
        CMB.AddItem SAFETEXT$(!NOMBRE)
        .MoveNext
      Loop
    End With
End Sub
Function PROXIMO_NRO_FUNCION&()
   PROXIMO_NRO_FUNCION& = 1 + XVALO(VALOBADA("FUNCIONES", "MAX(IDFUNCION)", "IDFUNCION > 0", "NOMESS"))
End Function
Function PROXIMO_NRO_RESERVA&()
   PROXIMO_NRO_RESERVA& = 1 + XVALO(VALOBADA("BOOKINGS", "MAX(IdReserva)", "IdReserva > 0", "NOMESS"))
End Function

Sub CARGALIST_FUNCIONES(MSF As MSFlexGrid, FECHADESDE, FECHAHASTA)
    'MUESTRA EN EL LIST LAS FUNCIONES QUE COINCIDEN CON EL DIA SELECCIONADO
    MSF.Rows = 1
    SQLX$ = "SELECT IDFUNCION,FECHA, OBRA, LUGAR"
    SQLX$ = SQLX$ + " From FUNCIONES WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE FECHA >= '" & FECHADESDE & "'"
    SQLX$ = SQLX$ + " AND FECHA <= '" & FECHAHASTA & "'"
    SQLX$ = SQLX$ + " ORDER BY FECHA"
    Set RS = READSET(SQLX$)
     '
     With RS
       Do While Not .EOF
         j& = j& + 1
         MSF.Rows = j& + 1
         MSF.TextMatrix(j&, 1) = SAFETEXT$(!FECHA)
         MSF.TextMatrix(j&, 2) = SAFETEXT$(!OBRA)
         MSF.TextMatrix(j&, 3) = RESERVAMASLIBERADA&(XVALO(!IDFUNCION))
         MSF.TextMatrix(j&, 4) = SAFETEXT$(!LUGAR)
'
'
'
'         Call REPLA(TX$, AJUSTI$(SAFETEXT$(!FECHA), 24), 1, 24)
'         Call REPLA(TX$, AJUSTI$(SAFETEXT$(!OBRA), 6), 26, 6)
'         Call REPLA(TX$, AJUSTD$(RESERVAMASLIBERADA&(XVALO(!IDFUNCION)), 6), 34, 6)
'         Call REPLA(TX$, AJUSTI$(SAFETEXT$(!LUGAR), 6), 42, 6)
'         LST.AddItem TX$
         .MoveNext
       Loop
     End With
    RS.Close
    Set RS = Nothing
    
End Sub
Function EXISTEFUNCION%(FECHAYHORA, ZONA$, OBRA$)
    'VALIDO SI EXISTE LA FUNCION , LO HAGO CONTRA EL LIST POR QUE TENIA PROBLEMAS PARA COMPARAR HORA Y FECHA CONTRA LA BASE
    With FUNCIONES1.MSF1
        For i& = 1 To .Rows - 1
           If UCase$(FECHAYHORA) = UCase$(TRIM$(.TextMatrix(i&, 1))) Then
              If TRIM$(UCase$(ZONA$)) = UCase$(TRIM$(.TextMatrix(i&, 4))) Then
                 If TRIM$(UCase$(OBRA$)) = UCase$(TRIM$(.TextMatrix(i&, 2))) Then
                    .Row = i&
                    .SetFocus
                    EXISTEFUNCION% = 1
                    Exit For
                 End If
              End If
           End If
        Next i&
    End With
End Function
Sub GENERA_TABLA_LUGARES()

      Call CREACAMPO("", "LUGARES", "Cod_lugar", 10, 16, 1)
      Call CREACAMPO("", "LUGARES", "REFERENCIA", 10, 16)
      Call CREACAMPO("", "LUGARES", "STATUS", 3, 0)
      
      'SI LA TABLA ESTA VACIA LE CARGO LA INFO LA PRIMERA VEZ
      If CantRegistros&("LUGARES", , "NOMESS") < 1 Then
        SQLX$ = " SELECT * FROM LUGARES "
        Dim RST As ADODB.Recordset
        Set RST = New ADODB.Recordset
        
        RST.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RST
            If Err = 0 Then
                SQLX$ = "SELECT substring(lugar,1,128) lugar "
                SQLX$ = SQLX$ + " From funciones WITH(NOLOCK)"
                SQLX$ = SQLX$ + " Group By substring(lugar,1,128)"
                SQLX$ = SQLX$ + " order by lugar "
                Set RS = READSET(SQLX$)
                '
                With RS
                  Do While Not .EOF
                    RST.AddNew
                    RST!COD_LUGAR = TRIM$(SAFETEXT$(!LUGAR))
                    Debug.Print !LUGAR
                    RST.Update
                    .MoveNext
                  Loop
                End With
            End If
        End With
      End If

End Sub
Sub AGREGA_CAMPO_STATUS_OBRAS()
   Call CREACAMPO("", "OBRAS", "STATUS", 3, 0)
   Call PoneEnCeroNull("OBRAS", "STATUS", REGAF&, "NOMESS")
   
End Sub
Sub iniciar_variables()
    
    pos% = InStrRev(App.Path, "\")
    LUDAT$ = Mid$(App.Path, 1, pos% - 1) + "\DATOS\"
    LUBAS$ = App.Path + "\"
    LUCOM$ = "C:\PROGRAMACION_INDEPENDIENTE\"
    
End Sub

Sub ABRECONEXION_PUNTODEVENTA()
   '
   On Error Resume Next
   ESTADODECONEXION% = FLEXCONN.State
   If ESTADODECONEXION% = 1 Then
      Exit Sub
   End If
   CONNSQL$ = TRIM$(LOADFILE$(LUDAT$ + "CONEXI.DRV"))
   If CONNSQL$ <> "" Then
      ' VERIFICA LA CONEXION
      Set FLEXCONN = New ADODB.Connection
      FLEXCONN.ConnectionString = CONNSQL$
      On Error Resume Next
      FLEXCONN.Open
      If Err Then
         MsgBox "Imposible Abrir Conexión a Servidor SQL con " + CONNSQL$, vbCritical
         Close: End
      End If
      On Error GoTo 0
      '
      FLEXCONN.Execute ("SET DATEFORMAT dmy")
      TIBASE$ = "SQL"
      TXT$ = CONNSQL$
      PINI% = InStr(UCase$(TXT$), "DATA SOURCE=") + 12
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      SERNA$ = Left$(TXT1$, PFIN% - 1)
      PINI% = InStr(UCase$(TXT$), "INITIAL CATALOG=") + 16
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      BINIC$ = Left$(TXT1$, PFIN% - 1)
      '
      BASEACTIVA$ = BINIC$
      TIBASE$ = "SQL"
      Exit Sub
   Else
      MsgBox "Archivo de Ruta de Conexión a Servidor Sql " & vbCrLf & "(CONNSQL.DRV) Inexistente o Vacio ", vbCritical
      Close: End
   End If
   '
   On Error Resume Next
   XX$ = CONN_SQLS.Label6.Caption
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "No se han Establecido los Parámetros\de Conexión al Servidor SQL-Server.")
      Call FINAL("")
   End If
   On Error GoTo 0
   CONN_SQLS.Show 1
   Call COMUNI("Re-Inicie Ahora la Aplicación.")
   Call FINAL("")
   '
End Sub


