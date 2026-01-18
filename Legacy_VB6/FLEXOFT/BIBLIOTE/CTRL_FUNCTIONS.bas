Attribute VB_Name = "CTRL_FUNCTIONS"
Public BaDaConsul As Database
Public BaDaMoviExt As Database
Public BaDaMensaFlex As Database
Public BaDaWald As Database
Public BASEWALD As ADODB.Connection
Public BASEWALDCLIENTE As ADODB.Connection
Public BASEINTERREMI_ENVAP As ADODB.Connection
Public BASEINTERREMI As ADODB.Connection
Public BASEINTERREMI_2_CONT As ADODB.Connection 'INTERREMI2 LA QUE SE CONFIGURA DE COTIZACO Y SE ESCRIBE EN REMITOS PARADO EN CONTEPLAST


Sub abreBaDaWald()
'   Static CTX%
'   If CTX% < 1 Then
'      CTX% = 1
         BDX$ = CONTROL("CONWALDB", "RUTAWALD")
         If BDX$ <> "" Then
          Set BASEWALD = New ADODB.Connection
          BASEWALD.Provider = "microsoft.jet.oledb.4.0"
          BASEWALD.ConnectionString = BDX$
          On Error Resume Next
          BASEWALD.Open
          If Err <> 0 Then
                Call COMUNI("Fallo al Intentar Abrir Base de Datos\" & BDX$)
          End If
          On Error GoTo 0
         End If
'     End If

End Sub
Sub abreBaMDBINTER_REMI() 'DONDE SE DUPLICAN LOS REGISTROS DE CONTEPLAST EN TABLAMDB
         BDX$ = TRIM$(CONTROL("CONWALDB", "ESCRMOVI"))
         If BDX$ <> "" Then
            Set BASEINTERREMI = New ADODB.Connection
            BASEINTERREMI.Provider = "microsoft.jet.oledb.4.0"
            BASEINTERREMI.ConnectionString = BDX$
            On Error Resume Next
            BASEINTERREMI.Open
            If Err > 0 Then
               Call COMUNI("Fallo al Intentar Abrir Base de Datos\" & BDX$)
            End If
            On Error GoTo 0
         End If

End Sub
Sub abreBaMDBINTER_REMI_2() 'DONDE SE DUPLICAN LOS REGISTROS DE CONTEPLAST EN TABLAMDB EN INTERREMI2
         BDX$ = TRIM$(CONTROL("CONWALDB", "ESCRMOV2")) ' ESTA SE CONFIGURA DESDE COTIZACO
         If BDX$ <> "" Then
            Set BASEINTERREMI_2_CONT = New ADODB.Connection
            BASEINTERREMI_2_CONT.Provider = "microsoft.jet.oledb.4.0"
            BASEINTERREMI_2_CONT.ConnectionString = BDX$
            On Error Resume Next
            BASEINTERREMI_2_CONT.Open
            If Err > 0 Then
               Call COMUNI("Fallo al Intentar Abrir Base de Datos\" & BDX$)
            End If
            On Error GoTo 0
         End If

End Sub
Sub abreBaDaWaldClientes()

         BDX$ = CONTROL("CONWALDB", "RTLECTCL")
         If BDX$ <> "" Then
          Set BASEWALDCLIENTE = New ADODB.Connection
          BASEWALDCLIENTE.Provider = "microsoft.jet.oledb.4.0"
          BASEWALDCLIENTE.ConnectionString = BDX$
          On Error Resume Next
          BASEWALDCLIENTE.Open
          If Err > 0 Then
                Call COMUNI("Fallo al Intentar Abrir Base de Datos\" & BDX$)
          End If
          On Error GoTo 0
         End If

End Sub
Sub abreBaMDBINTER_REMI_ENVAP() 'DONDE SE DUPLICAN LOS REGISTROS DE CONTEPLAST EN TABLAMDB
         BDX$ = TRIM$(CONTROL("TRCTPAWL", "DROPTRRM"))  'el dropbox de envaplast apuntando a la carpeta compartida:
                                                       'IMPOR_IMPLANSOFT\TRANSFREMITO\INTER_REMI2.mdb
         If BDX$ <> "" Then
          Set BASEINTERREMI_ENVAP = New ADODB.Connection
          BASEINTERREMI_ENVAP.Provider = "microsoft.jet.oledb.4.0"
          BASEINTERREMI_ENVAP.ConnectionString = BDX$
          On Error Resume Next
          BASEINTERREMI_ENVAP.Open
          If Err <> 0 Then
                Call COMUNI("Fallo al Intentar Abrir Base de Datos\" & BDX$)
          End If
          On Error GoTo 0
         End If

End Sub


Sub ABRETACON()
   ' crea la tabla de consultas en MDB
   Static CTX%
   If CTX% < 1 Then
      On Error GoTo ERRORSQLCTRLMDB
      CTX% = 1
      Call CREACAMPO_MDB("SQLCTRL", "SQL_QUERIES", "Cod_Query", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_QUERIES", "Des_Query", 10, 254)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_QUERIES", "Tip_Query", 3, 2)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_QUERIES", "Par_Query", 12, 0)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_QUERIES", "SQL_Query", 12, 0)
      '
      Call CREACAMPO_MDB("SQLCTRL", "SQL_GRAFICS", "Cod_Query", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_GRAFICS", "Cod_Graph", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_GRAFICS", "Des_Graph", 10, 254)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_GRAFICS", "Par_Graph", 12, 0)
      '
      Call CREACAMPO_MDB("SQLCTRL", "SQL_PROCEDS", "Cod_PROCEX", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_PROCEDS", "Des_PROCEX", 10, 254)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_PROCEDS", "Tip_PROCEX", 3, 2)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_PROCEDS", "Par_PROCEX", 12, 0)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_PROCEDS", "SQL_PROCEX", 12, 0)
      '
      Call CREACAMPO_MDB("SQLCTRL", "USER_PARMS", "Usuario", 3, 2, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_PARMS", "Proceso", 10, 128, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_PARMS", "Clave", 10, 128, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_PARMS", "ValoCon", 12, 0)
      '
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "Usuario", 3, 2, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "CodigoPadre", 10, 32, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "CodigoHijo", 10, 32, 1)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "Descripcion", 10, 64)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "CodigoImagen", 10, 32)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "NumeOrden", 3, 2)
      Call CREACAMPO_MDB("SQLCTRL", "USER_MENUS", "CantiAcce", 5, 4)
      '
      Call CREACAMPO_MDB("SQLCTRL", "SQL_MENUS", "Form_Name", 10, 32)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_MENUS", "Cod_Query", 10, 32)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_MENUS", "Num_Orden", 3, 2)
      '
      Call CREACAMPO_MDB("SQLCTRL", "SQL_ACCDIR", "Ejecutable", 10, 32)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_ACCDIR", "LEYENDA", 10, 128)
      Call CREACAMPO_MDB("SQLCTRL", "SQL_ACCDIR", "CODIGO", 10, 32)
      '
      Call CREACAMPO_MDB("SQLCTRL", "INDI_APLI", "COD_APLI", 10, 32)
      Call CREACAMPO_MDB("SQLCTRL", "INDI_APLI", "DES_APLI", 10, 128)
      '
      Call CREACAMPO_MDB("SQLCTRL", "DER_ACCE", "NUM_USER", 3, 2, 1)
      Call CREACAMPO_MDB("SQLCTRL", "DER_ACCE", "COD_APLI", 10, 32, 1)
      Call CREACAMPO_MDB("SQLCTRL", "DER_ACCE", "DER_READ", 3, 2)
      Call CREACAMPO_MDB("SQLCTRL", "DER_ACCE", "DER_WRITE", 3, 2)
      '
   End If
   '
   On Error Resume Next
   Set BaDaConsul = OpenDatabase(App.Path + "\SQLCTRL.MDB")
   On Error GoTo 0
   Exit Sub
   '
ERRORSQLCTRLMDB:
'   Call COMUNI("Error De Apertura de Base de Datos SQCTRL\Informe esta Situación a su Soporte de Sistemas")
   
End Sub
Sub ABREMENSAFLEX()
   ' crea la tabla de consultas en MDB
   Static CTX%
   If CTX% < 1 Then
      CTX% = 1
      Call CREACAMPO_MDB("MENSAFLEX", "MENSAJES", "NumMensa", 4, 0, 1)
      Call CREACAMPO_MDB("MENSAFLEX", "MENSAJES", "MensaViejo", 12, 0)
      Call CREACAMPO_MDB("MENSAFLEX", "MENSAJES", "MensaEsp", 12, 0)
      Call CREACAMPO_MDB("MENSAFLEX", "MENSAJES", "MensaIng", 12, 0)
      Call CREACAMPO_MDB("MENSAFLEX", "MENSAJES", "MensaPor", 12, 0)
      '
   End If
   '
   On Error Resume Next
   Set BaDaMensaFlex = OpenDatabase(App.Path + "\MENSAFLEX.MDB")
   On Error GoTo 0
   '
End Sub


Sub CREACAMPO_MDB(BASEDA$, TABLAX$, NCAMPO$, TCAMPO, LCAMPO, Optional CONINDEX%)
   '
   Static wrkPredeterminado As Workspace
   Static BASACTI$, CANCAM%, VECTAB$(), VECCAM$(), VECTIP%(), VECLON%(), VECIDX%()
   '
   ' ESTO SE AGREGO PARA NO CREAR LOS CAMPOS YA CREADOS
   If BASACTI$ <> UCase$(BASEDA$) Then
      Set wrkPredeterminado = DBEngine.Workspaces(0)
      Dim BASEDAX As Database
      On Error Resume Next
      Set BASEDAX = OpenDatabase(App.Path + "\" + BASEDA$ + ".MDB")
      If Err Then GoTo 19
      On Error GoTo 0
      '
      CANCAM% = 0: NUEDIM% = 64
      ReDim Preserve VECTAB$(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
      For i% = 1 To BASEDAX.TableDefs.Count
        For J% = 1 To BASEDAX.TableDefs(i% - 1).Fields.Count
          CANCAM% = CANCAM% + 1
          If CANCAM% >= UBound(VECTAB$) Then
             NUEDIM% = UBound(VECTAB$) + 64
             ReDim Preserve VECTAB$(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
          End If
          VECTAB$(CANCAM%) = BASEDAX.TableDefs(i% - 1).Name
          VECCAM$(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Name
          VECTIP(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Type
          VECLON(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Size
          '
          VECIDX%(CANCAM%) = 0
          NOIN$ = "I" + VECCAM$(CANCAM%)
          For k% = 1 To BASEDAX.TableDefs(i% - 1).Indexes.Count
            If BASEDAX.TableDefs(i% - 1).Indexes(k% - 1).Name = NOIN$ Then
              VECIDX%(CANCAM%) = 1
            End If
          Next k%
          '
        Next J%
      Next i%
      '
      BASACTI$ = UCase$(BASEDA$)
      BASEDAX.Close
   End If
   '
19 For KKI% = 1 To CANCAM%
     If VECTAB$(KKI%) = TABLAX$ Then
       If VECCAM$(KKI%) = NCAMPO$ Then
         If VECTIP(KKI%) = TCAMPO Then
           If VECLON(KKI%) >= LCAMPO Then
             If VECIDX%(KKI%) = CONINDEX% Then
               ' EL CAMPO YA EXISTE
               Exit Sub
             End If
           End If
         End If
       End If
     End If
   Next KKI%
   '
   '
' AQUI CAE SI EL CAMPO NO EXISTE O HAY QUE MODIFICAR ALGUN DATO
   '
   Dim BASDAT As Database
   Dim CREABA As Database
   Dim CREATA As TableDef
   '
   ' crea la base de datos BASEDA$
   Set wrkPredeterminado = DBEngine.Workspaces(0)
   On Error Resume Next
   Set CREABA = wrkPredeterminado.CreateDatabase(App.Path + "\" + BASEDA$ + ".MDB", dbLangGeneral)
   If Err < 1 Or Err = 3204 Then GoTo 10 ' LA BASE DE DATOS YA EXISTE
       MsgBox "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Base de Datos " + App.Path + "\" + BASEDA$
       Close: End
       '
10 ' crea la tabla TABLAX$
   On Error Resume Next
   Set BASDAT = OpenDatabase(App.Path + "\" + BASEDA$ + ".MDB")
   Set CREATA = BASDAT.CreateTableDef(TABLAX$)
   CREATA.Fields.Append CREATA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
   BASDAT.TableDefs.Append CREATA
   If Err < 1 Or Err = 3010 Then GoTo 20    ' LA tabla ya existe
       TTXX$ = "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Tabla '" + TABLAX$ + "' en BD '" + BASEDA$ + "'."
       On Error GoTo 0
       TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Es Probable que la Base de Datos de Control Necesite ser Compactada"
       TTXX$ = TTXX$ + vbCrLf + "o Restaurada."
       If EXISTE%(App.Path + "\FSETUP04.EXE") > 0 Then
            TTXX$ = TTXX$ + vbCrLf + vbCrLf + "¿ Desea Ingresar al Panel de Control para la Reparación ?"
            If MsgBox(TTXX$, vbYesNo, "Reparacion Base de Datos de Control") = vbYes Then
               On Error Resume Next
               Shell "FSETUP04.EXE " + "#" + JOBID$ + "#", 1
               If Err Then
                  MsgBox "Error " & Err & " - Imposible Conectar con " & LUBAS$ + ACONE$ + " " + CONECTA$ + "#" + JOBID$ + "#"
                  On Error GoTo 0
                  Close: End
                  Exit Sub
               End If
               Close: End
            End If
          Else
            TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Comuniquese con su Soporte de Sistemas FLEXOFT."
            MsgBox TTXX$
            Close: End
       End If
       Close: End
       '
20 ' Verifica pre-existencia del campo en la tabla
   Set CREACA = BASDAT.TableDefs(TABLAX$)
   For ICAMPO& = 1 To CREACA.Fields.Count
     If TRIM$(UCase$(CREACA.Fields(ICAMPO& - 1).Name)) = TRIM$(UCase$(NCAMPO$)) Then
       GoTo 25
     End If
   Next ICAMPO&
   '
   ' Agrega el campo a la tabla
   With CREACA.Fields
     On Error Resume Next
     .Append CREACA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
     If Err Then
       If Err = 3191 Then     ' El Campo ya Existe
           GoTo 25
         ElseIf Err = 3211 Or Err = 3212 Or Err = 3261 Or Err = 3262 Then
           ERREX& = Err
           On Error GoTo 0
           USLOCK$ = TRIM$(Left$(LOADFILE$(App.Path + "\" + BASEDA$ + ".LDB"), 24))
           TTXX$ = "Imposible Crear Campo '" + NCAMPO$ + "'\en Tabla '" + TABLAX$ + "'\en Base de Datos '" + BASEDA$ + ".MDB'."
           TTXX$ = TTXX$ + "\   \Error " + TRIM$(Str$(ERREX&)) + "."
           TTXX$ = TTXX$ + "\   \La Base de Datos está en Uso Exclusivo\en la Estación de Trabajo '" + USLOCK$ + "'."
           TTXX$ = TTXX$ + "\   \Cierre FLEXOFT en esa Terminal y Re-Intente."
           HUBOERROR% = 1
           TTYY$ = REPLACAR$(TTXX$, "\", vbCrLf)
           MsgBox TTYY$, vbExclamation
           RUTIACT$ = RUTIANT$
           Exit Sub
         Else
           MsgBox "Error " & Err & " al Crear Campo '" + NCAMPO$ + "' en Tabla '" + TABLAX$ + "' en BD '" + BASEDA$ + "'."
           RUTIACT$ = RUTIANT$
           Exit Sub
       End If
     End If
     On Error GoTo 0
     '
25   On Error Resume Next
     TICAM% = TCAMPO
     If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
           CREACA.Fields(NCAMPO$).AllowZeroLength = True
        ElseIf TICAM% = 1 Then
           CREACA.Fields(NCAMPO$).DefaultValue = True
        ElseIf TICAM% = 8 Then    ' FECHA
           CREACA.Fields(NCAMPO$).DefaultValue = CDate(0)
        Else
           CREACA.Fields(NCAMPO$).DefaultValue = 0
     End If
     On Error GoTo 0
   End With
   '
   If CONINDEX% = 1 Then
      On Error Resume Next
      BASDAT.Execute "CREATE INDEX I" + NCAMPO$ + " ON " + TABLAX$ + " ([" + NCAMPO$ + "]);"
      On Error GoTo 0
   End If
   '
End Sub
Sub CREACAMPO_MDB2(BASEDA$, TABLAX$, NCAMPO$, TCAMPO, LCAMPO, Optional CONINDEX%, Optional RUTA1$)
   '
   Static wrkPredeterminado As Workspace
   Static BASACTI$, CANCAM%, VECTAB$(), VECCAM$(), VECTIP%(), VECLON%(), VECIDX%()
   '
   'EN LA RUTA QUE SE LE PASA COMO ARCHIVO QUE DEBE SER SOLO LA UNIDAD EJ.: c:\
   RUTAX$ = TRIM$(RUTA1$ + "BDACCESS ") 'LE AGREGO LA CARPETA BDACCES SI NO EXISTE LA CREA.
   On Error Resume Next
   GetAttr (RUTAX$)
   If Err Then
       On Error GoTo 0
       MkDir RUTAX$
   End If
   On Error GoTo 0
   RUTAX$ = RUTAX$ + "\" ' LE AGREGO LA BARRA QUEDANDDO: c:\BDACCESS\
   ' ESTO SE AGREGO PARA NO CREAR LOS CAMPOS YA CREADOS
   If BASACTI$ <> UCase$(BASEDA$) Then
      Set wrkPredeterminado = DBEngine.Workspaces(0)
      Dim BASEDAX As Database
      On Error Resume Next
      Set BASEDAX = OpenDatabase(RUTAX$ + BASEDA$ + ".MDB")
      If Err Then GoTo 19
      On Error GoTo 0
      '
      CANCAM% = 0: NUEDIM% = 64
      ReDim Preserve VECTAB$(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
      For i% = 1 To BASEDAX.TableDefs.Count
        For J% = 1 To BASEDAX.TableDefs(i% - 1).Fields.Count
          CANCAM% = CANCAM% + 1
          If CANCAM% >= UBound(VECTAB$) Then
             NUEDIM% = UBound(VECTAB$) + 64
             ReDim Preserve VECTAB$(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
          End If
          VECTAB$(CANCAM%) = BASEDAX.TableDefs(i% - 1).Name
          VECCAM$(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Name
          VECTIP(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Type
          VECLON(CANCAM%) = BASEDAX.TableDefs(i% - 1).Fields(J% - 1).Size
          '
          VECIDX%(CANCAM%) = 0
          NOIN$ = "I" + VECCAM$(CANCAM%)
          For k% = 1 To BASEDAX.TableDefs(i% - 1).Indexes.Count
            If BASEDAX.TableDefs(i% - 1).Indexes(k% - 1).Name = NOIN$ Then
              VECIDX%(CANCAM%) = 1
            End If
          Next k%
          '
        Next J%
      Next i%
      '
      BASACTI$ = UCase$(BASEDA$)
      BASEDAX.Close
   End If
   '
19 For KKI% = 1 To CANCAM%
     If VECTAB$(KKI%) = TABLAX$ Then
       If VECCAM$(KKI%) = NCAMPO$ Then
         If VECTIP(KKI%) = TCAMPO Then
           If VECLON(KKI%) >= LCAMPO Then
             If VECIDX%(KKI%) = CONINDEX% Then
               ' EL CAMPO YA EXISTE
               Exit Sub
             End If
           End If
         End If
       End If
     End If
   Next KKI%
   '
   '
' AQUI CAE SI EL CAMPO NO EXISTE O HAY QUE MODIFICAR ALGUN DATO
   '
   Dim BASDAT As Database
   Dim CREABA As Database
   Dim CREATA As TableDef
   '
   ' crea la base de datos BASEDA$
   Set wrkPredeterminado = DBEngine.Workspaces(0)
   On Error Resume Next
   Set CREABA = wrkPredeterminado.CreateDatabase(RUTAX$ + BASEDA$ + ".MDB", dbLangGeneral)
   If Err < 1 Or Err = 3204 Then GoTo 10 ' LA BASE DE DATOS YA EXISTE
       MsgBox "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Base de Datos " + RUTAX$ + BASEDA$
       Close: End
       '
10 ' crea la tabla TABLAX$
   On Error Resume Next
   Set BASDAT = OpenDatabase(RUTAX$ + BASEDA$ + ".MDB")
   Set CREATA = BASDAT.CreateTableDef(TABLAX$)
   CREATA.Fields.Append CREATA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
   BASDAT.TableDefs.Append CREATA
   If Err < 1 Or Err = 3010 Then GoTo 20    ' LA tabla ya existe
       TTXX$ = "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Tabla '" + TABLAX$ + "' en BD '" + BASEDA$ + "'."
       On Error GoTo 0
       TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Es Probable que la Base de Datos de Control Necesite ser Compactada"
       TTXX$ = TTXX$ + vbCrLf + "o Restaurada."
       If EXISTE%(RUTAX$ + "FSETUP04.EXE") > 0 Then
            TTXX$ = TTXX$ + vbCrLf + vbCrLf + "¿ Desea Ingresar al Panel de Control para la Reparación ?"
            If MsgBox(TTXX$, vbYesNo, "Reparacion Base de Datos de Control") = vbYes Then
               On Error Resume Next
               Shell "FSETUP04.EXE " + "#" + JOBID$ + "#", 1
               If Err Then
                  MsgBox "Error " & Err & " - Imposible Conectar con " & LUBAS$ + ACONE$ + " " + CONECTA$ + "#" + JOBID$ + "#"
                  On Error GoTo 0
                  Close: End
                  Exit Sub
               End If
               Close: End
            End If
          Else
            TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Comuniquese con su Soporte de Sistemas FLEXOFT."
            MsgBox TTXX$
            Close: End
       End If
       Close: End
       '
20 ' Verifica pre-existencia del campo en la tabla
   Set CREACA = BASDAT.TableDefs(TABLAX$)
   For ICAMPO& = 1 To CREACA.Fields.Count
     If TRIM$(UCase$(CREACA.Fields(ICAMPO& - 1).Name)) = TRIM$(UCase$(NCAMPO$)) Then
       GoTo 25
     End If
   Next ICAMPO&
   '
   ' Agrega el campo a la tabla
   With CREACA.Fields
     On Error Resume Next
     .Append CREACA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
     If Err Then
       If Err = 3191 Then     ' El Campo ya Existe
           GoTo 25
         ElseIf Err = 3211 Or Err = 3212 Or Err = 3261 Or Err = 3262 Then
           ERREX& = Err
           On Error GoTo 0
           USLOCK$ = TRIM$(Left$(LOADFILE$(RUTAX$ + BASEDA$ + ".LDB"), 24))
           TTXX$ = "Imposible Crear Campo '" + NCAMPO$ + "'\en Tabla '" + TABLAX$ + "'\en Base de Datos '" + BASEDA$ + ".MDB'."
           TTXX$ = TTXX$ + "\   \Error " + TRIM$(Str$(ERREX&)) + "."
           TTXX$ = TTXX$ + "\   \La Base de Datos está en Uso Exclusivo\en la Estación de Trabajo '" + USLOCK$ + "'."
           TTXX$ = TTXX$ + "\   \Cierre FLEXOFT en esa Terminal y Re-Intente."
           HUBOERROR% = 1
           TTYY$ = REPLACAR$(TTXX$, "\", vbCrLf)
           MsgBox TTYY$, vbExclamation
           RUTIACT$ = RUTIANT$
           Exit Sub
         Else
           MsgBox "Error " & Err & " al Crear Campo '" + NCAMPO$ + "' en Tabla '" + TABLAX$ + "' en BD '" + BASEDA$ + "'."
           RUTIACT$ = RUTIANT$
           Exit Sub
       End If
     End If
     On Error GoTo 0
     '
25   On Error Resume Next
     TICAM% = TCAMPO
     If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
           CREACA.Fields(NCAMPO$).AllowZeroLength = True
        ElseIf TICAM% = 1 Then
           CREACA.Fields(NCAMPO$).DefaultValue = True
        ElseIf TICAM% = 8 Then    ' FECHA
           CREACA.Fields(NCAMPO$).DefaultValue = CDate(0)
        Else
           CREACA.Fields(NCAMPO$).DefaultValue = 0
     End If
     On Error GoTo 0
   End With
   '
   If CONINDEX% = 1 Then
      On Error Resume Next
      BASDAT.Execute "CREATE INDEX I" + NCAMPO$ + " ON " + TABLAX$ + " ([" + NCAMPO$ + "]);"
      On Error GoTo 0
   End If
   '
End Sub

'
Function VALOTACO(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   'PASANDOLE COMO ARGUMENTO EL NOMBRE DEL CAMPO , EL DE LA TABLA Y LA CONDICION _
    DEVUELVE EN VALOBADA EL VALOR DE LA ENCONTRADO SEGUN LOS ARGUMENTOS ENVIADOS
   
   Call ABRETACON
   SQLX$ = "SELECT " + NOCAM$ + " AS VALORSE "
   SQLX$ = SQLX$ + " FROM " + NOTAB$
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   '
   VALOTACO = ""
   On Error Resume Next
   Set RS1 = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
   RS1.MoveFirst
   If Err Then
      On Error GoTo 0
      If MODOMESS$ <> "NOMESS" Then
         MsgBox "Imposible Ejecutar Consulta '" + SQLX$ + "'"
      End If
      GoTo 90
   End If
   On Error GoTo 0
   With RS1
     If Not (.BOF And .EOF) Then
       VTP% = VarType(!VALORSE)
       If VTP% = 8 Or VTP% < 2 Then
           VALOTACO = !VALORSE
         Else
           VALOTACO = !VALORSE
       End If
     End If
   End With
   '
90 On Error Resume Next
   RS1.Close: Set RS1 = Nothing
   BaDaConsul.Close
   On Error GoTo 0
   '
End Function

Function VALOMENSA(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   'PASANDOLE COMO ARGUMENTO EL NOMBRE DEL CAMPO , EL DE LA TABLA Y LA CONDICION _
    DEVUELVE EN VALOBADA EL VALOR DE LA ENCONTRADO SEGUN LOS ARGUMENTOS ENVIADOS
   
   Call ABREMENSAFLEX
   SQLX$ = "SELECT " + NOCAM$ + " AS VALORSE "
   SQLX$ = SQLX$ + " FROM " + NOTAB$
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   '
   VALOMENSA = ""
   On Error Resume Next
   Set RS1 = BaDaMensaFlex.OpenRecordset(SQLX$, dbOpenSnapshot)
   RS1.MoveFirst
   If Err Then
      On Error GoTo 0
      If MODOMESS$ <> "NOMESS" Then
         MsgBox "Imposible Ejecutar Consulta '" + SQLX$ + "'"
      End If
      GoTo 90
   End If
   On Error GoTo 0
   With RS1
     If Not (.BOF And .EOF) Then
       VTP% = VarType(!VALORSE)
       If VTP% = 8 Or VTP% < 2 Then
           VALOMENSA = SAFETEXT$(!VALORSE)
         Else
           VALOMENSA = !VALORSE
       End If
     End If
   End With
   '
90 On Error Resume Next
   RS1.Close: Set RS1 = Nothing
   BaDaMensaFlex.Close
   On Error GoTo 0
   '
End Function

Sub COMPACTABD_VIE()
     '
     Screen.MousePointer = 11
     NOMECOR$ = "SQLCTRL"
     '
     DIREPRO$ = App.Path
     DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
     On Error Resume Next
        MkDir DIRENEW$
     On Error GoTo 0
     '
     DIREVIE$ = DIREPRO$ + "\VIEJOS\"
     ' CREA LA CARPETA \VIEJOS
     On Error Resume Next
     MkDir Left$(DIREVIE$, Len(DIREVIE$) - 1) ' PARA SUPRIMIR LA ULTIMA ANTIBARRA
     On Error GoTo 0
     '
     On Error Resume Next
     FIDTBK = CDbl(FileDateTime(DIRENEW$ + "\" + NOMECOR$ + ".MDB"))
     On Error GoTo 0
     If FIDTBK > CDbl(Now) - 1 Then GoTo 99
     '
     On Error Resume Next
     BaDaConsul.Close
        If Err Then Call MUERRO("1) BADACONSUL.CLOSE - ERROR " & Err.Number)
     On Error GoTo 0
     '
     On Error Resume Next
     BASDAT.Close
        If Err Then Call MUERRO("1) BASDAT.CLOSE - ERROR " & Err.Number)
     On Error GoTo 0
     '
     On Error Resume Next
     Kill DIRENEW$ + "\" + NOMECOR$ + ".MDB"
        If Err Then Call MUERRO("2) Kill " & DIRENEW$ + "\" + NOMECOR$ + ".MDB - ERROR " & Err.Number)
     On Error GoTo 0
     '
60   On Error Resume Next
     CompactDatabase App.Path + "\" + NOMECOR$ + ".MDB", DIRENEW$ + "\" + NOMECOR$ + ".MDB"
        If Err Then
           Call MUERRO("3) CompactDatabase " & App.Path + "\" + NOMECOR$ + ".MDB - ERROR " & Err.Number)
           GoTo 99
        End If
     On Error GoTo 0
     '
     ' AQUI PUDO COMPACTAR
     ANAMCO$ = NOMECOR$ + ".MD"
     On Error Resume Next
     Kill DIREVIE$ + ANAMCO$ + "9"
     On Error GoTo 0
     If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
         MsgBox "Imposible Instalar. No Pudo Borrarse " + DIREVIE$ + ANAMCO$ + "9"
         On Error GoTo 0
         GoTo 99
     End If
     '
     On Error Resume Next
     For KI% = 8 To 0 Step -1
         Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
     Next KI%
     On Error GoTo 0
     If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
       MsgBox "Imposible Instalar. No Pudo Renombrarse " + DIREVIE$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "1"
       On Error GoTo 0
       GoTo 99
     End If
     '
     On Error Resume Next
     FileCopy DIRENEW$ + "\" + ANAMCO$ + "B", DIREVIE$ + ANAMCO$ + "0"
     If Err Then
         MsgBox "Imposible Instalar. No Pudo Copiarse " + DIRENEW$ + "\" + ANAMCO$ + "B como " + DIREVIE$ + ANAMCO$ + "0"
         On Error GoTo 0
         GoTo 99
     End If
     '
     On Error Resume Next
     FileCopy DIRENEW$ + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + ".MDB"
     On Error GoTo 0
     '
99   On Error GoTo 0
     Screen.MousePointer = 1
     '
End Sub
'
Sub COMPACTABD()
     '
     Screen.MousePointer = 11
     NOMECOR$ = "SQLCTRL"
     '
     DIREPRO$ = App.Path
' DIREPRO$ = "X:\BLIPACK\FLEXOFT\WINDOWS"
     DIREVIE$ = DIREPRO$ + "\VIEJOS\"
     '
     ' CREA LA CARPETA \VIEJOS
     On Error Resume Next
     MkDir Left$(DIREVIE$, Len(DIREVIE$) - 1) ' PARA SUPRIMIR LA ULTIMA ANTIBARRA
     On Error GoTo 0
     '
     On Error Resume Next
'     FIDTKMP = CDbl(FileDateTime(DIREPRO$ + "\" + NOMECOR$ + ".KMP"))
     FIDTBKP = CDbl(FileDateTime(DIREVIE$ + NOMECOR$ + ".MD0"))
     On Error GoTo 0
     If FIDTBKP > CDbl(Now) - 1 Then GoTo 80       ' NO RESPALDA LA BASE ACTUAL
     '
     ' CIERRA LA BASE DE DATOS
     On Error Resume Next
     BaDaConsul.Close
        If Err Then Call MUERRO("1) BADACONSUL.CLOSE - ERROR " & Err.Number)
     On Error GoTo 0
     '
     On Error Resume Next
     BASDAT.Close
       If Err Then
         If Err.Number <> 424 Then           ' SI ES 424 ES QUE NO EXISTE BASDAT - ESTÁ TODO BIEN
           Call MUERRO("1) BASDAT.CLOSE - ERROR " & Err.Number)
         End If
       End If
     On Error GoTo 0
     '
     ' COPIA LA ACTUAL COMO .MD0
     On Error Resume Next
     FileCopy DIREPRO$ + "\" + NOMECOR$ + ".MDB", DIREPRO$ + "\" + NOMECOR$ + ".MD0"
     If Err Then
        Call MUERRO("Imposible Instalar. No Pudo Copiarse " + DIREPRO$ + "\" + NOMECOR$ + ".MDB como " + DIREPRO$ + "\" + NOMECOR$ + ".MD0")
        ' SI ES ERROR 70 EN USO POR EL WEB_BROWSR
        On Error GoTo 0
        GoTo 99
     End If
     '
     If EXISTE%(DIREPRO$ + "\" + NOMECOR$ + ".MD0") < 1 Then
        Call MUERRO("Imposible Instalar. No Existe " + DIREPRO$ + "\" + NOMECOR$ + ".MD0")
        On Error GoTo 0
        GoTo 99
     End If
     '
     ' CORRE LOS RESPALDOS
     ANAMCO$ = NOMECOR$ + ".MD"
     On Error Resume Next
     Kill DIREVIE$ + ANAMCO$ + "9"
     On Error GoTo 0
     If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
         Call MUERRO("Imposible Instalar. No Pudo Borrarse " + DIREVIE$ + ANAMCO$ + "9")
         On Error GoTo 0
         GoTo 99
     End If
     '
     On Error Resume Next
     For KI% = 8 To 0 Step -1
       If EXISTE%(DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%))) > 0 Then
         Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
       End If
     Next KI%
     On Error GoTo 0
     If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
       Call MUERRO("Imposible Instalar. No Pudo Renombrarse " + DIREVIE$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "1")
       On Error GoTo 0
       GoTo 99
     End If
     '
     ' COPIA MD0 A DIRECTORIO VIEJO
     On Error Resume Next
     FileCopy DIREPRO$ + "\" + ANAMCO$ + "0", DIREVIE$ + ANAMCO$ + "0"
     If Err Then
         Call MUERRO("Imposible Instalar. No Pudo Copiarse " + DIRENEW$ + "\" + ANAMCO$ + "B como " + DIREVIE$ + ANAMCO$ + "0")
         On Error GoTo 0
         GoTo 99
     End If
     '
' *********************************************************************************
     
80   On Error Resume Next
     FILENKMP = FileLen(DIREPRO$ + "\" + NOMECOR$ + ".KMP")
     FILENAPP = FileLen(DIREPRO$ + "\" + NOMECOR$ + ".MDB")
     On Error GoTo 0
     If FILENAPP < 1.25 * FILENKMP Then GoTo 99   ' NO HACE FALTA COMPACTAR
     '
     ' RECOMIENDA COMPACTAR LA BASE DE DATOS DE CONTROL
     TTXX$ = "La Base de Datos de Control Requiere Mantenimiento."
     TTXX$ = TTXX$ + "\Se Recomienda Ejecutar el Proceso de Compactación."
     TTXX$ = TTXX + "\   \Oportunamente Informe  al Supervisor del Sistema."
     TTYY$ = REPLACAR$(TTXX$, "\", vbCrLf)
     'SE COMENTO PARA QUE NO APAAAREZCA PERMANENTE MENTE
'     MsgBox TTYY$, vbExclamation
     '
99   On Error GoTo 0
     Screen.MousePointer = 1
     '
End Sub
'
Sub MUERRO(TTTT$)
    '
    Static MUESTRAERR%
    '
    If MUESTRAERR% = 0 Then
       MUESTRAERR% = (-1)
       UL$ = "C:"   'Left$(App.Path, 2)
'       If SERIVOL$(UL$) = "445E-73A2" Then MUESTRAERR% = 1  ' MAQUINA ERNESTO
    End If
    '
    If MUESTRAERR% = 1 Then
       MsgBox TTTT$
    End If
    '
End Sub
'
