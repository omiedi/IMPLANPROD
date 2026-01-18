Attribute VB_Name = "ABREBASE"
Public Contable As Database
Public Tabejcon As Recordset
Public Plancuen As Recordset
Public Encasien As Recordset
Public Detasien As Recordset
Public Balasec As Recordset
'
Public ACTICARGA%
'
Public Articulos As Database
Public Master As Recordset
Public Formulas As Recordset
Public Secoper As Recordset
'
Public Badaclie As Database
Public Deudor12 As Recordset
'
Public Badaprov As Database
Public Proved12 As Recordset
'
Public CueCorri As Database
Public CajaBanc As Recordset
Public SacrePen As Recordset
Public SaDeuPen As Recordset
Public ApliPago As Recordset
Public ApliCobra As Recordset
Public EncMovCc As Recordset 'nueva tabla p/imporclieCtaCte
Public DetMovCc As Recordset 'nueva tabla p/imporclieCtaCte
'
Public Stocks As Database
Public Movisto As Recordset
Public SalStock As Recordset
Public Reserva As Recordset
Public Bolrecep As Recordset
Public MaConsig As Recordset
'
Public Ofabric As Database
Public OrdeFabr As Recordset
Public OperFab As Recordset
Public InfoFab As Recordset  'nueva tabla

Public Compras As Database
Public OcomEnca As Recordset
Public OcomDeta As Recordset
'
Public Ventas As Database
Public PresuVen As Recordset
Public PedEnca As Recordset
Public PedDeta As Recordset
'
Public Controcen As Database
Public BloqArchA As Recordset
Public BloqArchB As Recordset
'
Public STOCKORIGI As Database
Public EncaStock As Recordset
Public DetaStock As Recordset
'
Public OTrabajo As Database
Public OrdeTrab As Recordset
Public PlanCob As Recordset
Public GrupOtra As Recordset
Public ListaMat As Recordset
'
Public StockVie As Database   'MOVIMIENTOS STOCK ANTIGUO.
Public Movisto1 As Recordset  '1 XQ MOVISTO RECORDSET YA EXISTE.
'
Public PresuCam As Database
Public EncaPresu As Recordset
Public DetaPresu As Recordset
'OT-05-0856-WAR-28/11/05///
Public Maquinarias As Database
Public Equipos As Recordset
Public GaleImag As Recordset
Public RepuMaq As Recordset
Public Contadores As Recordset
Public SubConjuntos As Recordset
'
Public Omantin As Database
Public OrdeMant As Recordset
Public RecuProp As Recordset
'OT-05-0856-WAR-FIN///
Sub ABRETACEN()
   '
   Static YABIER%, LUACCA$
   Screen.MousePointer = 11
   '
   If LUACC$ <> LUACCA$ Then
     YABIER% = 0
     LUACCA$ = LUACC$
   End If
   '
   If YABIER% > 0 Then GoTo 90
   '
   'If EXISTE%(LUACC$ + "BCONTROL.MDB") < 1 Then ' crea la base de datos
   'On Error Resume Next
   Dim wrkPredeterminado As Workspace
   Set wrkPredeterminado = DBEngine.Workspaces(0)
     '
   On Error Resume Next
   ' crea la base de datos
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + "BCONTROL.MDB", dbLangGeneral)
   Set Controcen = OpenDatabase(LUACC$ + "BCONTROL.MDB")
   '
   ' crea la tabla "BloArchA"
   Set CREATABLA = Controcen.CreateTableDef("BloArchA")
   CREATABLA.Fields.Append CREATABLA.CreateField("IdenTabla", dbText, 32)
   Controcen.TableDefs.Append CREATABLA
   '
   ' crea la tabla "BloArchB"
   Set CREATABLA = Controcen.CreateTableDef("BloArchB")
   CREATABLA.Fields.Append CREATABLA.CreateField("IdenTabla", dbText, 32)
   Controcen.TableDefs.Append CREATABLA
   '
   ' agrega campos a "BloArchB"
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("UserName", dbText, 32)
     .Append CREATABLA.CreateField("TermName", dbText, 32)
     .Append CREATABLA.CreateField("ExecName", dbText, 32)
     .Append CREATABLA.CreateField("CtrlName", dbText, 32)
     .Append CREATABLA.CreateField("InitDate", dbDate)
   End With
   '
90 Set BloqArchA = Controcen.OpenRecordset("BloArchA", dbOpenDynaset)
   Set BloqArchB = Controcen.OpenRecordset("BloArchB", dbOpenDynaset)
   On Error GoTo 0
   YABIER% = 1
   Screen.MousePointer = 1
   '
End Sub

Sub ABREPLANCUEN()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("CONTAGEN")
     Set Contable = OpenDatabase(LUACC$ + "CONTAGEN.MDB")
   End If
   '
' Crea el índice ÍndNue en la tabla Empleados.
'On Error Resume Next
'Contable.Execute "CREATE INDEX ICueCoInt ON DetaAsien ([CueCoInt]);"
'Contable.Execute "CREATE INDEX IFechasien ON DetaAsien ([Fechasien]);"
'On Error GoTo 0

   On Error Resume Next
   Tabejcon.MoveFirst
     If Err Then GoTo 90
   Plancuen.MoveFirst
     If Err Then GoTo 90
   Encasien.MoveFirst
     If Err Then GoTo 90
   Detasien.MoveFirst
     If Err Then GoTo 90
   GoTo 99
   '
90 Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenDynaset, dbReadOnly)
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset, dbReadOnly)
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset, dbReadOnly)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset, dbReadOnly)
   '
99 YABIER% = 1
   Screen.MousePointer = SCMP
   End Sub

Sub ABREMASTER()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("ARTICULOS")
     Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
     YABIER% = 1
   End If
   '
90 Set Master = Articulos.OpenRecordset("MASTER", dbOpenDynaset)
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRESTRUC()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("ARTICULOS")
     Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
     YABIER% = 1
   End If
   '
90 Set Formulas = Articulos.OpenRecordset("FORMULAS", dbOpenDynaset)
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRESECOP()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("ARTICULOS")
     Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
     YABIER% = 1
   End If
   '
90 Set Secoper = Articulos.OpenRecordset("SECOPER", dbOpenDynaset)
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub VABRESTRUC()
   Static YABIER%
   If YABIER% > 0 Then GoTo 90
   '
   Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
   YABIER% = 1
   '
   On Error Resume Next
   CAMPOS% = Articulos.TableDefs("Formulas").Fields.Count
   If CAMPOS% < 1 Then
     Set CREAMASTER = Articulos.CreateTableDef("Formulas")
     CREAMASTER.Fields.Append CREAMASTER.CreateField("CodiConj", dbInteger)
     Articulos.TableDefs.Append CREAMASTER
   End If
   On Error GoTo 0
   '
   Set CREACAMPOS = Articulos.TableDefs("Formulas")
   On Error Resume Next
   With CREACAMPOS.Fields
     .Append CREACAMPOS.CreateField("CodXConj", dbText)
     .Append CREACAMPOS.CreateField("CodiElem", dbInteger)
     .Append CREACAMPOS.CreateField("CodXElem", dbText)
     .Append CREACAMPOS.CreateField("UsoBruto", dbSingle)
     .Append CREACAMPOS.CreateField("Pladispo", dbInteger)
     .Append CREACAMPOS.CreateField("Marborra", dbInteger)
   End With
   On Error GoTo 0
   '
   CACAM% = Articulos.TableDefs("Formulas").Fields.Count
   For KU% = 1 To CACAM%
     If Articulos.TableDefs("Formulas").Fields(KU% - 1).Type = 10 Then
       Articulos.TableDefs("Formulas").Fields(KU% - 1).AllowZeroLength = True
     End If
   Next KU%
   '
90 Set Formulas = Articulos.OpenRecordset("FORMULAS", dbOpenDynaset)
   '
End Sub

Sub VABRESECOP()
   '
   Static YABIER%, LUACCA$
   Screen.MousePointer = 11
   '
   If LUACC$ <> LUACCA$ Then
     YABIER% = 0
     LUACCA$ = LUACC$
   End If
   '
   If YABIER% > 0 Then GoTo 90
   '
   Dim wrkPredeterminado As Workspace
   Set wrkPredeterminado = DBEngine.Workspaces(0)
     '
   On Error Resume Next
   ' crea la base de datos
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + "ARTICULOS.MDB", dbLangGeneral)
   Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
   '
   ' crea la tabla "SECOPER"
   Set CREATABLA = Articulos.CreateTableDef("SECOPER")
   CREATABLA.Fields.Append CREATABLA.CreateField("CodIConj", dbInteger)
   Articulos.TableDefs.Append CREATABLA
   '
   ' agrega campos a "SECOPER"
   Set CREACAMPOS = Articulos.TableDefs("SECOPER")
   With CREACAMPOS.Fields
     .Append CREACAMPOS.CreateField("CodXConj", dbText, 24)
     .Append CREACAMPOS.CreateField("NumeOper", dbInteger)
     .Append CREACAMPOS.CreateField("DescOper", dbText, 64)
     .Append CREACAMPOS.CreateField("StatOper", dbInteger)
     .Append CREACAMPOS.CreateField("CodMaq0", dbText, 6)
     .Append CREACAMPOS.CreateField("CodMaq1", dbText, 6)
     .Append CREACAMPOS.CreateField("CodMaq2", dbText, 6)
     .Append CREACAMPOS.CreateField("CodMaq3", dbText, 6)
     .Append CREACAMPOS.CreateField("CodMaq4", dbText, 6)
     .Append CREACAMPOS.CreateField("CodMaq5", dbText, 6)
     .Append CREACAMPOS.CreateField("HorsPrep", dbSingle)
     .Append CREACAMPOS.CreateField("HorsFijo", dbSingle)
     .Append CREACAMPOS.CreateField("MinuStan", dbSingle)
     .Append CREACAMPOS.CreateField("HorsLimp", dbSingle)
     .Append CREACAMPOS.CreateField("HorsMHer", dbSingle)
     .Append CREACAMPOS.CreateField("PieCiclo", dbSingle)
     .Append CREACAMPOS.CreateField("PorDemor", dbSingle)
     .Append CREACAMPOS.CreateField("Cadencia", dbSingle)
     .Append CREACAMPOS.CreateField("CoefSolp", dbSingle)
     .Append CREACAMPOS.CreateField("LoteStan", dbSingle)
     .Append CREACAMPOS.CreateField("CantiOps", dbSingle)
     .Append CREACAMPOS.CreateField("CategOps", dbInteger)
     .Append CREACAMPOS.CreateField("ProgCNum", dbText, 24)
     .Append CREACAMPOS.CreateField("HoProFab", dbText, 24)
      '
     .Append CREACAMPOS.CreateField("ObseOper", dbMemo)
      '
     .Append CREACAMPOS.CreateField("ProTerOp", dbInteger)
     .Append CREACAMPOS.CreateField("ProvOper", dbText, 12)
     .Append CREACAMPOS.CreateField("PreUnid", dbDouble)
     .Append CREACAMPOS.CreateField("MonePrec", dbInteger)
     .Append CREACAMPOS.CreateField("FechPrec", dbDate)
   End With
   '
   CACAM% = Articulos.TableDefs("SECOPER").Fields.Count
   For KU% = 1 To CACAM%
     TICAM = Articulos.TableDefs("SECOPER").Fields(KU% - 1).Type
     If TICAM = 10 Or TICAM = 12 Then
       Articulos.TableDefs("SECOPER").Fields(KU% - 1).AllowZeroLength = True
     End If
   Next KU%
   '
90 Set Secoper = Articulos.OpenRecordset("SECOPER", dbOpenDynaset)
   Screen.MousePointer = 1
   '
End Sub
'

Sub ABRECLIEN()
   '
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("BADACLIEN")
     Set Badaclie = OpenDatabase(LUACC$ + "BADACLIEN.MDB")
   End If
   '
90 Set Deudor12 = Badaclie.OpenRecordset("DEUDOR12", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABREPROVE()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("BADAPROVE")
     Set Badaprov = OpenDatabase(LUACC$ + "BADAPROVE.MDB")
   End If
   '
90 Set Proved12 = Badaprov.OpenRecordset("Proved12", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRECABAN()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("CUECORRI")
     Set CueCorri = OpenDatabase(LUACC$ + "CUECORRI.MDB")
   End If
   '
' Crea el índice ÍndNue en la tabla Empleados.
'On Error Resume Next
'CueCorri.Execute "CREATE INDEX ICueContaI ON cajabanc ([CueContaI]);"
'CueCorri.Execute "CREATE INDEX IFechConta ON cajabanc ([FechConta]);"
'CueCorri.Execute "CREATE INDEX IFechEmisi ON cajabanc ([FechEmisi]);"
'CueCorri.Execute "CREATE INDEX ITipoMovim ON cajabanc ([TipoMovim]);"
'CueCorri.Execute "CREATE INDEX INumeComp ON cajabanc ([NumeComp]);"
'On Error GoTo 0

   On Error Resume Next
   CajaBanc.MoveFirst
     If Err Then GoTo 90
   SacrePen.MoveFirst
     If Err Then GoTo 90
   SaDeuPen.MoveFirst
     If Err Then GoTo 90
   ApliPago.MoveFirst
     If Err Then GoTo 90
   ApliCobra.MoveFirst
     If Err Then GoTo 90
   GoTo 99
   '
90 Set CajaBanc = CueCorri.OpenRecordset("CAJABANC", dbOpenDynaset)
   Set SacrePen = CueCorri.OpenRecordset("SACREPEN", dbOpenDynaset)
   Set SaDeuPen = CueCorri.OpenRecordset("SADEUPEN", dbOpenDynaset)
   Set ApliPago = CueCorri.OpenRecordset("APLIPAGO", dbOpenDynaset)
   Set ApliCobra = CueCorri.OpenRecordset("APLICOBRA", dbOpenDynaset)
   'Set EncMovCc = CueCorri.OpenRecordset("ENCMOVCC", dbOpenDynaset) 'Add p/import.
   'Set DetMovCc = CueCorri.OpenRecordset("DETMOVCC", dbOpenDynaset) 'add p/import.
   '
99 YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRECOMPRAS()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("COMPRAS")
     Call AGRECAMPO("COMPRAS", "OCOMENCA", "OcomRevis", dbInteger, 2)
     Set Compras = OpenDatabase(LUACC$ + "COMPRAS.MDB")
   End If
   '
90 Set OcomEnca = Compras.OpenRecordset("OCOMENCA", dbOpenDynaset)
   Set OcomDeta = Compras.OpenRecordset("OCOMDETA", dbOpenDynaset)
   '
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABREPRESVEN()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("VENTAS")
     Set Ventas = OpenDatabase(LUACC$ + "VENTAS.MDB")
   End If
   '
90 Set PresuVen = Ventas.OpenRecordset("PRESUVEN", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABREPEDCLI()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("VENTAS")
     Set Ventas = OpenDatabase(LUACC$ + "VENTAS.MDB")
   End If
   '
90 Set PedEnca = Ventas.OpenRecordset("PEDENCA", dbOpenDynaset)
   Set PedDeta = Ventas.OpenRecordset("PEDDETA", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABRESTOCKS()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("STOCKS")
     Set Stocks = OpenDatabase(LUACC$ + "STOCKS.MDB")
     '
     On Error Resume Next
     Stocks.Execute "CREATE INDEX ICod_Inte ON Movisto ([Cod_Inte]);"
     Stocks.Execute "CREATE INDEX ICodiMovi ON Movisto ([CodiMovi]);"
     Stocks.Execute "CREATE INDEX IDoPriLar ON Movisto ([DoPriLar]);"
     Stocks.Execute "CREATE INDEX ICod_Inte ON Bolrecep ([Cod_Inte]);"
     Stocks.Execute "CREATE INDEX ICodiInte ON SalStock ([CodiInte]);"
     On Error GoTo 0
     '
   End If
   '
   On Error Resume Next
   Movisto.MoveFirst
     If Err Then GoTo 90
   SalStock.MoveFirst
     If Err Then GoTo 90
   Bolrecep.MoveFirst
     If Err Then GoTo 90
   MaConsig.MoveFirst
     If Err Then GoTo 90
   GoTo 99
   '
90 On Error GoTo 0
   Set Movisto = Stocks.OpenRecordset("MOVISTO", dbOpenDynaset)
   Set SalStock = Stocks.OpenRecordset("SALSTOCK", dbOpenDynaset)
   Set Bolrecep = Stocks.OpenRecordset("BOLRECEP", dbOpenDynaset)
   Set MaConsig = Stocks.OpenRecordset("MACONSIG", dbOpenDynaset)
   '
99 YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABRESTOCKVIE()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("STOCKVIE")
     Set StockVie = OpenDatabase(LUACC$ + "STOCKVIE.MDB")
   End If
   '
90 Set Movisto1 = StockVie.OpenRecordset("MOVISTO", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRESERVA()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("STOCKS")
     Set Stocks = OpenDatabase(LUACC$ + "STOCKS.MDB")
     On Error Resume Next
     Stocks.Execute "CREATE INDEX ICod_Inte ON Reserva ([Cod_Inte]);"
     On Error GoTo 0
   End If
   '
   On Error Resume Next
   Reserva.MoveFirst
     If Err Then GoTo 90
   GoTo 99
   '
90 On Error GoTo 0
   Set Reserva = Stocks.OpenRecordset("RESERVA", dbOpenDynaset)
   '
99 YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABREORFAB()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("OFABRIC")
     Set Ofabric = OpenDatabase(LUACC$ + "OFABRIC.MDB")
   End If
   '
90 Set OrdeFabr = Ofabric.OpenRecordset("ORDEFABR", dbOpenDynaset)
   Set OperFab = Ofabric.OpenRecordset("OPERFAB", dbOpenDynaset)
   Set InfoFab = Ofabric.OpenRecordset("INFOFAB", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub BLOQTABLA(BATAB$)
   '
   BASETABLA$ = TRIM$(UCase$(BATAB$))
   '
   Call ABRETACEN
10 With BloqArchA
     .FindFirst "IDENTABLA = '" & BASETABLA$ & "'"
     If .NoMatch = True Then
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !IDENTABLA = BASETABLA$
       .Update
       BloqArchA.Close
       GoTo 10
     End If
   End With
   '
20 With BloqArchB
     .FindFirst "IDENTABLA = '" & BASETABLA$ & "'"
     If .NoMatch = True Then
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !IDENTABLA = BASETABLA$
       .Update
       BloqArchB.Close
       GoTo 20
     End If
   End With
   '
30 On Error Resume Next
   With BloqArchA
     .Edit
     If Err Then
       TTXX$ = "En uso Exclusivo por " + TRIM$(BloqArchB!USERNAME)
       TTXX$ = TTXX$ + " en " + TRIM$(BloqArchB!TERMNAME)
       TTXX$ = TTXX$ + " desde " + Format$(BloqArchB!INITDATE, "mm/mm/yy  hh:mm:ss")
       MsgBox TTXX$
       Resume 30
       GoTo 30
     End If
   End With
   '
   With BloqArchB
     .Edit
     !USERNAME = USERNAME$
     !TERMNAME = NOMTER$
     !INITDATE = Now
     .Update
   End With
   '
End Sub

Sub LIBERTABLA(BATAB$)
   '
   BASETABLA$ = TRIM$(UCase$(BATAB$))
   '
   Call ABRETACEN
10 With BloqArchA
     .FindFirst "IDENTABLA = '" & BASETABLA$ & "'"
     If .NoMatch = True Then
       Exit Sub
     End If
   End With
   '
20 With BloqArchB
     .FindFirst "IDENTABLA = '" & BASETABLA$ & "'"
     If .NoMatch = True Then
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !IDENTABLA = BASETABLA$
       .Update
       BloqArchB.Close
       GoTo 20
     End If
   End With
   '
   On Error Resume Next
   With BloqArchB
     .Edit
     !USERNAME = ""
     !TERMNAME = ""
     !INITDATE = ""
     .Update
   End With
   '
   BloqArchA.Update
   '
End Sub

Function ESTABIER%(YABI%, LUAI$)
   ESTX% = 0
   If YABI% > 0 Then
     If LUAI$ = LUACC$ Then
       ESTX% = 1
       GoTo 90
     End If
   End If
   '
   YABI% = 0
   LUAI$ = LUACC$
   '
90 ESTABIER% = ESTX%
End Function

Sub VERISTRU(BASEDAT$)
   '
   Dim BADASTRU As Database
   '
   LCD$ = LUDAT$
   LCB$ = LUBAS$
   '
   If BASEDAT$ = "" Then
     Call MENSERR(24, "Error de Invocación Rutina 'VERISTRU'.\Nombre de Base de Datos no Válido.")
     Call FUERASYS
   End If
   '
 'LUBAS$ = "X:\BRASBOM\FLEXOFT\WINDOWS\"
   If EXISTE%(LUBAS$ + "BADASTRU.MDB") < 1 Then
     Call MENSERR(24, "Error de Instalación Básica.\Falta 'BADASTRU.MDB' En Carpeta del Sistema Básico.")
     Call FUERASYS
   End If
   '
'  AA1 = FileLen(LUBAS$ + "BADASTRU.MDB")
'  AA2 = FileLen(LUDAT$ + "BADASTRU.MDB")
'  AA3 = FileDateTime(LUBAS$ + "BADASTRU.MDB")
'  AA4 = FileDateTime(LUDAT$ + "BADASTRU.MDB")
'   If EXISTE%(LUDAT$ + "BADASTRU.MDB") > 0 Then
'     If FileLen(LUBAS$ + "BADASTRU.MDB") = FileLen(LUDAT$ + "BADASTRU.MDB") Then
'       If FileDateTime(LUBAS$ + "BADASTRU.MDB") = FileDateTime(LUDAT$ + "BADASTRU.MDB") Then
'         GoTo 90
'       End If
'     End If
     '
     If EXISTE%(LUDAT$ + "BADASTRU.MDB") < 1 Then GoTo 18
     If FileLen(LUBAS$ + "BADASTRU.MDB") <> FileLen(LUDAT$ + "BADASTRU.MDB") Then GoTo 18
     If FileDateTime(LUBAS$ + "BADASTRU.MDB") <> FileDateTime(LUDAT$ + "BADASTRU.MDB") Then GoTo 18
     '
     ' comparar si existen todos los campos y son iguales
     Dim BASELUDA As Database
     On Error Resume Next
     Set BASELUDA = OpenDatabase(LUACC$ + BASEDAT$ + ".MDB")
       If Err = 3343 Then   ' BASE DE DATOS DAÑABA
         ' ESTO ES PARA PRODUCIR LA AUTO-REPARACIóN DE LA BASE DE DATOS
         ' POR COMPACTACIóN
         On Error GoTo 0
         '
         On Error Resume Next
         Kill LUACC$ + BASEDAT$ + "1.MDB"
         On Error GoTo 0
         On Error Resume Next
         CompactDatabase LUACC$ + BASEDAT$ + ".MDB", LUACC$ + BASEDAT$ + "1.MDB"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         FileCopy LUACC$ + BASEDAT$ + ".MDB", LUACC$ + BASEDAT$ + ".OLD"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         FileCopy LUACC$ + BASEDAT$ + "1.MDB", LUACC$ + BASEDAT$ + ".MDB"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         Set BASELUDA = OpenDatabase(LUACC$ + BASEDAT$ + ".MDB")
         If Err Then GoTo 18
         GoSub RECUBASE
         GoTo 12   ' pudo recuperar el error
       End If
       '
       If Err Then GoTo 18 ' No existe la base de datos
       '
12   On Error GoTo 0
     LUBASTRU$ = LUBAS$
     If VERDEMO% > 0 Then LUBASTRU$ = LUDAT$ ' porque en la version demo da un error
     Set BADASTRU = OpenDatabase(LUBASTRU$ + "BADASTRU.MDB")
     Dim LICAMBASE As Recordset
     Set LICAMBASE = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS WHERE BaDatos = '" & UCase$(BASEDAT$) & "'", dbOpenSnapshot)
     '
     COINTOCAM% = 0
     On Error Resume Next
     With LICAMBASE
       .MoveLast
       If Err < 1 Then
         .MoveFirst
         On Error GoTo 0
         Do While Not .EOF
           NOMTABLA$ = TRIM$(UCase$(!TABLA))
           NOMCAMPO$ = TRIM$(UCase$(!CAMPO))
           TICAMPO% = !TIPO
           LOCAMPO% = !LONGITUD
           '
           For K1% = 1 To BASELUDA.TableDefs.Count
             If TRIM$(UCase$(BASELUDA.TableDefs(K1% - 1).Name)) = NOMTABLA$ Then
               For K2% = 1 To BASELUDA.TableDefs(K1% - 1).Fields.Count
                 If TRIM$(UCase$(BASELUDA.TableDefs(K1% - 1).Fields(K2% - 1).Name)) = NOMCAMPO$ Then
                   If BASELUDA.TableDefs(K1% - 1).Fields(K2% - 1).Type = TICAMPO% Then
                     'If (TICAMPO% <> 10 And TICAMPO% <> 12) Or LOCAMPO% = 0 Or BASELUDA.TableDefs(K1% - 1).Fields(K2% - 1).Size = LOCAMPO% Then
                       GoTo 15
                     'End If
                   End If
                 End If
               Next K2%
             End If
           Next K1%
           '
           BADASTRU.Close
           If INTERPRE% = 1 Then
             Call COMUNI("No Coinciden " + TRIM$(NOMTABLA$) + " y " + NOMCAMPO$ & " y " & TICAMPO%)
           End If
           GoTo 18   ' ALGO NO COINCIDIO - HAY QUE REORGANIZAR LA B.DATOS
           '
15       .MoveNext
         Loop
       End If
     End With
     '
     BADASTRU.Close
     COINTOCAM% = 1
     GoTo 30    ' COINCIDEN TODOS LOS CAMPOS
   'End If
   '
18 On Error GoTo 0
   TTXX$ = "FLEXOFT debe Ahora Actualizar\"
   TTXX$ = TTXX$ + "la Estructura de las Bases de Datos.\  \"
   TTXX$ = TTXX$ + "Cierre las Aplicaciones FLEXOFT\"
   TTXX$ = TTXX$ + "en TODAS las Estaciones de Trabajo de la Red,\"
   TTXX$ = TTXX$ + "Dejando Abierta Unicamente esta Aplicación."
   If COMALTER%(TTXX$ + "\\Continuar\Cancelar") <> 1 Then
     Call FUERASYS
   End If
   '
   On Error Resume Next
   BADASTRU.Close
   On Error GoTo 0
   '
20 If LUBAS$ <> LUDAT$ Then
     LUBA1$ = REPLACAR$(LUBAS$, "\", "/")
     LUDA1$ = REPLACAR$(LUDAT$, "\", "/")
     On Error Resume Next
     FileCopy LUBAS$ + "BADASTRU.MDB", LUDAT$ + "BADASTRU.MDB"
     If Err Then
       ERREX& = Err
       On Error GoTo 0
25     If COMALTER%("Error (" + TRIM$(Str$(ERREX&)) + ") de Copia de Archivo 'BADASTRU.MDB'\desde " + LUBA1$ + " hacia " + LUDA1$ + ".\\Reintentar\Abandonar") = 1 Then GoTo 20
       If INTERPRE% = 1 Then GoTo 90
       Call FUERASYS
     End If
   End If
   '
   If EXISTE%(LUDAT$ + "BADASTRU.MDB") > 0 Then
     If FileLen(LUBAS$ + "BADASTRU.MDB") = FileLen(LUDAT$ + "BADASTRU.MDB") Then
       If FileDateTime(LUBAS$ + "BADASTRU.MDB") = FileDateTime(LUDAT$ + "BADASTRU.MDB") Then
         GoTo 30
       End If
     End If
   End If
   GoTo 25
   '
30 Dim wrkPredeterminado As Workspace
   Set wrkPredeterminado = DBEngine.Workspaces(0)
   '
   LUBASTRU$ = LUBAS$
   If VERDEMO% > 0 Then LUBASTRU$ = LUDAT$ ' porque en la version demo da un error
   Set BADASTRU = OpenDatabase(LUBASTRU$ + "BADASTRU.MDB")
   Dim LISTABLAS As Recordset
   Set LISTABLAS = BADASTRU.OpenRecordset("SELECT DISTINCT * FROM BATABLAS WHERE BaDatos = '" & UCase$(BASEDAT$) & "'", dbOpenSnapshot)
   '
   ' crea la base de datos
   On Error Resume Next
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + BASEDAT$ + ".MDB", dbLangGeneral)
   Dim BASEDATO As Database
   Set BASEDATO = OpenDatabase(LUACC$ + BASEDAT$ + ".MDB")
   If Err = 3343 Then
     On Error GoTo 0
     BD1$ = REPLACAR$(LUACC$ + BASEDAT$ + ".MDB", "\", "/")
     Call MENSERR(24, "Error 3343. Imposible Abrir Base de Datos\" + BD1$ + ".\Consulte !!!")
     Call FUERASYS
   End If
   On Error GoTo 0
   '
   CATABLAS% = 0
   On Error Resume Next
   LISTABLAS.MoveLast
   CATABLAS% = LISTABLAS.RecordCount
   '
   YACREADA% = 0
   If CATABLAS% > 0 Then
     LISTABLAS.MoveFirst
     '
     Do While Not LISTABLAS.EOF
       NOMTABLA$ = LISTABLAS!TABLA
       On Error GoTo 0
       '
       If COINTOCAM% = 0 Then
         Set LISCAMPOS = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS WHERE Badatos = '" & UCase$(BASEDAT$) & "' AND Tabla = '" & NOMTABLA$ & "'", dbOpenSnapshot)
         LISCAMPOS.MoveLast
         CACAMPOS% = LISCAMPOS.RecordCount
         '
         YACREADA% = 0
         If CACAMPOS% > 0 Then
           LISCAMPOS.MoveFirst
           Do While Not LISCAMPOS.EOF
             '
             NOMCAMPO$ = LISCAMPOS!CAMPO
             TIPOVA = LISCAMPOS!TIPO
             LOCAM% = LISCAMPOS!LONGITUD
             '
             On Error Resume Next
             If YACREADA% = 0 Then
                Set CREATABLA = BASEDATO.CreateTableDef(NOMTABLA$)
                CREATABLA.Fields.Append CREATABLA.CreateField(NOMCAMPO$, TIPOVA, LOCAM%)
                BASEDATO.TableDefs.Append CREATABLA
                Set CREACAMPOS = BASEDATO.TableDefs(NOMTABLA$)
                YACREADA% = 1
             End If
             '
             With CREACAMPOS.Fields
               .Append CREACAMPOS.CreateField(NOMCAMPO$, TIPOVA, LOCAM%)
               If Err = 3197 Then
                 Call COMUNI("Imposible Actualizar Campo " + NOMCAMPO$ + ".\Recomendado Compactar Bases de Datos.")
               End If
             End With
             On Error GoTo 0
             '
           LISCAMPOS.MoveNext
           Loop
         End If
       End If
       '
       CACAMPOS1% = BASEDATO.TableDefs(NOMTABLA$).Fields.Count
       '
       On Error Resume Next
       For KU% = 1 To CACAMPOS1% ' PARA QUE TOME LOS CAMPOS REALES DE LA B.DATOS
         TICAM% = BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).Type
         If TICAM% = 1 Or TICAM% = 10 Or TICAM% = 12 Then
             'MsgBox BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).Name
             BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).AllowZeroLength = True
           ElseIf TICAM% = 1 Then
             BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).DefaultValue = True
           ElseIf TICAM% = 8 Then    ' FECHA
             BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).DefaultValue = CVDAT(0)
           Else
             BASEDATO.TableDefs(NOMTABLA$).Fields(KU% - 1).DefaultValue = 0
         End If
       Next KU%
       On Error GoTo 0
       '
       LISTABLAS.MoveNext
     Loop
   End If
   '
   BASEDATO.Close
   BADASTRU.Close
   Exit Sub
   '
   '
RECUBASE:
   '
   TERECU$ = Space$(128)
   HOII% = HOY(HO$)
   PRXX$ = TRIM$(UCase$(App.EXEName))
   USX$ = USERNAME$
   ITX$ = IDENTER$
   TRX$ = NOMTER$
   '
   Call REPLA(TERECU$, PRXX$ + ".EXE", 1, 16)
   Call REPLA(TERECU$, USX$, 17, 24)
   Call REPLA(TERECU$, TRX$, 45, 24)
   Call REPLA(TERECU$, HO$, 69, 8)
   Call REPLA(TERECU$, Time$, 79, 8)
   Call REPLA(TERECU$, BASEDAT$, 89, 28)
   Call REPLA(TERECU$, ITX$, 119, 8)
   '
   NRX1% = FILEOPEN%(LUDAT$ + "RECUBASE.PRN", 3, 128)
   Print #NRX1%, TERECU$
   Close #NRX1%
Return

90 End Sub

Sub ABREMOVSTK()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("STOCKORIGI")
     Set STOCKORIGI = OpenDatabase(LUACC$ + "STOCKORIGI.MDB")
   End If
   '
90 Set EncaStock = STOCKORIGI.OpenRecordset("ENCASTOCK", dbOpenDynaset)
   Set DetaStock = STOCKORIGI.OpenRecordset("DETASTOCK", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABREORTRA()
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("OTRABAJO")
     Set OTrabajo = OpenDatabase(LUACC$ + "OTRABAJO.MDB")
     YABIER% = 1
   End If
   '
90 Set OrdeTrab = OTrabajo.OpenRecordset("ORDETRAB", dbOpenDynaset)
   Set PlanCob = OTrabajo.OpenRecordset("PLACOB", dbOpenDynaset)
   Set GrupOtra = OTrabajo.OpenRecordset("GRUPOTRA", dbOpenDynaset)
   Set ListaMat = OTrabajo.OpenRecordset("LISTAMAT", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABREPRECAM()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("PRESUCAM")
     Set PresuCam = OpenDatabase(LUACC$ + "PRESUCAM.MDB")
     YABIER% = 1
   End If
   '
90 Set EncaPresu = PresuCam.OpenRecordset("ENCAPRESU", dbOpenDynaset)
   Set DetaPresu = PresuCam.OpenRecordset("DETAPRESU", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'OT-05-0856-WAR-28/11/05///
Sub ABREMAQUI()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("MAQUINARIAS")
     Set Maquinarias = OpenDatabase(LUACC$ + "MAQUINARIAS.MDB")
     YABIER% = 1
   End If
   '
90 Set Equipos = Maquinarias.OpenRecordset("EQUIPOS", dbOpenDynaset)
   Set GaleImag = Maquinarias.OpenRecordset("GALEIMAG", dbOpenDynaset)
   Set RepuMaq = Maquinarias.OpenRecordset("REPUMAQ", dbOpenDynaset)
   Set Contadores = Maquinarias.OpenRecordset("CONTADORES", dbOpenDynaset)
   Set SubConjuntos = Maquinarias.OpenRecordset("SUBCONJUNTOS", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub

Sub ABREOMAN()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$) < 1 Then
     Call VERISTRU("OMANTIN")
     Set Omantin = OpenDatabase(LUACC$ + "OMANTIN.MDB")
   End If
   '
90 Set OrdeMant = Omantin.OpenRecordset("ORDEMANT", dbOpenDynaset)
   Set RecuProp = Omantin.OpenRecordset("RECUPROP", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'OT-05-0856-WAR-FIN///

Sub AGRECAMPO(BASEDA$, TABLAX$, NCAMPO$, TCAMPO, LCAMPO)
   '
   Set BASDAT = OpenDatabase(LUACC$ + BASEDA$ + ".MDB")
   Set CREACA = BASDAT.TableDefs(TABLAX$)
   '
   ' Verifica pre-existencia del campo en la tabla
   For ICAMPO& = 1 To CREACA.Fields.Count
     If CREACA.Fields(ICAMPO& - 1).Name = NCAMPO$ Then
       Exit Sub
     End If
   Next ICAMPO&
   '
   ' Agrega el campo a la tabla
   With CREACA.Fields
     On Error Resume Next
     .Append CREACA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
     If Err Then
       If Err = 3191 Then
           Exit Sub
         Else
           Call MENSERR(24, "Error " & Err & " al Crear Campo '" + NCAMPO$ + "' en Tabla '" + TABLAX$ + "'.")
       End If
     End If
   End With
   '
End Sub
'

