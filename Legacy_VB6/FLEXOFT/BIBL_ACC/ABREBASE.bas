Attribute VB_Name = "ABREBASE"
Public wrkPredeterminado As Workspace

Public Contable As Database
Public Tabejcon As Recordset
Public Plancuen As Recordset
Public Encasien As Recordset
Public Detasien As Recordset
Public Balasec As Recordset
Public TaProyec As Recordset '\\\OT-06-0244-WAR-15/06/06///
'
Public Articulos As Database
Public Master As Recordset
Public Formulas As Recordset
Public Secoper As Recordset
Public TmusOper As Recordset '\\\OT-06-0252-22/06/06///
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
Public InfoProy As Recordset '\\\OT-06-0252-WAR-26/06/06///

Public compras As Database
Public OcomEnca As Recordset
Public OcomDeta As Recordset
Public EncoCabi As Recordset '\\\OT-06-0149-WAR-19/04/06///
Public DetoCabi As Recordset '\\\OT-06-0149-WAR-19/04/06///
Public AutoEnt As Recordset '\\\OT-06-0149-WAR-19/04/06///
'
Public Ventas As Database
Public PresuVen As Recordset
Public PedEnca As Recordset
Public PedDeta As Recordset
Public ListaPrec As Recordset
Public Preventa As Recordset
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
Public Despecial As Recordset
Public Docadjun As Recordset
'
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
Public SeTaReco As Recordset
Public ORecorri As Recordset
'
Public BDEQUIPOS As Database
Public BASEENUSO As Database
'
Public ACTICARGA%
Dim HUBOERROR%
'
Function FILTSQL$(XXYZT$)
    FILTSQL$ = XXYZT$
End Function
'
Function VALOBADA(AAA1$, AAA2$, AAA3$, Optional AAA4$)
    VALOBADA = ""
End Function
'
Sub EJEMPLOPARACREARINDICES()
   Call APERBASDAT("CABAN")
   AA1 = CueCorri.TableDefs("CAJABANC").Indexes.Count
   For AB1 = 1 To 13
   MsgBox CueCorri.TableDefs("CAJABANC").Indexes(AB1).Name
   Next AB1
End Sub
'
Sub APERBASDAT(NBASE$)
   '
   RUTIACT$ = "ABREBASDAT"
   On Error GoTo 90    ' BUSCA CAPTURAR ERROR 91
   '
   NBAS$ = UCase$(NBASE$)
   Select Case NBAS$
     Case "MASTER"
        Call ABREMASTER
     Case "STRUC"
        Call ABRESTRUC
     Case "CABAN"
        Call ABRECABAN
     Case "STOCKS"
        Call ABRESTOCKS
     Case "STOCKVIE"
        Call ABRESTOCKVIE
     Case "COMPRAS"
        Call ABRECOMPRAS
     Case "PLANCUEN"
        Call ABREPLANCUEN
     Case "RESERVA"
        Call ABRESERVA
     Case "ORFAB"
        Call ABREORFAB
     Case "PEDCLI"
        Call ABREPEDCLI
     Case "CLIEN"
        Call ABRECLIEN
     Case "PROVE"
        Call ABREPROVE
     Case Else
        MsgBox "Imposible Abrir B.Datos '" & NBAS$ & "'", vbExclamation
        Call FINAL("")
   End Select
   '
   RUTIACT$ = ""
   '
Exit Sub
'*******************************************************************************
90 TTXX$ = "Se ha Producido un Error no Recuperable (" & Err.Number & ") en la Rutina " & RUTIACT$ & " al Abrir " & NBAS$
   TTXX$ = TTXX$ + ".\Capture o Imprima esta Pantalla, Enviándola por Mail a asistec@flexoft.com.ar."
   TTXX$ = TTXX$ + "\   \Será Respondido a la Brevedad. Muchas Gracias por su Colaboración."
   Resume 91
   '
91 On Error GoTo 0
   RUTIACT$ = ""
   MsgBox FORMESS$(TTXX$), vbExclamation
   Call FINAL("")
End Sub
'
Sub ABREMASTER()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREMASTER"
   '
   If ESTABIER%(YABIER%, LUACCA$, "ARTICULOS") < 1 Then
     Call VERISTRU("ARTICULOS", YAINDEX%)
     Set Articulos = OpenDatabase(LUACREPLA$("ARTICULOS") + "ARTICULOS.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       Articulos.Execute "CREATE INDEX ICodint ON Master ([CodInt]);"
       On Error GoTo 0
     End If
     '
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Master.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set Master = Articulos.OpenRecordset("MASTER", dbOpenDynaset)
   'Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABRESTRUC()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRESTRUC"
   '
'   SCMP = Screen.MousePointer
'   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "ARTICULOS") < 1 Then
     Call VERISTRU("ARTICULOS", YAINDEX%)
     Set Articulos = OpenDatabase(LUACREPLA$("ARTICULOS") + "ARTICULOS.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Formulas.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set Formulas = Articulos.OpenRecordset("FORMULAS", dbOpenDynaset)
   'Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABRESECOP()
   '
   Static YABIER%, LUACCA$
   '
'   SCMP = Screen.MousePointer
'   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "ARTICULOS") < 1 Then
     Call VERISTRU("ARTICULOS", YAINDEX%)
     Set Articulos = OpenDatabase(LUACREPLA$("ARTICULOS") + "ARTICULOS.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Secoper.MoveFirst
     If Err Then GoTo 90
   Exit Sub
   '
90 On Error GoTo 0
   Set Secoper = Articulos.OpenRecordset("SECOPER", dbOpenDynaset)
   Set TmusOper = Articulos.OpenRecordset("TMUSOPER", dbOpenDynaset)
   'Screen.MousePointer = SCMP
   '
End Sub
'
Sub ABRECLIEN()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRECLIEN"
   '
   If ESTABIER%(YABIER%, LUACCA$, "BADACLIEN") < 1 Then
     Call VERISTRU("BADACLIEN", YAINDEX%)
     Set Badaclie = OpenDatabase(LUACREPLA$("BADACLIEN") + "BADACLIEN.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       Badaclie.Execute "CREATE INDEX INume_Cli ON Deudor12 ([Nume_Cli]);"
       Badaclie.Execute "CREATE INDEX IAgendaOriginal ON AgendaOriginal ([Nume_Cli]);"
       Badaclie.Execute "CREATE INDEX IID_DIA ON AgendaOriginal ([ID_DIA]);"
       Badaclie.Execute "CREATE INDEX IID_ORIGINAL ON Agenda ([ID_ORIGINAL]);"
       Badaclie.Execute "CREATE INDEX IPEDIDO ON Agenda ([PEDIDO]);"
       'CREA LOS CAMPOS ID
       If EXISTECAMPOENBADACLIEN%("AgendaOriginal", "ID") < 1 Then
          sqlx$ = "ALTER TABLE AgendaOriginal ADD COLUMN ID COUNTER CONSTRAINT ClavePrincipal PRIMARY KEY"
          Badaclie.Execute (sqlx$)
       End If
       If EXISTECAMPOENBADACLIEN%("Agenda", "ID") < 1 Then
         sqlx$ = "ALTER TABLE Agenda ADD COLUMN ID COUNTER CONSTRAINT ClavePrincipal PRIMARY KEY"
         Badaclie.Execute (sqlx$)
      End If
      On Error GoTo 0
     End If
     '
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Deudor12.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set Deudor12 = Badaclie.OpenRecordset("DEUDOR12", dbOpenDynaset)
   '
   RUTIACT$ = RUTIANT$
End Sub
Function EXISTECAMPOENBADACLIEN%(TABLA$, CAMPO$)
    EXISTECAMPOENBADACLIEN% = 0
    Dim obj_Database As Database
    Dim obj_TableDef As TableDef
    Dim obj_Field As Field
    ' Abre la base de datos
    Set obj_Database = Workspaces(0).OpenDatabase(LUDAT$ + "BdAccess\BADACLIEN")
    Set obj_TableDef = obj_Database.TableDefs(TABLA$)
    
    On Error Resume Next
    'referencia al campo que se quiere comprobar. Si da error al referirnos a el, es por que no existe
    Set obj_Field = obj_TableDef.Fields(CAMPO$)
    ' Comprueba si hay error
    If Err.Number = 0 Then
        EXISTECAMPOENBADACLIEN% = 1
    End If
      
    ' cierra la base
    obj_Database.Close
      
    If Not obj_Database Is Nothing Then
       Set obj_Database = Nothing
    End If
    On Error GoTo 0


End Function


'End Function
'
Sub ABREPROVE()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREPROVE"
   '
   If ESTABIER%(YABIER%, LUACCA$, "BADAPROVE") < 1 Then
     Call VERISTRU("BADAPROVE", YAINDEX%)
     Set Badaprov = OpenDatabase(LUACREPLA$("BADAPROVE") + "BADAPROVE.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Proved12.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set Proved12 = Badaprov.OpenRecordset("Proved12", dbOpenDynaset)
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABRECABAN()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRECABAN"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "CUECORRI") < 1 Then
     Call VERISTRU("CUECORRI", YAINDEX%)
     Set CueCorri = OpenDatabase(LUACREPLA$("CUECORRI") + "CUECORRI.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       CueCorri.Execute "CREATE INDEX ICueContaI ON cajabanc ([CueContaI]);"
       CueCorri.Execute "CREATE INDEX IFechConta ON cajabanc ([FechConta]);"
       CueCorri.Execute "CREATE INDEX IFechEmisi ON cajabanc ([FechEmisi]);"
       CueCorri.Execute "CREATE INDEX ITipoMovim ON cajabanc ([TipoMovim]);"
       CueCorri.Execute "CREATE INDEX IOpciMovim ON cajabanc ([OpciMovim]);"
       CueCorri.Execute "CREATE INDEX IVariMovim ON cajabanc ([VariMovim]);"
       CueCorri.Execute "CREATE INDEX ICodiComCor ON cajabanc ([CodiCom_Cor]);"
       CueCorri.Execute "CREATE INDEX ICodiComLar ON cajabanc ([CodiCom_Lar]);"
       CueCorri.Execute "CREATE INDEX INuAsidiar ON cajabanc ([NuAsidiar]);"
       CueCorri.Execute "CREATE INDEX INumeComp ON cajabanc ([NumeComp]);"
       CueCorri.Execute "CREATE INDEX INumeTalo ON cajabanc ([NumeTalo]);"
       CueCorri.Execute "CREATE INDEX INuClien ON cajabanc ([NuClien]);"
       CueCorri.Execute "CREATE INDEX INuProve ON cajabanc ([NuProve]);"
       CueCorri.Execute "CREATE INDEX IDescriMov ON cajabanc ([DescriMov]);"
       CueCorri.Execute "CREATE INDEX IValAbsComp ON cajabanc ([ValAbsComp]);"
       On Error GoTo 0
     End If
     '
     YABIER% = 1
   End If
   '
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
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set CajaBanc = CueCorri.OpenRecordset("CAJABANC", dbOpenDynaset)
   Set SacrePen = CueCorri.OpenRecordset("SACREPEN", dbOpenDynaset)
   Set SaDeuPen = CueCorri.OpenRecordset("SADEUPEN", dbOpenDynaset)
   Set ApliPago = CueCorri.OpenRecordset("APLIPAGO", dbOpenDynaset)
   Set ApliCobra = CueCorri.OpenRecordset("APLICOBRA", dbOpenDynaset)
   'Set EncMovCc = CueCorri.OpenRecordset("ENCMOVCC", dbOpenDynaset) 'Add p/import.
   'Set DetMovCc = CueCorri.OpenRecordset("DETMOVCC", dbOpenDynaset) 'add p/import.
   '
'99 YABIER% = 1
'   Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABREPLANCUEN()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREPLANCUEN"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "CONTAGEN") < 1 Then
     Call VERISTRU("CONTAGEN", YAINDEX%)
     Set Contable = OpenDatabase(LUACREPLA$("CONTAGEN") + "CONTAGEN.MDB")
     YABIER% = 1
   End If
   '
   If YAINDEX% < 1 Then
     On Error Resume Next
     Contable.Execute "CREATE INDEX ICodInCue ON Plancuen ([CodInCue]);"
     On Error GoTo 0
   End If
' Crea el índice ÍndNue en la tabla Empleados.
' On Error Resume Next
' Contable.Execute "CREATE INDEX ICueCoInt ON DetaAsien ([CueCoInt]);"
' Contable.Execute "CREATE INDEX IFechasien ON DetaAsien ([Fechasien]);"
' On Error GoTo 0

   On Error Resume Next
   Tabejcon.MoveFirst
     If Err Then GoTo 90
   Plancuen.MoveFirst
     If Err Then GoTo 90
   Encasien.MoveFirst
     If Err Then GoTo 90
   Detasien.MoveFirst
     If Err Then GoTo 90
   TaProyec.MoveFirst '\\\OT-06-0244-WAR-15/06/06///
     If Err Then GoTo 90 '\\\OT-06-0244-WAR-15/06/06///
   RUTIACT$ = RUTIANT$
   Exit Sub   'GoTo 99
   '
90 On Error GoTo 0
   Set Tabejcon = Contable.OpenRecordset("TABEJCON", dbOpenDynaset, dbReadOnly)
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenDynaset, dbReadOnly)
   Set Encasien = Contable.OpenRecordset("ENCABEASIEN", dbOpenDynaset, dbReadOnly)
   Set Detasien = Contable.OpenRecordset("DETAASIEN", dbOpenDynaset, dbReadOnly)
   Set TaProyec = Contable.OpenRecordset("TAPROYEC", dbOpenDynaset) '\\\OT-06-0244-WAR-15/06/06///
   '
'99 YABIER% = 1
'   Screen.MousePointer = SCMP
   RUTIACT$ = RUTIANT$
End Sub

Sub ABRECOMPRAS()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRECOMPRAS"
   '
   If ESTABIER%(YABIER%, LUACCA$, "COMPRAS") < 1 Then
     Call VERISTRU("COMPRAS", YAINDEX%)
     Set compras = OpenDatabase(LUACREPLA$("COMPRAS") + "COMPRAS.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   OcomEnca.MoveFirst
     If Err Then GoTo 90
   OcomDeta.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set OcomEnca = compras.OpenRecordset("OCOMENCA", dbOpenDynaset)
   Set OcomDeta = compras.OpenRecordset("OCOMDETA", dbOpenDynaset)
   '\\\OT-06-0149-WAR-18/04/06///
   Set EncoCabi = compras.OpenRecordset("ENCOCABI", dbOpenDynaset)
   Set DetoCabi = compras.OpenRecordset("DETOCABI", dbOpenDynaset)
   Set AutoEnt = compras.OpenRecordset("AUTOENT", dbOpenDynaset)
   '\\\OT-06-0149-WAR-FIN///
   '
   'YABIER% = 1
   'Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub

Sub ABREPRESVEN()
   '
   Static YABIER%, LUACCA$
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "VENTAS") < 1 Then
     Call VERISTRU("VENTAS", YAINDEX%)
     Set Ventas = OpenDatabase(LUACREPLA$("VENTAS") + "VENTAS.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   PresuVen.MoveFirst
     If Err Then GoTo 90
   Exit Sub
   '
90 On Error GoTo 0
   Set PresuVen = Ventas.OpenRecordset("PRESUVEN", dbOpenDynaset)
   'YABIER% = 1
   'Screen.MousePointer = SCMP
   '
End Sub

Sub ABREPEDCLI()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREPEDCLI"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "VENTAS") < 1 Then
     Call VERISTRU("VENTAS", YAINDEX%)
     Set Ventas = OpenDatabase(LUACREPLA$("VENTAS") + "VENTAS.MDB")
     YABIER% = 1
   End If
   '
   On Error Resume Next
   PedEnca.MoveFirst
     If Err Then GoTo 90
   PedDeta.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set PedEnca = Ventas.OpenRecordset("PEDENCA", dbOpenDynaset)
   Set PedDeta = Ventas.OpenRecordset("PEDDETA", dbOpenDynaset)
   Set ListaPrec = Ventas.OpenRecordset("LISTAPRE", dbOpenDynaset)
   Set Preventa = Ventas.OpenRecordset("PREVENTA", dbOpenDynaset)
   'YABIER% = 1
   'Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub

Sub ABRESTOCKS()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRESTOCKS"
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "STOCKS") < 1 Then
     Call VERISTRU("STOCKS", YAINDEX%)
     Set Stocks = OpenDatabase(LUACREPLA$("STOCKS") + "STOCKS.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       Stocks.Execute "CREATE INDEX ICod_Inte ON Movisto ([Cod_Inte]);"
       Stocks.Execute "CREATE INDEX ICodiMovi ON Movisto ([CodiMovi]);"
       Stocks.Execute "CREATE INDEX IDoPriCor ON Movisto ([DoPriCor]);"
       Stocks.Execute "CREATE INDEX IDoPriLar ON Movisto ([DoPriLar]);"
       Stocks.Execute "CREATE INDEX INume_Clie ON Movisto ([Nume_Clie]);"
       Stocks.Execute "CREATE INDEX IFechMov ON Movisto ([FechMov]);"
       Stocks.Execute "CREATE INDEX IFeReMovi ON Movisto ([FeReMovi]);"
       Stocks.Execute "CREATE INDEX IReferCor ON Movisto ([ReferCor]);"
       Stocks.Execute "CREATE INDEX ICod_Inte ON Bolrecep ([Cod_Inte]);"
       Stocks.Execute "CREATE INDEX ICodiInte ON SalStock ([CodiInte]);"
       Stocks.Execute "CREATE INDEX INuPedCli ON Movisto ([NuPedCli]);"
       On Error GoTo 0
     End If
     YABIER% = 1
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
   RUTIACT$ = RUTIANT$
   Exit Sub   'GoTo 99
   '
90 On Error GoTo 0
   Set Movisto = Stocks.OpenRecordset("MOVISTO", dbOpenDynaset)
   Set SalStock = Stocks.OpenRecordset("SALSTOCK", dbOpenDynaset)
   Set Bolrecep = Stocks.OpenRecordset("BOLRECEP", dbOpenDynaset)
   Set MaConsig = Stocks.OpenRecordset("MACONSIG", dbOpenDynaset)
   '
'99 YABIER% = 1
'   Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub

Sub ABRESTOCKVIE()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRESTOCKVIE"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "STOCKVIE") < 1 Then
     Call VERISTRU("STOCKVIE", YAINDEX%)
     Set StockVie = OpenDatabase(LUACREPLA$("STOCKVIE") + "STOCKVIE.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       StockVie.Execute "CREATE INDEX ICod_Inte ON Movisto1 ([Cod_Inte]);"
       StockVie.Execute "CREATE INDEX ICodiMovi ON Movisto1 ([CodiMovi]);"
       StockVie.Execute "CREATE INDEX IFechMov ON Movisto1 ([FechMov]);"
       StockVie.Execute "CREATE INDEX IFeReMovi ON Movisto1 ([FeReMovi]);"
       On Error GoTo 0
     End If
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Movisto1.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set Movisto1 = StockVie.OpenRecordset("MOVISTO", dbOpenDynaset)
'   YABIER% = 1
'   Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABREEQUIPOS()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREEQUIPOS"
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "EQUIPOS") < 1 Then
     Call VERISTRU("EQUIPOS", YAINDEX%)
     Set BDEQUIPOS = OpenDatabase(LUACREPLA$("EQUIPOS") + "EQUIPOS.MDB")
     '
     YABIER% = 1
     '
   End If
   '
   On Error Resume Next
   TANUMSE.MoveFirst
     If Err Then GoTo 90
   Exit Sub
   '
90 On Error GoTo 0
   Set TANUMSE = BDEQUIPOS.OpenRecordset("TANUMSE", dbOpenDynaset)
   RUTIACT$ = RUTIANT$
End Sub

Sub ABRESERVA()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABRESERVA"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "STOCKS") < 1 Then
     Call VERISTRU("STOCKS", YAINDEX%)
     Set Stocks = OpenDatabase(LUACREPLA$("STOCKS") + "STOCKS.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       Stocks.Execute "CREATE INDEX ICod_Inte ON Reserva ([Cod_Inte]);"
       On Error GoTo 0
     End If
     YABIER% = 1
   End If
   '
   On Error Resume Next
   Reserva.MoveFirst
     If Err Then GoTo 90
   RUTIACT$ = RUTIANT$
   Exit Sub   ' GoTo 99
   '
90 On Error GoTo 0
   Set Reserva = Stocks.OpenRecordset("RESERVA", dbOpenDynaset)
   '
'99 YABIER% = 1
'   Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Sub ABREORFAB()
   '
   Static YABIER%, LUACCA$
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ABREORFAB"
   '
   'SCMP = Screen.MousePointer
   'Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "OFABRIC") < 1 Then
     Call VERISTRU("OFABRIC", YAINDEX%)
     Set Ofabric = OpenDatabase(LUACREPLA$("OFABRIC") + "OFABRIC.MDB")
     '
     If YAINDEX% < 1 Then
       On Error Resume Next
       Ofabric.Execute "CREATE INDEX IOrfNop1 ON OperFab ([IdSubOr] , [NumeOper]);"
       Ofabric.Execute "CREATE INDEX IOrfNop2 ON InfoFab ([IdSubOr] , [NumeOper]);"
       On Error GoTo 0
     End If
     YABIER% = 1
   End If
   '
   On Error Resume Next
   OrdeFabr.MoveFirst
     If Err Then GoTo 90
   OperFab.MoveFirst
     If Err Then GoTo 90
   InfoFab.MoveFirst
     If Err Then GoTo 90
   InfoProy.MoveFirst '\\\OT-06-0252-WAR-26/06/06///
     If Err Then GoTo 90 '\\\OT-06-0252-WAR-26/06/06///
   RUTIACT$ = RUTIANT$
   Exit Sub
   '
90 On Error GoTo 0
   Set OrdeFabr = Ofabric.OpenRecordset("ORDEFABR", dbOpenDynaset)
   Set OperFab = Ofabric.OpenRecordset("OPERFAB", dbOpenDynaset)
   Set InfoFab = Ofabric.OpenRecordset("INFOFAB", dbOpenDynaset)
   On Error Resume Next
   Set InfoProy = Ofabric.OpenRecordset("INFOPROY", dbOpenDynaset) '\\\OT-06-0252-WAR-26/06/06///
   On Error GoTo 0
   'YABIER% = 1
   'Screen.MousePointer = SCMP
   '
   RUTIACT$ = RUTIANT$
End Sub
'
Function ESTABIER%(YABI%, LUAI$, NOMBASE$)
   '
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "ESTABIER"
   '
   ESTX% = 0
   If YABI% > 0 Then
     If LUAI$ = LUACREPLA$(NOMBASE$) Then
       ESTX% = 1
       GoTo 90
     End If
   End If
   '
   YABI% = 0
   LUAI$ = LUACREPLA$(NOMBASE$)
   '
90 ESTABIER% = ESTX%
   RUTIACT$ = RUTIANT$
   '
End Function

Sub VERISTRU(BASEDAT$, YAINDEX%)
   '
   Static CTXX%, CAREGI&, VECBASE$(), FECHAC#()
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "VERISTRU"
   '
   If CTXX% < 1 Then
     Set wrkPredeterminado = DBEngine.Workspaces(0)
     TIBASE$ = "MDB"
     CTXX% = 1
   End If
   '
   YAINDEX% = 0
   LCD$ = LUDAT$
   LCB$ = LUBAS$
   '
   ' FECHA Y HORA DEL EJECUTABLE ACTUAL
   FDT# = 0
   On Error Resume Next
   FDT# = FileDateTime(App.PATH + "\" + App.EXEName + ".EXE")
   On Error GoTo 0
   If FDT# < 1 Then
     RUTIACT$ = RUTIANT$
     Exit Sub
     FDT# = Now
   End If
   '
   If CAREGI& < 1 Then
     VBAS$ = LOADFILE$(LUDAT$ + "ACSTRUBA.DAT")
     CAREGI& = Len(VBAS$) / 64
     ReDim Preserve VECBASE$(CAREGI& + 128), FECHAC#(CAREGI& + 128)
     For U& = 1 To CAREGI&
       VECBASE$(U&) = UCase$(TRIM$(Mid$(VBAS$, 64 * U& - 63, 56)))
       FECHAC#(U&) = CVD(Mid$(VBAS$, 64 * U& - 7, 8))
     Next U&
   End If
   '
   FULAC# = 0: REVEC& = 0
   For U& = 1 To CAREGI&
     If VECBASE$(U&) = UCase$(TRIM$(BASEDAT$)) Then
       FULAC# = FECHAC#(U&)
       REVEC& = U&
       GoTo 20
     End If
   Next U&
   CAREGI& = CAREGI& + 1
   VECBASE$(CAREGI&) = UCase$(TRIM$(BASEDAT$))
   FECHAC#(CAREGI&) = 0
   REVEC& = CAREGI&
   '
20 ' verifica si existe la base de datos
   EXIBA% = EXISTE%(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB")
   If EXIBA% < 1 Then GoTo 50 ' CREA LA BASE DE DATOS
   '
30 ' Verifica integridad de la base de datos
   Dim BASELUDA As Database
   VUELTA% = 1
35 On Error Resume Next
   Set BASELUDA = OpenDatabase(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB")
   '
   If Err = 0 Then
       On Error GoTo 0
       BASELUDA.Close
       GoTo 40                 ' VERIFICADA INTEGRIDAD DE LA BD
     ElseIf Err = 3343 Then    ' BASE DE DATOS DAÑABA
       ' ESTO ES PARA PRODUCIR LA AUTO-REPARACIóN DE LA BASE DE DATOS
       ' POR COMPACTACIóN
       On Error GoTo 0
       '
       If VUELTA% > 1 Then
         TTXX$ = "La Base de Datos '" + BASEDAT$ + ".MDB' Aparece Dañada (Error 3343)"
         TTXX$ = TTXX$ + "\   \Cierre FLEXOFT en TODAS las Estaciones de Trabajo de la Red"
         TTXX$ = TTXX$ + "\y Vuelva a Ejecutar la Presente Aplicación '" + App.EXEName + ".EXE'."
         TTXX$ = TTXX$ + "\   \En caso de Repetirse el Presente Mensaje, Ejecute 'Compactar Bases"
         TTXX$ = TTXX$ + "\de Datos' desde el Menú 'Setup' del Escritorio Principal FLEXOFT."
         TTXX$ = TTXX$ + "\  \Si nada de lo Anterior Resulta, Contacte a su Soporte FLEXOFT."
         Call MENSERR(24, TTXX$)
         Call FUERASYS
       End If
       '
       On Error Resume Next
       Kill LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB"
       On Error GoTo 0
       On Error Resume Next
       CompactDatabase LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB"
       If Err Then GoTo 38
       On Error GoTo 0
       '
       On Error Resume Next
       FileCopy LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + ".OLD"
       If Err Then GoTo 38
       On Error GoTo 0
       '
       On Error Resume Next
       FileCopy LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB"
       If Err Then GoTo 38
       On Error GoTo 0
       '
38     On Error GoTo 0
       BASELUDA.Close
       GoSub RECUBASE          ' ESCRIBE RECORD EN ARCHIVO DE RECUPERACION
       VUELTA% = 2
       GoTo 35
       '
     Else
       TTXX$ = "Error (" + TRIM$(Str$(Err)) + ") al Abrir la Base de Datos '" + BASEDAT$ + ".MDB'"
       TTXX$ = TTXX$ + "\en Aplicación '" + App.EXEName + ".EXE' del " + Format$(FDT#, "dd/mm/yy  hh:mm") + "."
       TTXX$ = TTXX$ + "\  \Contacte a su Soporte FLEXOFT"
       TTXX$ = TTXX$ + "\con los Datos del Presente Mensaje."
       Call MENSERR(24, TTXX$)
       Call FUERASYS
   End If
   '
40 ' Verifica que fecha de ultima actualicacion de la BD
   ' sea mayor o igual que la de este ejecutable
   If FULAC# >= FDT# - 0.0001 Then
     YAINDEX% = 1
     GoTo 90
   End If
   '
50 SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   HUBOERROR% = 0
   Call GENERBASE(BASEDAT$)
   If HUBOERROR% < 1 Then
     FECHAC#(REVEC&) = FDT#
     VBAS$ = ""
     For U& = 1 To CAREGI&
       VBAS$ = VBAS$ + AJUSTI$(VECBASE$(U&), 56)
       VBAS$ = VBAS$ + MKD$(FECHAC#(U&))
     Next U&
     Call SAVEFILE(LUDAT$ + "ACSTRUBA.DAT", VBAS$)
   End If
   Screen.MousePointer = SCMP
   GoTo 90
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

   RUTIACT$ = RUTIANT$
90 End Sub

Sub VERISTRUVIE()
Stop

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
     Set BASELUDA = OpenDatabase(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB")
       If Err = 3343 Then   ' BASE DE DATOS DAÑABA
         ' ESTO ES PARA PRODUCIR LA AUTO-REPARACIóN DE LA BASE DE DATOS
         ' POR COMPACTACIóN
         On Error GoTo 0
         '
         On Error Resume Next
         Kill LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB"
         On Error GoTo 0
         On Error Resume Next
         CompactDatabase LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         FileCopy LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + ".OLD"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         FileCopy LUACREPLA$(BASEDAT$) + BASEDAT$ + "1.MDB", LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB"
         If Err Then GoTo 18
         On Error GoTo 0
         '
         On Error Resume Next
         Set BASELUDA = OpenDatabase(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB")
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
           LOCAMPO% = !Longitud
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
       'If INTERPRE% = 1 Then GoTo 90
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
30 'Dim wrkPredeterminado As Workspace
   'Set wrkPredeterminado = DBEngine.Workspaces(0)
   '
   LUBASTRU$ = LUBAS$
   If VERDEMO% > 0 Then LUBASTRU$ = LUDAT$ ' porque en la version demo da un error
   Set BADASTRU = OpenDatabase(LUBASTRU$ + "BADASTRU.MDB")
   Dim LISTABLAS As Recordset
   Set LISTABLAS = BADASTRU.OpenRecordset("SELECT DISTINCT * FROM BATABLAS WHERE BaDatos = '" & UCase$(BASEDAT$) & "'", dbOpenSnapshot)
   '
   ' crea la base de datos
   On Error Resume Next
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", dbLangGeneral)
   Dim BASEDATO As Database
   Set BASEDATO = OpenDatabase(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB")
   If Err = 3343 Then
     On Error GoTo 0
     BD1$ = REPLACAR$(LUACREPLA$(BASEDAT$) + BASEDAT$ + ".MDB", "\", "/")
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
             LOCAM% = LISCAMPOS!Longitud
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

End Sub

Sub ABREMOVSTK()
   '
   Static YABIER%, LUACCA$
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   If ESTABIER%(YABIER%, LUACCA$, "STOCKORIGI") < 1 Then
     Call VERISTRU("STOCKORIGI", YAINDEX%)
     Set STOCKORIGI = OpenDatabase(LUACREPLA$("STOCKORIGI") + "STOCKORIGI.MDB")
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
   If ESTABIER%(YABIER%, LUACCA$, "OTRABAJO") < 1 Then
     Call VERISTRU("OTRABAJO", YAINDEX%)
     Set OTrabajo = OpenDatabase(LUACREPLA$("OTRABAJO") + "OTRABAJO.MDB")
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
   If ESTABIER%(YABIER%, LUACCA$, "PRESUCAM") < 1 Then
     Call VERISTRU("PRESUCAM", YAINDEX%)
     Set PresuCam = OpenDatabase(LUACREPLA$("PRESUCAM") + "PRESUCAM.MDB")
     YABIER% = 1
   End If
   '
90 Set EncaPresu = PresuCam.OpenRecordset("ENCAPRESU", dbOpenDynaset)
   Set DetaPresu = PresuCam.OpenRecordset("DETAPRESU", dbOpenDynaset)
   Set Despecial = PresuCam.OpenRecordset("DESPECIAL", dbOpenDynaset)
   Set Docadjun = PresuCam.OpenRecordset("DOCADJUN", dbOpenDynaset)
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
   If ESTABIER%(YABIER%, LUACCA$, "MAQUINARIAS") < 1 Then
     Call VERISTRU("MAQUINARIAS", YAINDEX%)
     Set Maquinarias = OpenDatabase(LUACREPLA$("MAQUINARIAS") + "MAQUINARIAS.MDB")
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
   If ESTABIER%(YABIER%, LUACCA$, "OMANTIN") < 1 Then
     Call VERISTRU("OMANTIN", YAINDEX%)
     Set Omantin = OpenDatabase(LUACREPLA$("OMANTIN") + "OMANTIN.MDB")
   End If
   '
90 Set OrdeMant = Omantin.OpenRecordset("ORDEMANT", dbOpenDynaset)
   Set RecuProp = Omantin.OpenRecordset("RECUPROP", dbOpenDynaset)
   Set SeTaReco = Omantin.OpenRecordset("SETARECO", dbOpenDynaset)
   Set ORecorri = Omantin.OpenRecordset("ORECORRI", dbOpenDynaset)
   YABIER% = 1
   Screen.MousePointer = SCMP
   '
End Sub
'OT-05-0856-WAR-FIN///

Function LUACREPLA$(BASDATOS$)
   '
   Static BADAT$(512), RUTBA$(512), YAPAR%(512), CANBA%, CTRX1%
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "LUACREPLA"
   '
   If CTRX1% < 1 Then
      TTXX$ = LOADFILE$(LUDAT$ + "RUTBASES.DRV")
      CANBA% = 0
      For KKI% = 1 To Len(TTXX$) Step 128
        TTYY$ = DESENCRI$(Mid$(TTXX$, KKI%, 128))
        If TRIM$(Mid$(TTYY$, 17, 112)) <> "" Then
           CANBA% = CANBA% + 1
           BADAT$(CANBA%) = UCase$(TRIM$(Left$(TTYY$, 16)))
           RUTBA$(CANBA%) = TRIM$(Mid$(TTYY$, 17, 112))
           If Right$(RUTBA$(CANBA%), 1) <> "\" Then
             RUTBA$(CANBA%) = RUTBA$(CANBA%) + "\"
           End If
        End If
      Next KKI%
      CTRX1% = 1
   End If
   '
   LUACREPLA$ = LUACC$
   For KKJ% = 1 To CANBA%
     If UCase$(BASDATOS$) = BADAT$(KKJ%) Then
        LUACREPLA$ = RUTBA$(KKJ%)
        ' RUTIACT$ = RUTIANT$
        GoTo 20      'Exit Function
     End If
   Next KKJ%
   KKJ% = 0
   '
20 RUTIACT$ = RUTIANT$
   '
   If KKJ% > 0 Then
      If YAPAR%(KKJ%) < 1 Then
        If LUACREPLA$ <> LUACC$ Then
          Call APAREDAT(BASDATOS$, LUACREPLA$, LUACC$, YAPAR%(KKJ%))
        End If
      End If
   End If
   '
End Function

Sub APAREDAT(BADAT$, LUREP$, LUACC$, YAPAR%)
   YAPAR% = 1
   PPXX1% = InStr(UCase$(LUACC$), "\BDACCESS")
   PPXX2% = InStr(UCase$(LUREP$), "\BDACCESS")
   If PPXX1% > 1 Then
     If PPXX2% > 1 Then
       LDX1$ = Left$(LUACC$, PPXX1%)
       LDX2$ = Left$(LUREP$, PPXX2%)
       If LDX1$ <> LDX2$ Then
         BDX$ = TRIM$(UCase$(BADAT$))
         Select Case BDX$
            Case "ARTICULOS"
               ARAP$ = "MASTER": GoSub 60
               ARAP$ = "ECOMAST": GoSub 60
         End Select
       End If
     End If
   End If
   Exit Sub
   '
60 FDT1# = FILEDATE#(LDX1$ + ARAP$ + ".DAT")
   FDT2# = FILEDATE#(LDX2$ + ARAP$ + ".DAT")
   If Abs(FDT1# - FDT2#) > 0.00001 Then
      If EXISTE%(LDX2$ + ARAP$ + ".DAT") Then
         YAPAR% = 0
         Call CIERRARCH(ARAP$)
         On Error Resume Next
         Call FileCopy(LDX2$ + ARAP$ + ".DAT", LDX1$ + ARAP$ + ".DAT")
         On Error GoTo 0
      End If
   End If
Return
'
End Sub
'
Sub CREACAMPO(BASEDA$, TABLAX$, NCAMPO$, TCAMPO, Optional LCAMPO)
   '
   Static CTXX%
   'Static wrkPredeterminado As Workspace
   Static BASDAT, CREABA, CREATA, CREACA
   RUTIANT$ = RUTIACT$
   RUTIACT$ = "CREACAMPO"
   '
'   If CTXX% < 1 Then
'     Set wrkPredeterminado = DBEngine.Workspaces(0)
'     CTXX% = 1
'   End If
   '
   ' crea la base de datos BASEDA$
   On Error Resume Next
  'ESTA ES LA ORIGINAL PERO NO ME ESTA FUNCIONANDO.
   Set CREABA = DBEngine.CreateDatabase(LUACREPLA$(BASEDA$) + BASEDA$ + ".MDB", dbLangGeneral)
'   Set CREABA = wrkPredeterminado.CreateDatabase(LUACREPLA$(BASEDA$) + BASEDA$ + ".MDB", dbLangGeneral)
   If Err < 1 Or Err = 3204 Then     ' LA BASE DE DATOS YA EXISTE
       GoTo 10
     Else
       MsgBox "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Base de Datos " + LUACREPLA$(BASEDA$) + BASEDA$
       Close: End
   End If
   '
10  On Error Resume Next
   ' crea la tabla TABLAX$
   Set BASDAT = OpenDatabase(LUACREPLA$(BASEDA$) + BASEDA$ + ".MDB")
   Set CREATA = BASDAT.CreateTableDef(TABLAX$)
   CREATA.Fields.Append CREATA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
   BASDAT.TableDefs.Append CREATA
   If Err < 1 Or Err = 3010 Then     ' LA tabla ya existe
       GoTo 20
     Else
       MsgBox "Error (" + TRIM$(Str$(Err)) + ") - Imposible Crear Tabla '" + TABLAX$ + "' en BD '" + BASEDA$ + "'."
       Close: End
   End If
   '
20 Set CREACA = BASDAT.TableDefs(TABLAX$)
   ' Verifica pre-existencia del campo en la tabla
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
           If Left$(UCase$(App.EXEName), 5) = "WINTA" Then GoTo 90
           USLOCK$ = TRIM$(Left$(LOADFILE$(LUACREPLA$(BASEDA$) + BASEDA$ + ".LDB"), 24))
           TTXX$ = "Imposible Crear Campo '" + NCAMPO$ + "'\en Tabla '" + TABLAX$ + "'\en Base de Datos '" + BASEDA$ + ".MDB'."
           TTXX$ = TTXX$ + "\   \Error " + TRIM$(Str$(ERREX&)) + "."
           TTXX$ = TTXX$ + "\   \La Base de Datos está en Uso Exclusivo\en la Estación de Trabajo '" + USLOCK$ + "'."
           TTXX$ = TTXX$ + "\   \Cierre FLEXOFT en esa Terminal y Re-Intente."
           HUBOERROR% = 1
           Call MENSERR(24, TTXX$)
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
           CREACA.Fields(NCAMPO$).DefaultValue = CVDAT(0)
        Else
           CREACA.Fields(NCAMPO$).DefaultValue = 0
     End If
     On Error GoTo 0
     '
   End With
   '
90 RUTIACT$ = RUTIANT$
End Sub
'
Function VALOBADAMDB(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   
   sqlx$ = "SELECT " + NOCAM$ + " AS VALORSE "
   sqlx$ = sqlx$ + " FROM " + NOTAB$
   sqlx$ = sqlx$ + " WHERE " + CONDI$
   '
   VALOBADAMDB = ""
   On Error Resume Next
   If TRIM$(UCase$(NOTAB$)) = "CAJABANC" Or TRIM$(UCase$(NOTAB$)) = "SADEUPEN" Or TRIM$(UCase$(NOTAB$)) = "APLICOBRA" Then
     Set MIRS = CueCorri.OpenRecordset(sqlx$, dbOpenSnapshot)
   ElseIf TRIM$(UCase$(NOTAB$)) = "MOVISTO" Then
     Set MIRS = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
   ElseIf TRIM$(UCase$(NOTAB$)) = "PEDENCA" Then
     Set MIRS = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
   End If
   If Err Then
     On Error GoTo 0
     If MODOMESS$ <> "NOMESS" Then
        TTXX$ = "Imposible Ejecutar Consulta\   \'" + sqlx$ + "'"
        If COMALTER%(TTXX$ + "\\Continuar\Interrumpir") = 2 Then
           Call FINAL("")
        End If
     End If
     GoTo 90
   End If
   On Error GoTo 0
   With MIRS
     If Not (.BOF And .EOF) Then
       VTP% = VarType(!VALORSE)
       If VTP% = 8 Or VTP% < 2 Then
           VALOBADAMDB = SAFETEXT(!VALORSE)
         Else
           VALOBADAMDB = !VALORSE
       End If
     End If
   End With
   '
90 On Error Resume Next
   MIRS.Close
   Set MIRS = Nothing
   On Error GoTo 0
End Function
'
Sub GENERBASE(NOMBASE$)
   '
   If TRIM$(UCase$(NOMBASE$)) = "ARTICULOS" Then Call CREATABLA_ARTICULOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BADACLIEN" Then Call CREATABLA_BADACLIEN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BADAPROVE" Then Call CREATABLA_BADAPROVE: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BCONTROL" Then Call CREATABLA_BCONTROL: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "COMPRAS" Then Call CREATABLA_COMPRAS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "CONTAGEN" Then Call CREATABLA_CONTAGEN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "CUECORRI" Then Call CREATABLA_CUECORRI: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "EMPLEADOS" Then Call CREATABLA_EMPLEADOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "EQUIPOS" Then Call CREATABLA_EQUIPOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "MAQUINARIAS" Then Call CREATABLA_MAQUINARIAS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "MENSAFLEX" Then Call CREATABLA_MENSAFLEX: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OFABRIC" Then Call CREATABLA_OFABRIC: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OMANTIN" Then Call CREATABLA_OMANTIN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OTRABAJO" Then Call CREATABLA_OTRABAJO: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "PRESUCAM" Then Call CREATABLA_PRESUCAM: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKORIGI" Then Call CREATABLA_STOCKORIGI: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKS" Then Call CREATABLA_STOCKS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKVIE" Then Call CREATABLA_STOCKVIE: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "VENTAS" Then Call CREATABLA_VENTAS: Exit Sub
   '
End Sub
'
Sub CREATABLA_ARTICULOS()
   '
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Cod_Mod", 10, 16)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Ubic", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "CantPiezas", 6, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Frecuen", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "FechUlt_Con", 8, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "FechPro_Con", 8, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Material", 3, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Desc_Mat", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Cod_Inte", 3, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "AnotaMod", 12, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "ObservaMod", 12, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto1", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto2", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto3", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto4", 10, 128)
   '
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Cod_Inte", 3, 0)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Descripcion", 10, 24)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Ruta", 10, 128)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "Formulas", "CodIConj", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXConj", 10, 16)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodiElem", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXElem", 10, 16)
   Call CREACAMPO("ARTICULOS", "Formulas", "NuorIt", 3, 2)
   Call CREACAMPO("ARTICULOS", "Formulas", "UsoBruto", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Unis_Form", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Lar_Corte", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Pladispo", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "RefeComp", 12, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodIConj", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodxConj", 10, 16)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodiElem", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodXElem", 10, 16)
   Call CREACAMPO("ARTICULOS", "FormuPat", "NuorIt", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "UsoBruto", 6, 4)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Operacion", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Marborra", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Min", 6, 0)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Max", 6, 0)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Prec", 6, 0)
   '
   Call CREACAMPO("ARTICULOS", "GRAFICOS", "Codigo", 3, 2)
   Call CREACAMPO("ARTICULOS", "GRAFICOS", "Nume_Ord", 3, 2)
   '
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Codigo", 10, 16)
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Codint", 3, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cod_Oper", 3, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cod_Herramental", 10, 16)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cantidad", 6, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Marborra", 3, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Tipo_Herr", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "Master", "Ferevis", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Observa", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Especificacion", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DesCompAut", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DescriComp", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "DepoEntra", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DepoCon", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Descridos", 10, 128)
   Call CREACAMPO("ARTICULOS", "Master", "Status", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codint", 4, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Alias", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Umedida", 10, 4)
   Call CREACAMPO("ARTICULOS", "Master", "Timate", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Critrep", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Monecos", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Fecosto", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PorDescCos", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Cosuni", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Pornacio", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PorComVen", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Provhabi", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stomin", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stomax", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Lorepos", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Lindivis", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Pesuni", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PesMetro", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Color", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Talle", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Descompo", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Grupiva", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Grucoven", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CLinea", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "Cfamilia", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "NPlano", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Revision", 10, 16)
   Call CREACAMPO("ARTICULOS", "Master", "Trazable", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Revi_Formu", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FechApFor", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "UsuApFor", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Cod_Interna", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Bien_Cap", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo_Ncm", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo_Sec", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Pro_Umed", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Iva_Id", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "ArchiPdf", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Familia", 10, 4)
   Call CREACAMPO("ARTICULOS", "Master", "Lindiven", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CanDeci", 10, 2)
   Call CREACAMPO("ARTICULOS", "Master", "ConsuProm", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CosCal", 7, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeCosCal", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stat_Movi", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Clas_ATC", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Vers_BC", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "FeVig_BC", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeRev_BC", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CantApro", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "LeTiCom", 10, 16)
   Call CREACAMPO("ARTICULOS", "Master", "PiezasXCaja", 4, 0)
   Call CREACAMPO("ARTICULOS", "Master", "M3Caja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "KiloNetoCaja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "KiloBrutoCaja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Autono_Proy", 10, 10)
   Call CREACAMPO("ARTICULOS", "Master", "UsuPrePes", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "UsuPreDol", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "MANO_OBRA_REAL", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Tamano_Etiqueta", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodIConj", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodXConj", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "NumeOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "AltOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "DescOper", 10, 64)
   Call CREACAMPO("ARTICULOS", "SECOPER", "StatOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq0", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq1", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq2", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq3", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq4", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq5", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsPrep", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsFijo", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "MinuStan", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsMHer", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsLimp", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PieCiclo", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PorDemor", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Cadencia", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CoefSolp", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "LoteStan", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CantiOps", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Cantiops2", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CategOps", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Categops2", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProgCNum", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HoProFab", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ObseOper", 12, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProTerOp", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProvOper", 10, 12)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PreUnid", 7, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "MonePrec", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "FechPrec", 8, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Nuoper_Prep", 6, 4)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Catoper_Prep", 3, 2)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Fech_Revis", 8, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Usua_Revis", 10, 32)
   '
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_Inte", 3, 0)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_Oper", 3, 0)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_TMU", 10, 4)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Nume_Ord", 3, 2)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cantidad", 7, 8)
   '
End Sub
'
Sub CREATABLA_BADACLIEN()
   '
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Codi_Cli", 10, 12)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Raso_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Origen", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "RamoActi", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "ProduPrin", 10, 128)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Empleados", 4, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Tele_Cli2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Fax_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo1", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector1", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno1", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Contacto2", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector2", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Contacto3", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo3", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector3", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno3", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "PagWeb", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechBaja", 10, 8)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Mot_Baja", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechUltEnvio", 8, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Grupos", 10, 30)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechProxCto", 8, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "His_Contac", 12, 0)
   '
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre0_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre1_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "NTrans_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre2_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Ent", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Rep", 10, 8)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Gru_Div", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nume_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Codi_Cli", 10, 12)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Raso_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Fant_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Stat_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Faxi_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "PagWeb", 10, 128)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Piva_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cuit_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nibr_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Prvd_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cpag_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lpre_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tcpr_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CaIb_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Dsta_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lcre_Cli", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Sald_Cli", 7, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CuMa_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "GrCo_Cli", 3, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Obser_Entre", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Recar_Embal", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mar_Ya_Fac", 3, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DePosIva", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "LCred_CCte", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "EnvMaiDeuda", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Ctact_Cob", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tele_Cob", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mail_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasRec_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasPago_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Segui_Cob", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "MailFactura_Cob", 10, 64)
   '
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Suc", 3, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Deno_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Domi_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Loca_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Tele_Suc", 10, 64)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Codi_Tra", 3, 0)
   '
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Codi_Tra", 3, 0)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Deno_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Raso_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Domi_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Loca_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Cuit_Tra", 10, 32)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Piva_Tra", 3, 0)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Tele_Tra", 10, 64)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Faxi_Tra", 10, 64)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Mail_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Ctac_Tra", 10, 128)
   '
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "RasoVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "DomiVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "LocaVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_1", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_2", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "EmailVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Observaciones", 12, 0)
   '
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "CodTrabajo", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "CodexTRab", 10, 24)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "DescTrab", 10, 64)

   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "TipoTrabajo", 10, 64)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Dia", 10, 24)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "OrdenSemana", 10, 16)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Horario", 10, 32)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Continuidad", 3, 2)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Cuotas", 3, 48)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "MARBORRA", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "ORDEN", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "ID_DIA", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "APARTIRDE", 10, 16)
   '
   Call CREACAMPO("BADACLIEN", "Agenda", "CodTrabajo", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "TipoTrabajo", 10, 64)
   Call CREACAMPO("BADACLIEN", "Agenda", "CodexTRab", 10, 24)
   Call CREACAMPO("BADACLIEN", "Agenda", "DescTrab", 10, 64)
   Call CREACAMPO("BADACLIEN", "Agenda", "Fecha_Original", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "OrdenSemana", 10, 16)
   Call CREACAMPO("BADACLIEN", "Agenda", "Horario_Orig", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "Continuidad", 3, 2)
   Call CREACAMPO("BADACLIEN", "Agenda", "Cuotas", 3, 48)
   Call CREACAMPO("BADACLIEN", "Agenda", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "MARBORRA", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "ORDEN", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "Dia_Organizado", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "Horario_Organizado", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "Dia_Realizado", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "FACTURADO", 10, 8)
   Call CREACAMPO("BADACLIEN", "Agenda", "MESANO_ORIG", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "ID_ORIGINAL", 4, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "REALIZADO", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "OPER_REALIZO", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "PEDIDO", 4, 0)
   
End Sub
'
Sub CREATABLA_BADAPROVE()
   '
   Call CREACAMPO("BADAPROVE", "Proved12", "Lcre_Cli", 6, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Sald_Cli", 7, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CuMa_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "GrCo_Cli", 3, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "C_Ret", 10, 10)
   Call CREACAMPO("BADAPROVE", "Proved12", "Posivat_Cli", 10, 10)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Codi_Cli", 10, 12)
   Call CREACAMPO("BADAPROVE", "Proved12", "Raso_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Fant_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Stat_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADAPROVE", "Proved12", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tele2_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Faxi_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Piva_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cuit_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nibr_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Prvd_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Hora_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Ord_Cheq", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cpag_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CCom_Cli", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Lpre_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tcpr_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CaIb_Cli", 10, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Dsta_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Calinteg", 6, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "ObsCompras", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "LuEntCompras", 12, 0)
   '
End Sub
'
Sub CREATABLA_BCONTROL()
   '
   Call CREACAMPO("BCONTROL", "AUDITRA", "FeReTra", 8, 8)
   Call CREACAMPO("BCONTROL", "AUDITRA", "TipoTra", 10, 8)
   Call CREACAMPO("BCONTROL", "AUDITRA", "DescriTra", 10, 64)
   Call CREACAMPO("BCONTROL", "AUDITRA", "NuCompro", 10, 24)
   Call CREACAMPO("BCONTROL", "AUDITRA", "UserTra", 10, 16)
   Call CREACAMPO("BCONTROL", "AUDITRA", "IdTermi", 10, 24)
   '
   Call CREACAMPO("BCONTROL", "BloArchA", "IdenTabla", 10, 32)
   '
   Call CREACAMPO("BCONTROL", "BloArchB", "CtrlName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "ExecName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "TermName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "IdenTabla", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "InitDate", 8, 0)
   Call CREACAMPO("BCONTROL", "BloArchB", "UserName", 10, 32)
   '
   Call CREACAMPO("BCONTROL", "MENUES", "CodPadre", 10, 32)
   Call CREACAMPO("BCONTROL", "MENUES", "Relacion", 4, 4)
   Call CREACAMPO("BCONTROL", "MENUES", "CodHijo", 10, 128)
   Call CREACAMPO("BCONTROL", "MENUES", "Descripcion", 10, 64)
   '
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "NomLiSele", 10, 32)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "ACol_1", 3, 2)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "ACol_2", 3, 2)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Tabla_Origen", 10, 256)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Lista_Campos", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Condi_Selec", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Fech_Actu", 8, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "LiSelec", 12, 0)
   '
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "Cod_Query", 10, 16)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "Des_Query", 10, 256)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "PAR_Query", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "SQL_Query", 12, 0)
   '
End Sub
'
Sub CREATABLA_COMPRAS()
   '
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Aten", 4, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Femi_Aten", 8, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Npro_Aten", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Raso_Aten", 10, 64)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cant_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Acuen_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Sald_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Stat_Aten", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Observa_Aten", 10, 64)
   '
   Call CREACAMPO("COMPRAS", "COTIRECI", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "NuOrd_item", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Npro_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cod_Exte", 10, 24)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Descritem", 12, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cant_Solint", 6, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Uni_Sol", 10, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Fech_Coti", 8, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Mone_Emis", 3, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "COTIRECI", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Plaz_Entre", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Venci_Ofer", 8, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Observa", 12, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "CondPag", 3, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Sect_Solint", 10, 12)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Pre_Unit_Pes", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "DETAIVA", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Norit", 3, 2)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Tasa_Iva", 10, 16)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Base_Impo", 7, 8)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Impo_Iva", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Descritem", 12, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Sol", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cant_Solint", 6, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Coti_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Desti_Solint", 10, 255)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Iconta_Solint", 10, 32)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Status", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Delibre", 12, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "IContab", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Sect_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Colo_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Medi_Ocom", 10, 20)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Valo_Unid", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Unids_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Uni_Com", 10, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Coef_Unids", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Desc_Item", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "PrUn_Neto", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "FVig_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Consimat", 10, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cant_Rec", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Oper", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Doc_Orig", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "N_OrServ", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nro_Otrab", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nro_Grup", 10, 10)
   '
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Femi_Solint", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_1", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_2", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_3", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_4", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Obse_Solint", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Sect_Solint", 10, 12)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_1", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_2", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_3", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_4", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "AnotaSolint", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "HayNews", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Descri_Solint", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Status", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FVig_Ocom", 8, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FGen_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FApro_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Apro", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FMod_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Mod", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FAnu_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Anul", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Moti_Anul", 10, 16)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Tcam_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "ToBru_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Pdes_Ocom", 10, 24)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Ides_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iiva_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iotr1_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iotr2_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Total_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Cpag_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Embala_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "AtSr_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Obse_Ocom", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Lentrega", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Anex_Ocom", 10, 128)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "OcomRevis", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "MACONSIG", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "MACONSIG", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Secc_Ejec", 6, 4)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Deno_Secc", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Nro_Orfa", 10, 12)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Nro_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Lote_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Unidad", 10, 4)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Asig", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Entreg", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Fech_Entreg", 8, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "NRemi_Consig", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Valo_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Moneda", 3, 2)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Valo_Tot", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Rendid", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Saldo_Pend", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Observa", 12, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "NuOrdItem", 3, 0)
   '
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cantidad", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "FechRecep", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Nume_Ocom", 4, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Periodo", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cosuni", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "StokAct", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Stomin", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Lorepos", 6, 0)
   '
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Cli", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Delibre", 12, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "IContab", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Sect_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Colo_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Medi_Ocom", 10, 20)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Valo_Unid", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Unids_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Uni_Com", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Coef_Unids", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Desc_Item", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "PrUn_Neto", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Consimat", 10, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cant_Rec", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Oper", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Doc_Orig", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "N_OrServ", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nro_Otrab", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nro_Grup", 10, 10)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "IdenReque", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "PrUn_Neto_Pes", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "OCOMENCA", "AtSr_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Obse_Ocom", 12, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Lentrega", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Anex_Ocom", 10, 128)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "OcomRevis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Cli", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FGen_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FApro_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Apro", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FMod_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Mod", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FAnu_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Anul", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Moti_Anul", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Tcam_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "ToBru_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Pdes_Ocom", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Ides_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iiva_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iotr1_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iotr2_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Total_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Cpag_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Embala_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "PorcIva", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Id_Comprob", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "CodCarpImp", 10, 12)
   '
   Call CREACAMPO("COMPRAS", "PROVECOD", "Num_Prov", 4, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Cod_inte", 4, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Precio", 7, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Moneda", 3, 2)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Fecha", 8, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "MarBorra", 3, 0)
   '
End Sub
'
Sub CREATABLA_CONTAGEN()
   '
   Call CREACAMPO("CONTAGEN", "DetaAsien", "NuIdAsien", 10, 16)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "OrNuPase", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "RefePase", 10, 64)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "FechAsien", 8, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "OrFeAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "SucuAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "NuDfAsien", 4, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "StatPase", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CueCoInt", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CueCoExt", 10, 16)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "IDebePase", 5, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "IHabePase", 5, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "MarConsi", 10, 8)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CenCos", 10, 8)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "Niv_Clasi", 3, 2)
   '
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "NuIdAsien", 10, 16)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "RefeAsien", 10, 64)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "FechAsien", 8, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "OrFeAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "SucuAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "NuDfAsien", 4, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "TipoAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "StatAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "EditAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "QPasAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "DebeAsien", 5, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Observa", 12, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Nume_Usuar", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Niv_Clasi", 3, 2)
   '
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CueContaI", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CenCosto", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Fecha", 8, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Concepto", 10, 128)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Ingreso", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Gastos", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Marborra", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CodCuenta", 10, 16)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "MesAno", 10, 24)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Tipo", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Imp_Real", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Ing_Gas", 10, 8)
   '
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "CodExCue", 10, 16)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Denocue", 10, 120)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "CodInCue", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Statcue", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "NivAgru", 3, 4)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "TipoCue", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "NuOrdIt", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "PosImpu", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Observa", 12, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Salini", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Sudebe", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Suhabe", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Result", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Salfin", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "IndiAjus", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "MoneExtr", 3, 0)
   '
   Call CREACAMPO("CONTAGEN", "TABEJCON", "CodiEjer", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "DenoEjCon", 10, 48)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechAper", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechCierr", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Status", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Diverif", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "MarcActi", 3, 0)
   '
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Cod_Proy", 10, 8)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Deno_Proy", 10, 48)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Clien_Proy", 3, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "FeAlt_Proy", 8, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "FeCie_Proy", 8, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Stat_Proy", 3, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Descrip_Proy", 12, 0)
   '
End Sub
'
Sub CREATABLA_CUECORRI()
   '
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "NuClien", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiClie", 10, 16)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "RasoClie", 10, 48)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocAp_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocAp_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ImpoApli", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "RefeApli", 10, 32)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocOr_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocOr_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ItotOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "IIvaOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "Mone_Apli", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ItotOri_Dol", 7, 15)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ImpoApli_Dol", 7, 15)
   '
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "NuProve", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiProv", 10, 16)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "RasoProv", 10, 48)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocAp_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocAp_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ImpoApli", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "RefeApli", 10, 32)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocOr_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocOr_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ItotOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "IIvaOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "Monemi_Ori", 3, 0)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ImpoApli_Ori", 7, 8)
   '
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenAlta", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenBaja", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechEntre", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeHoReal", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuAsiDiar", 10, 16)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeDeposEnt", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DocumEnt", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DestiCheq", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "RefeSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "UsuaSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "PerLiComis", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoComis", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "VendeComis", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiCom_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiCom_Lar", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "RefeCom", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechConta", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechAcred", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "MarConsi", 10, 4)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CueContaI", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CenCosto", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ItemConc", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuClien", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Sucu_Clie", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuProve", 3, 3)
   Call CREACAMPO("CUECORRI", "CAJABANC", "TipoMovim", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "OpciMovim", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "VariMovim", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DescriMov", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiEmpr", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiSucu", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeTalo", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeComp", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Status", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Mone_Emis", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Tipo_Cam", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Impo_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ValAbsComp", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IDebeComp", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IHabeComp", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "BaseImpo", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IncSubIva", 10, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NroCai", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeVtoCai", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ObserGen", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "AtraCheque", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "AtraCance", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CondPag", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodRetGan", 10, 3)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodPerGan", 10, 3)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Niv_Clasi", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "LetraComis", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Vend_Cli", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdCae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodCarpImp", 10, 12)
   Call CREACAMPO("CUECORRI", "CAJABANC", "TextoCae", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Notas", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Incoterms", 10, 12)
   Call CREACAMPO("CUECORRI", "CAJABANC", "OCompra", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Pedidos", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cli_Real", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechVtoCae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "SimMone", 10, 16)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cod_Barra", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Descu1", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Descu2", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoDesc1", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoDesc2", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CUIT", 10, 13)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DespaImpo", 10, 24)
   '
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_OrIt", 4, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_CueImpu", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Cli", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_CueImpu", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "Compro", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Comp", 4, 0)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Cli", 4, 0)
   '
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Fecha", 8, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Referencia", 10, 64)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Ingresos", 7, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Gastos", 7, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Tipo", 3, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "MarBorra", 3, 0)
   '
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "NuClien", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "NuProve", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodCliPro", 10, 16)
   Call CREACAMPO("CUECORRI", "SACREPEN", "RaSoCliPro", 10, 48)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiCom_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiCom_Lar", 10, 24)
   Call CREACAMPO("CUECORRI", "SACREPEN", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Mone_Emis", 3, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Ibru_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldDebe_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldAcre_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "IBruComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "IIvaComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "OrVencim", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "FeVencim", 8, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldDebe", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldAcre", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Status", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Niv_Clasi", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Observa", 12, 0)
   '
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "NuClien", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Sucu_Clie", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "NuProve", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodCliPro", 10, 16)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "RaSoCliPro", 10, 48)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiCom_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiCom_Lar", 10, 24)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Mone_Emis", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Ibru_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldDebe_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldAcre_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "IbruComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "IIvaComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "OrVencim", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "FeVencim", 8, 0)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldDebe", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldAcre", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Status", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CondPag", 3, 0)
   '
End Sub
'
Sub CREATABLA_EMPLEADOS()
   '
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "Legajo", 3, 2)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "Competencia", 10, 8)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "FechAlta", 8, 0)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "FechVenc", 8, 0)
   '
   Call CREACAMPO("EMPLEADOS", "PADRON", "Legajo", 3, 2)
   Call CREACAMPO("EMPLEADOS", "PADRON", "ID_SYS", 10, 16)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Nombre", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Grupo", 3, 2)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Status", 3, 2)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Documento", 10, 13)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Sexo", 10, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaNacim", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Nacionalidad", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "EstadoCivil", 10, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Domicilio", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CodPostal", 10, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Localidad", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Sector", 6, 4)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Tarea", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Categoria", 10, 9)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Horario", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaIng", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CUIL", 10, 13)
   Call CREACAMPO("EMPLEADOS", "PADRON", "AFJP", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CajaJub", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaEgre", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "MotivoEgre", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Foto", 10, 100)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Anotaciones", 12, 100)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaAntig", 8, 10)
   '
End Sub
'
Sub CREATABLA_EQUIPOS()
   '
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroOrden", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Resultado", 10, 4)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Origen", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumControl", 4, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Tec_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aext_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aint_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Obse_Ensa", 12, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Dic_Ensa", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR070", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto5", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto6", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Instrumento", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR072", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto5", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto6", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Corrfuga7", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto7", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga8", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto8", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Tension", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Instrumento", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieInstru", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieTension", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieTiempo", 10, 16)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieVeredicto", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieObserva", 12, 0)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR074", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO1", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO2", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "CHN", 10, 6)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedA", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedB", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedC", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Origen", 10, 16)
   '
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeroSerie", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Mesanio", 10, 4)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeOrd", 3, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Exte", 10, 24)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Descrip", 10, 64)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "IdSubOr", 10, 12)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechAlta", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechEntrega", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumRemito", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumFactura", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeCli", 4, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Paciente", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Ten_Ensa", 3, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Dic_Ensa", 10, 4)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Tec_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Aext_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Obse_Ensa", 12, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Aint_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "DepoNumse", 3, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Historial", 12, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FeVenGar", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FeVenGaRep", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "UNumRep", 4, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Ubicacion", 10, 128)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO1", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO2", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO3", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO4", 10, 32)
   '
End Sub
'
Sub CREATABLA_MAQUINARIAS()
   '
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Num_Con", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Referen_Con", 10, 32)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Freq_Con", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FUL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Val_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FPL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Prom_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Observa_Con", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Codigo_Equ", 10, 6)
   '
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "CodLar_Equ", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Descri_Equ", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Marca_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Mode_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "NuSeRi_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Fabri_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Represe_Equ", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Ubica_Equ", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Alimenta_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Potenci_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Kilos_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Largo_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Ancho_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Alto_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Linea_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Posi_Equ", 10, 12)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Planta_Equ", 10, 4)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Finst_Equ", 8, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Fgara_Equ", 8, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Observa_Equ", 12, 0)
   '
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Norit_Imag", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Ruta_Imag", 12, 0)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Refe_Imag", 12, 0)
   '
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "CodLar_Equ", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Nro_Sub", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Orden_Tar", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Descri_Tar", 10, 128)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Frecu_Tar", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Uni_Frecu", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Obse_Tar", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodInte_Rep", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodExte_Rep", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Descrip_Rep", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "UM_Rep", 10, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Cant_Rep", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Observa_Rep", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Nro_Sub", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Descri_Sub", 10, 32)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Marca_Sub", 10, 24)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Modelo_Sub", 10, 20)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "NSerie_Sub", 10, 20)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Criticidad", 10, 2)
   '
End Sub
'
Sub CREATABLA_MENSAFLEX()
   '
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "NumMensa", 4, 0)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaViejo", 12, 0)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaEsp", 12, 0)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaPor", 12, 0)
   '
End Sub
'
Sub CREATABLA_OFABRIC()
   '
   Call CREACAMPO("OFABRIC", "DOCANRE", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "DOCANRE", "NomAsignado", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_NAME", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_DATA", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "GALIMREP", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Descripcion", 12, 0)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Norit_Imag", 3, 2)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Nom_Imag", 10, 256)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Imagen", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescriOF", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NuOrInf", 3, 2)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Inte", 3, 2)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NumeOper", 3, 3)
   Call CREACAMPO("OFABRIC", "INFOFAB", "AltOper", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "INFOFAB", "ReTraba", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "AtMulti", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CanOrPro", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CaNoConf", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "LegaOps", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FeHorLib", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Des", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Has", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Fri", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MaqLiber", 10, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemStan", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Rendim", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Observa", 10, 30)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Referen", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MotiRecha", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NroInfNC", 10, 32)
   Call CREACAMPO("OFABRIC", "INFOFAB", "RefeInf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "UserControl", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Operarios", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdInfofab", 10, 24)
   '
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nuor_Inf", 4, 4)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nume_Inf", 10, 12)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Fecha_Inf", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Refer_Inf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Cant_Inf", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Proy_Inf", 10, 8)
   Call CREACAMPO("OFABRIC", "INFOPROY", "LegaOps_Inf", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Total_Inf", 7, 8)
   '
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cod_Inte", 3, 2)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cant_Reque", 7, 4)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Fech_Nec", 8, 0)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Destino", 10, 32)
   '
   Call CREACAMPO("OFABRIC", "OPERFAB", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "OPERFAB", "NumeOper", 3, 2)
   Call CREACAMPO("OFABRIC", "OPERFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "OPERFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "OPERFAB", "EquiAux", 10, 18)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantiOps", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantProg", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantPend", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemPrep", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemFijo", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemVari", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemDemo", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemMHer", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemLimp", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemOtro", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemTota", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemPend", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "StatOper", 3, 2)
   Call CREACAMPO("OFABRIC", "OPERFAB", "ComiePro", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TermiPro", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TermiTemp", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Observa", 10, 48)
   Call CREACAMPO("OFABRIC", "OPERFAB", "HoBatchRec", 11, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "SecciPro", 6, 4)
   '
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumeOrFa", 4, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechGen", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Unidad", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanProy", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Destino", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Priorid", 3, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "StatuOrf", 3, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NuOrProd", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechLan", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "SecciPro", 6, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechCierr", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanApro", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanRech", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdenRetra", 10, 12)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechProyEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechTempEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechRecal", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Observa", 12, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Origen", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "AnotaOrf", 12, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumRem_Prov", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Nume_Ocom", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Nume_Cli", 4, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Id_Orfa", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NroPed", 4, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Armador ", 3, 0)
   '
   Call CREACAMPO("OFABRIC", "RECONCILIA", "IdSubOr", 10, 16)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "IdenLote", 10, 16)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Consumido", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Mermas", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Devolucion", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "ConsuNeto", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "ConsuFormu", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Diferencia", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Autorizado", 10, 64)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Fecha", 8, 0)
   '
   Call CREACAMPO("OFABRIC", "REPARA", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Referen", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "FechRecep", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Nume_Cli", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Raso_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Domi_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Loca_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Pais_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Tele_Cli", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Mail_Cli", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Descripcion", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "NumeroSerie", 10, 32)
   Call CREACAMPO("OFABRIC", "REPARA", "Accesorios", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Defec_Infor", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Observa", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Trab_Efect", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "PresuEsti", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Usuario_Repu", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Estima_Horas", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Or_Compra", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "FOr_Compra", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "NumRemito", 10, 18)
   Call CREACAMPO("OFABRIC", "REPARA", "FechEnt", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Nom_Tra", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Cuit_Cli", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Piva_Cli", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Conductor", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Patente", 10, 16)
   Call CREACAMPO("OFABRIC", "REPARA", "Lentre_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Status_Gar", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "CausaFalla", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "DeCauFalla", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "TipoFalla", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "DeTipFalla", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "Referen2", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "HoraMObra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Mobra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "ValMObra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Repu", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Tot", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "FechPres", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "FechApro", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "NroPed", 3, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "AnotaRep", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "CodRepa", 10, 16)
   Call CREACAMPO("OFABRIC", "REPARA", "NuOrdRep", 4, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "RefOrdRep", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "FechOrdRep", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Inspeccion", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Reparacion", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Sitio", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Factura", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Importe", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Tot_Repuestos", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Tot_Mano_Obra", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "UsRealizaRepa", 10, 32)
   '
   Call CREACAMPO("OFABRIC", "REPUREPA", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Cant_Sol", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "PrUn_Neto", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "MarBorra", 3, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Impo_Neto", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Nuord_Item", 3, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Origen", 3, 0)
   '
   Call CREACAMPO("OFABRIC", "RESERVASF", "Cod_Inte", 3, 2)
   '
End Sub
'
Sub CREATABLA_OMANTIN()
   '
   Call CREACAMPO("OMANTIN", "ORDEMANT", "NumeOrMa_OM", 4, 4)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Referen_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FechGen_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Codigo_Equ", 10, 6)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Descri_OM", 12, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Status_OM", 3, 1)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FechCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoraCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "UsuGen_OM", 10, 25)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "UsuCie_OM", 10, 25)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Proved_OM", 3, 2)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "CostoTer_OM", 7, 8)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FeProCom_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoProCom_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FeProFin_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoProFin_OM", 8, 0)
   '
   Call CREACAMPO("OMANTIN", "ORECORRI", "NumORec", 4, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "DesORec", 10, 128)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechEmis", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechCierr", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseEmis", 12, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseCierr", 12, 0)
   '
   Call CREACAMPO("OMANTIN", "RECUPROP", "NumeOrMa_OM", 4, 4)
   Call CREACAMPO("OMANTIN", "RECUPROP", "Descri_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "RECUPROP", "LegaEmp_OM", 3, 2)
   Call CREACAMPO("OMANTIN", "RECUPROP", "Empleado_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "RECUPROP", "FeCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "RECUPROP", "CaHora_OM", 6, 4)
   '
   Call CREACAMPO("OMANTIN", "SETARECO", "Cod_Reco", 10, 8)
   Call CREACAMPO("OMANTIN", "SETARECO", "Des_Reco", 10, 72)
   Call CREACAMPO("OMANTIN", "SETARECO", "Norit_Tar", 3, 0)
   Call CREACAMPO("OMANTIN", "SETARECO", "MaEquip", 10, 8)
   Call CREACAMPO("OMANTIN", "SETARECO", "Des_Larga", 12, 0)
   '
End Sub
'
Sub CREATABLA_OTRABAJO()
   '
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Prog_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Comp_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Deno_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Observa_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla05_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla07_Grup", 10, 10)
   '
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "NuOrden", 3, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Cantid", 7, 8)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Descri_Cor", 10, 48)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "NPlano", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Posic", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Material", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "CodExte", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Kilos", 6, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Clave", 10, 8)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Descri_Lar", 12, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Ocom", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_BolRe", 10, 250)
   '
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Refe_Otrab", 10, 64)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Nclie_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Femi_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Fentre_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuConf_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuCont_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuApro_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NroOcom_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeOcom_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NroOfer_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeOfer_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NuOrItem_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "CanPed_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Especi_Otrab", 12, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "PreVen_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Mone_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "CotiMone_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "LEntrega_Otrab", 10, 255)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Ajusta_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Vended_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "ComiVend_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsInge_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsAvor_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCCal_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCald_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMeca_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoInt_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoExt_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsIngeInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsAvorInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCCalInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCaldInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMecaInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoIntInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoExtInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Observa_Otrab", 12, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeRev_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NuRev_Otrab", 10, 6)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "IContab", 3, 2)
   '
   Call CREACAMPO("OTRABAJO", "PLACOB", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "PLACOB", "NuOrItem", 3, 2)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Refe_Pcob", 10, 64)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Impo_Pcob", 7, 8)
   Call CREACAMPO("OTRABAJO", "PLACOB", "FeVen_Pcob", 8, 0)
   Call CREACAMPO("OTRABAJO", "PLACOB", "NroFac_Pcob", 10, 16)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Moneda", 3, 2)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Coti_Pcob", 7, 8)
   '
End Sub
'
Sub CREATABLA_PRESUCAM()
   '
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Num_Orden", 3, 2)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Referen", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Descrip", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Importe", 7, 8)
   '
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Rubro_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "SubRub_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "NuorItem", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "CodInte_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "CodExte_Det", 10, 16)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Cantid_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "PreUni_Det", 7, 8)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Mone_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med1_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med2_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med3_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med4_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Impo_Det", 7, 8)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Descri_Det", 10, 128)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Hora_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Opera_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "ValoHor_Det", 7, 8)
   '
   Call CREACAMPO("PRESUCAM", "DOCADJUN", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DOCADJUN", "RutArch", 12, 0)
   '
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Observa_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FeRev_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NuRev_Presu", 10, 6)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "IContab", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Largo_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ForPag_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PlaEnt_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Precios_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ValOfer_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "BoniVars_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FormElec_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "HorasCons_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ImporCons_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Ancho_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Alto_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Espesor_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "GGMat_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "GGOba_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "DesBon_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Refe_Presu", 10, 64)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nclie_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Femi_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FVali_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuConf_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuCont_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuApro_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NroOcom_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FeOcom_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PediCoti_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FePedCot_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NuOrItem_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "CanPed_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Especi_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PreVen_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Mone_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "CotiMone_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "LEntrega_Presu", 10, 255)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Ajusta_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Vended_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ComiVend_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Anti_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Cuot_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Inter_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Plazo_Entre", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Plazo_Pago", 3, 2)
   '
End Sub
'
Sub CREATABLA_STOCKORIGI()
   '
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Com", 10, 10)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Com", 4, 0)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Tmv", 10, 10)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Fecha_Mov", 8, 0)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Sig", 10, 1)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Dep", 10, 2)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Art", 10, 16)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Canti", 7, 8)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Trn", 4, 0)
   '
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Com", 10, 10)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "N_Com", 4, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Tmv", 10, 10)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "FechaMov", 8, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Ente", 10, 16)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Dep", 10, 2)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "N_Trn", 4, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "Descrip", 10, 64)
   '
End Sub
'
Sub CREATABLA_STOCKS()
   '
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodiCompro", 10, 16)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodSucu", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "IdenReque", 10, 24)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Cod_Inte", 3, 2)
   Call CREACAMPO("STOCKS", "BALAREQMO", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DoCompra", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuCert_Apro", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Posicion", 10, 24)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Reque", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Asig", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Saldo", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuLisPick", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DeLisPick", 10, 64)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Consu", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "MarqAsi", 3, 2)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CreadoPor", 10, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Destino", 10, 255)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Ident_Asig", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DepoMovi", 3, 1)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Nume_Cli", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Aped", 7, 8)
   '
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeBoRe", 6, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "IdenLote", 10, 12)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechRecep", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DescripItem", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Unidad", 10, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CantNom", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "M_Largo", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "M_Ancho", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DepoEnt", 3, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumProv", 3, 2)
   Call CREACAMPO("STOCKS", "BOLRECEP", "RaSo_Prov", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "LoteCol_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CerCal_Prov", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumRem_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumFac_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumConVig", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumOC", 4, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Nume_Aten", 4, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "IContab", 3, 2)
   Call CREACAMPO("STOCKS", "BOLRECEP", "MoneVaUni", 10, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Valo_Unit", 6, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToApro", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToRech", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToFalta", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechLibe", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeOrdSer", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Observa", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UsuApro", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FeHoApro", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Marca", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeParte", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechFab", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Plazo_Uso", 8, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Traz_Cli", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Cli", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Tipo", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Carac", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DepoUbic", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "AnotaOrf", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "MotiRecha", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "ObseRecha", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Destirecha", 3, 0)
   '
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Os", 10, 16)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Entre", 10, 16)
   Call CREACAMPO("STOCKS", "DETACARTEM", "FechEntre", 8, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "NorItCa", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Parte", 10, 32)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Inter", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Medio", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Pedid", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Recu", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Tota", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Enviar", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cant_Entre", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "SubCod", 10, 8)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Dopricor", 10, 12)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Doprilar", 10, 24)
   Call CREACAMPO("STOCKS", "DETACARTEM", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "DETACARTEM", "FechEmis", 8, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "IdCartem", 10, 48)
   '
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrdSe", 4, 4)
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrden", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "Modelo", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "NumeroSerie", 10, 32)
   Call CREACAMPO("STOCKS", "DETAOSER", "HorasFunc", 3, 2)
   '
   Call CREACAMPO("STOCKS", "ENCACARTEM", "FechEntre", 8, 0)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Os", 10, 16)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Entre", 10, 16)
   '
   Call CREACAMPO("STOCKS", "ENCAOSER", "NuOrdSe", 4, 4)
   Call CREACAMPO("STOCKS", "ENCAOSER", "FeOrdSe", 8, 0)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Nume_Cli", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "TipoOrdSe", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Observa", 12, 0)
   '
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Propiedad", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Concepto", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Producto", 12, 0)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Bultos", 3, 5)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Origen", 3, 3)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Deposito", 3, 3)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Nume_Clie", 3, 4)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Nume_Prov", 3, 4)
   Call CREACAMPO("STOCKS", "ENCAREMI", "CantPalet", 3, 4)
   '
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NumPallet", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NuOrdPallet", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "DesdeCaja", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "HastaCaja", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "CantiCajas", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "Codint", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "PiezasXCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "PiezasTotal", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "M3Caja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "M3Total", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloNetoCaja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloNetoTot", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloBrutoCaja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloBrutoTot", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "Marborra", 1, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NumLtaEmp", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NroPed", 4, 0)
   '
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cod_Inte", 3, 2)
   Call CREACAMPO("STOCKS", "LOTINGRE", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DocuIngre", 10, 24)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechAlta", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cantidad", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanAlta", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanConsu", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanReser", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalNeto", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "KilBruto", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "LOTINGRE", "AnchoMili", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "MetroIni", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalMetro", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechFab", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechVenc", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DepoUbic", 3, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Marca", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Origen", 10, 255)
   Call CREACAMPO("STOCKS", "LOTINGRE", "LoteCol_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Costo_Rep", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Moneco_Rep", 3, 2)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Fechco_Rep", 8, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "ValoHis", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "ValoRep", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Libera", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "TiMate", 3, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cod_Desti", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "MarInve", 3, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Potencia", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Tara", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Precinto", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalMeCua", 6, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Ref_Apro", 12, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "NuCert_Apro", 4, 4)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Usu_Apro", 10, 64)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Fech_Apro", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "OperProd", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "TurnoProd", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Nro_Analisis", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "EtiqxCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "RolxCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "EtiqxRollo", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CodCarpImp", 10, 12)
   '
   Call CREACAMPO("STOCKS", "MACONSIG", "Femi_Ocom", 8, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Npro_Ocom", 3, 2)
   Call CREACAMPO("STOCKS", "MACONSIG", "RaSo_Prov", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Secc_Ejec", 6, 4)
   Call CREACAMPO("STOCKS", "MACONSIG", "Deno_Secc", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Orfa", 10, 12)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Ocom", 10, 12)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MACONSIG", "Descritem", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Lote_Ocom", 10, 16)
   Call CREACAMPO("STOCKS", "MACONSIG", "Unidad", 10, 4)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Asig", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Entreg", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Fecha_Entreg", 8, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "NRemi_Consig", 10, 16)
   Call CREACAMPO("STOCKS", "MACONSIG", "Valo_Unit", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Moneda", 3, 2)
   Call CREACAMPO("STOCKS", "MACONSIG", "Valo_Tot", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Rendid", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Saldo_Pend", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Observa", 12, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Devu", 7, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "MarBorra", 3, 0)
   '
   Call CREACAMPO("STOCKS", "MASUMCLI", "Nume_Cli", 3, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "NumeIng", 4, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Yacimiento", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Referencia", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "TarAcordadas", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Observaciones", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "DepInterviniente", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechRecep", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_1", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_2", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_3", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_4", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "EstadoRec", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Ubicacion", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechVerif", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Trab_Efect", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechTrabEfect", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "EstEntrega", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechEnt", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Lentre_Cli", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Patente", 10, 16)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Conductor", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Piva_Cli", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Cuit_Cli", 10, 24)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Nom_Tra", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Fech_Remito", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "NumRemito", 10, 18)
   '
   Call CREACAMPO("STOCKS", "MOVIREPA", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIREPA", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CodiMovi", 10, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoSecCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NuPedCli", 4, 4)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DepoMovi", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValoUnitIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValTotItIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "FHor_Verif", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CodCarImp", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Causa", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Verificado", 3, 1)
   '
   Call CREACAMPO("STOCKS", "MOVISTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTO", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVISTO", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "CodiMovi", 10, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoSecCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVISTO", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "NuPedCli", 4, 4)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTO", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValoUnitIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValTotItIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVISTO", "DepoMovi", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Verificado", 3, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Causa", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "CodCarpImp", 10, 12)
   Call CREACAMPO("STOCKS", "MOVISTO", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "FHor_Verif", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "TipoTran", 3, 0)
   '
   Call CREACAMPO("STOCKS", "MOVIVIE", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIVIE", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIVIE", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CodiMovi", 10, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DosecCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DosecLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NuPedCli", 4, 4)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DepoMovi", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantSaldo", 7, 0)
   '
   Call CREACAMPO("STOCKS", "RESERVA", "Cod_Inte", 3, 2)
   Call CREACAMPO("STOCKS", "RESERVA", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKS", "RESERVA", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "RESERVA", "IdSubOr", 10, 12)
   Call CREACAMPO("STOCKS", "RESERVA", "NuOrItem", 3, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "FechRes", 8, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "CantRes", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "CantCons", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "Saldo_Entre", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "RESERVA", "Deposito", 3, 2)
   Call CREACAMPO("STOCKS", "RESERVA", "UbicaLot", 10, 32)
   Call CREACAMPO("STOCKS", "RESERVA", "UsoUnit", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "Observa", 10, 128)
   Call CREACAMPO("STOCKS", "RESERVA", "PNL_Des", 10, 24)
   Call CREACAMPO("STOCKS", "RESERVA", "PLC_Has", 10, 24)
   Call CREACAMPO("STOCKS", "RESERVA", "Nume_Solint", 4, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Nume_Ocom", 4, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Marborra", 3, 0)
   '
   Call CREACAMPO("STOCKS", "SALSTOCK", "CodiInte", 3, 2)
   Call CREACAMPO("STOCKS", "SALSTOCK", "VeSalSto", 11, 256)
   '
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodiCompro", 10, 5)
   Call CREACAMPO("STOCKS", "TRANMOVI", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodSucu", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Descridos", 10, 255)
   Call CREACAMPO("STOCKS", "TRANMOVI", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CanKilos", 7, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "TipoTrans", 3, 2)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CreadoPor", 10, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Destino", 10, 255)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Ingresado", 3, 0)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Ident_Asig", 10, 32)
   '
   Call CREACAMPO("STOCKS", "TRANRECEP", "CodiCompro", 10, 5)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "TRANRECEP", "DocuIngre", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "TRANRECEP", "CanKilos", 7, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NumRemProv", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechVenc", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechFab", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Origen", 10, 255)
   Call CREACAMPO("STOCKS", "TRANRECEP", "AproRech", 10, 1)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Ingresado", 3, 0)
   '
   Call CREACAMPO("STOCKS", "UBIPRE", "DepoUbic", 3, 0)
   Call CREACAMPO("STOCKS", "UBIPRE", "Codiubic", 10, 16)
   Call CREACAMPO("STOCKS", "UBIPRE", "Referencia", 10, 64)
   Call CREACAMPO("STOCKS", "UBIPRE", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "UBIPRE", "MarBorra", 3, 0)
   '
End Sub
'
Sub CREATABLA_STOCKVIE()
   '
   Call CREACAMPO("STOCKVIE", "MOVISTO", "FechMov", 8, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Descrip", 10, 64)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CodiMovi", 10, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoSecCor", 10, 12)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NuPedCli", 4, 4)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DepoMovi", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NumUsuar", 3, 0)
   '
End Sub
'
Sub CREATABLA_VENTAS()
   '
   Call CREACAMPO("VENTAS", "DESCOMBI", "Nume_Cli", 4, 0)
   Call CREACAMPO("VENTAS", "DESCOMBI", "GrCo_Cli", 3, 2)
   Call CREACAMPO("VENTAS", "DESCOMBI", "Cod_Inte", 3, 2)
   Call CREACAMPO("VENTAS", "DESCOMBI", "GruCoVen", 3, 0)
   Call CREACAMPO("VENTAS", "DESCOMBI", "PorDescu", 10, 12)
   Call CREACAMPO("VENTAS", "DESCOMBI", "Modo_Desc", 3, 0)
   '
   Call CREACAMPO("VENTAS", "DETCOTIZA", "NroCoti", 10, 16)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "ClaValor", 10, 48)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor1", 12, 0)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor2", 12, 0)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor3", 12, 0)
   '
   Call CREACAMPO("VENTAS", "LISTAPRE", "Cod_Lista", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Denom_Lista", 10, 64)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Vigencia", 8, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Formato", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Decimales", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Moneda", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Desc_Abrev", 10, 64)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Descu_lista", 6, 32)
   Call CREACAMPO("VENTAS", "LISTAPRE", "IncluIva", 3, 0)
   '
   Call CREACAMPO("VENTAS", "PEDDETA", "PreList", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanAsig", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanPed", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "PreUnid", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoTot", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoAnti", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "CantEnt", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanSaldo", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "FeUltEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FApro_Nope", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuPrepa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuDespa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroRemi", 10, 24)
   Call CREACAMPO("VENTAS", "PEDDETA", "TPedItem", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroPed", 4, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "TipoPed", 10, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroClie", 4, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "RaSoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroOcom", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "FeSolEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FinVigen", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Status", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "StaDespa", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "ListPre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "CondPag", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "PorComVend", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NuOrItem", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "CodiInt", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "CodiExt", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "Descrip", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "SubCod", 10, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "TColor", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "TTalle", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "TLargo", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "TAncho", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "TEspesor", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CBase", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "VenRep", 10, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "MoneEmis", 3, 2)
   Call CREACAMPO("VENTAS", "PEDDETA", "TipoCam", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescLista", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescCombi", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescRubro", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescTotal", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NumeroSerie", 10, 32)
   '
   Call CREACAMPO("VENTAS", "PEDENCA", "NroPed", 4, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Referencia", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "TipoPed", 10, 8)
   Call CREACAMPO("VENTAS", "PEDENCA", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "NroClie", 4, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "RasoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "NroOcom", 10, 16)
   Call CREACAMPO("VENTAS", "PEDENCA", "DomiEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "LocaEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "TranspEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "NTranspo", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "NSucEnt", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Observa", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FeSolEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FinVigen", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Status", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "VenRep", 10, 8)
   Call CREACAMPO("VENTAS", "PEDENCA", "ListPre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "CondPag", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "PorComVend", 6, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "TDescApli", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "AnotaPed", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FApro_Nope", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Usua_Apro ", 3, 2)
   Call CREACAMPO("VENTAS", "PEDENCA", "Rec_ConPa", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Motivo", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Destino", 10, 128)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Vehiculo", 10, 128)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Fesalida", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_FeRegreso", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_CondDespa", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Autoriza1", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Autoriza2", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Observaciones", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_NSerieDev", 10, 24)
   Call CREACAMPO("VENTAS", "PEDENCA", "CodCarpImp", 10, 12)
   Call CREACAMPO("VENTAS", "PEDENCA", "TipoIngre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Dto_Ret_Cte", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "DtoxCant", 10, 14)
   '
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroPres", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Rubro", 10, 8)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NumeCRubro", 4, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Anio", 10, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroClie", 4, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Codi_Cli", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "RasoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Referen", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "AteSen", 10, 48)
   Call CREACAMPO("VENTAS", "PRESUVEN", "LisPre", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ConPag", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MonedEmi", 3, 2)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NuOrdIt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiInt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiExt", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Descrip", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Especif", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PreList", 7, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TDescuen", 10, 32)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PorcDesc", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ImpoDesc", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PreUnit", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cantit", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ImToNe", 7, 8)
   Call CREACAMPO("VENTAS", "PRESUVEN", "FechVent", 8, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroPedi", 6, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroFac_Cor", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroFac_Lar", 10, 24)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MarNoVig", 10, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CBase", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TTalle", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TColor", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroCoti", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Reventa", 10, 48)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Antecedentes", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Observa", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Imagen", 11, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Status", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Grco_Cli", 3, 2)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CantBultos", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PesoNeto", 10, 128)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PesoTotal", 10, 128)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NCM", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CondVta", 10, 128)
   '
   Call CREACAMPO("VENTAS", "PREVENTA", "Cod_Lista", 3, 2)
   Call CREACAMPO("VENTAS", "PREVENTA", "Cod_Pieza", 3, 2)
   Call CREACAMPO("VENTAS", "PREVENTA", "Num_Orden", 3, 2)
   Call CREACAMPO("VENTAS", "PREVENTA", "Desc_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Denom_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Carac_Control", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Num_Original", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Precio_Unitario", 7, 8)
   Call CREACAMPO("VENTAS", "PREVENTA", "Cant_Limite", 6, 8)
   Call CREACAMPO("VENTAS", "PREVENTA", "Porc_Limite", 7, 32)
   Call CREACAMPO("VENTAS", "PREVENTA", "MarBorra", 3, 0)
   '
   Call CREACAMPO("VENTAS", "TROQUELES", "CodTroquel", 10, 12)
   Call CREACAMPO("VENTAS", "TROQUELES", "DescriTroquel", 10, 64)
   Call CREACAMPO("VENTAS", "TROQUELES", "DenoTroquel", 10, 64)
   Call CREACAMPO("VENTAS", "TROQUELES", "FormTroquel", 10, 32)
   Call CREACAMPO("VENTAS", "TROQUELES", "AvanceTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "AnchoTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "DiamTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "BocasTroquel", 3, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "MediTroquel", 6, 0)
   '
End Sub
Function VALOBADA_MDB(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   
   sqlx$ = "SELECT " + NOCAM$ + " AS VALORSE "
   sqlx$ = sqlx$ + " FROM " + NOTAB$
   sqlx$ = sqlx$ + " WHERE " + CONDI$
   '
   NOTAB$ = TRIM$(UCase$(NOTAB$))
   X1Q = FLEXCONN_QUEBASE$(NOTAB$)
   On Error Resume Next
   Set MIRS = BASEENUSO.OpenRecordset(sqlx$, dbOpenSnapshot, dbReadOnly)

   VALOBADA_MDB = ""
    
   If Err Then
     On Error GoTo 0
     If MODOMESS$ <> "NOMESS" Then
        TTXX$ = "Imposible Ejecutar Consulta\   \'" + sqlx$ + "'"
        If COMALTER%(TTXX$ + "\\Continuar\Interrumpir") = 2 Then
           Call FINAL("")
        End If
     End If
     GoTo 90
   End If
   On Error GoTo 0
   With MIRS
     If Not (.BOF And .EOF) Then
       VTP% = VarType(!VALORSE)
       If VTP% = 8 Or VTP% < 2 Then
           VALOBADA_MDB = SAFETEXT(!VALORSE)
         Else
           VALOBADA_MDB = !VALORSE
       End If
     End If
   End With
   '
90 On Error Resume Next
   MIRS.Close
   Set MIRS = Nothing
   On Error GoTo 0
End Function

Function FLEXCONN_QUEBASE$(TABLAX$)
    BASEX$ = TRIM$(UCase$(TABLAX$))
    Select Case BASEX$
 
       Case "AUDITRA", "BlOARCHA", "BLOArchB", "MENUES", "MENUES", "SQL_LISELEC", "SQL_QUERIES"
           Set BCONTROL = OpenDatabase(LUACREPLA$("BCONTROL") + "BCONTROL.MDB")
           Set BASEENUSO = BCONTROL
           FLEXCONN_QUEBASE$ = "BCONTROL"
       Case "BSEGCOM", "SUCENTRE", "TRANSPORTES", "VENDEDOR", "AGENDA", "AGENDAORIGINAL"
           Set BASEENUSO = Badaclie
           FLEXCONN_QUEBASE$ = "BADACLIE"
       Case "Proved12"
           Set BASEENUSO = BADAPROVE
           FLEXCONN_QUEBASE$ = "BADAPROVE"
       Case "AUTOENT", "COTIRECI", "DETAIVA", "DETASOLINT", "DETOCABI", "ENCASOLINT", "ENCOCABI", "MACONSIG", "OCOEMI", "OCOMDETA", "OCOMENCA", "PROVECOD"
           Set BASEENUSO = compras
           FLEXCONN_QUEBASE$ = "compras"
       Case "DETAASIEN", "ENCABEASIEN", "PLACOPRE", "PLANCUEN", "TABEJCON", "TAPROYEC"
           Set BASEENUSO = CONTAGEN
           FLEXCONN_QUEBASE$ = "CONTAGEN"
       Case "CAJABANC", "APLIPAGO", "APLICOBRA", "DETMOVCC", "DETRECCC", "ENCMOVCC", "PLAGASPEN", "SACREPEN", "SADEUPEN"
           Set BASEENUSO = CueCorri
           FLEXCONN_QUEBASE$ = "CueCorri"
       Case "COMPXEMPLE", "PADRON"
           Set BASEENUSO = EMPLEADOS
           FLEXCONN_QUEBASE$ = "EMPLEADOS"
       Case "CONTROLFIN", "ENSAR070", "ENSAR072", "ENSAR074", "MEDIFIN", "TANUMSE"
           Set BASEENUSO = Equipos
           FLEXCONN_QUEBASE$ = "Equipos"
       Case "CONTADORES", "EQUIPOS", "GALEIMAG", "PLAMANPREV", "REPUMAQ", "SUBCONJUNTOS"
           Set BASEENUSO = Maquinarias
           FLEXCONN_QUEBASE$ = "Maquinarias"
       Case "MENSAJES"
           Set BASEENUSO = MENSAFLEX
           FLEXCONN_QUEBASE$ = "MENSAFLEX"
       Case "DOCANRE", "GALIMREP", "INFOFAB", "INFOPROY", "NEBRUNE", "OPERFAB", "ORDEFABR", "RECONCILIA", "REPARA", "REPUREPA", "RESERVASF" '       Case "STOCKVIE"
           Set BASEENUSO = Ofabric
           FLEXCONN_QUEBASE$ = "Ofabric"
       Case "ORDEMANT", "ORECORRI", "RECUPROP", "SETARECO"
           Set BASEENUSO = Omantin
           FLEXCONN_QUEBASE$ = "Omantin"
       Case "GRUPOTRA", "LISTAMAT", "ORDETRAB", "PLACOB"
           Set BASEENUSO = OTrabajo
           FLEXCONN_QUEBASE$ = "OTrabajo"
       Case "DESPECIAL", "DETAPRESU", "DOCADJUN", "ENCAPRESU"
           Set BASEENUSO = PresuCam
           FLEXCONN_QUEBASE$ = "PresuCam"
       Case "DETASTOCK", "ENCASTOCK"
           Set BASEENUSO = STOCKORIGI
           FLEXCONN_QUEBASE$ = "STOCKORIGI"
       Case "BALAREQMO", "BOLRECEP", "DETACARTEM", "DETAOSER", "ENCACARTEM", "ENCAOSER", "ENCAREMI", "LISEMPAQ", "LOTINGRE", "MACONSIG", "MASUMCLI", "MOVIREPA", "MOVISTO", "MOVIVIE", "RESERVA", "SALSTOCK", "TRANMOVI", "TRANRECEP", "UBIPRE"
           Set BASEENUSO = Stocks
           FLEXCONN_QUEBASE$ = "Stocks"
       Case "DETASTOCK", "MOVISTO"
           Set BASEENUSO = StockVie
           FLEXCONN_QUEBASE$ = "StockVie"
       Case "DESCOMBI", "DETCOTIZA", "LISTAPRE", "PEDDETA", "PEDENCA", "PRESUVEN", "PREVENTA", "TROQUELES"
           Set BASEENUSO = Ventas
           FLEXCONN_QUEBASE$ = "Ventas"
       Case "BADAMODE", "DOCUREL", "FORMULAS", "FORMUPAT", "GRAFICOS", "HERRAMEN", "HERRAOPER", "MASTER", "SECOPER", "TMUSOPER"
           Set BASEENUSO = Articulos
           FLEXCONN_QUEBASE$ = "Articulos"
    Case Else
       Call COMUNI("Atención !!! Base no Detectada, Verifique.")
       Call FINAL("")
    End Select
End Function
Function CANTREGISTROSMDB&(NOTAB$, CONDI$, Optional MODOMESS$)
   CANTREGISTROSMDB& = 0
   NOTAB$ = TRIM$(UCase$(NOTAB$))
   
    sqlx$ = "SELECT COUNT(*) AS CANTIDAD"
    sqlx$ = sqlx$ + " FROM " & NOTAB$
    If InStr(UCase$(NOTAB), "JOIN") < 1 Then
       'SQLX$ = SQLX$ + " WITH(NOLOCK) "
    End If
    If CONDI$ <> "" Then
      sqlx$ = sqlx$ + " WHERE  " & CONDI$
    End If
   '
   X1Q = FLEXCONN_QUEBASE$(NOTAB$)
   On Error Resume Next
   Set MIRS = BASEENUSO.OpenRecordset("select * from " & NOTAB$, dbOpenSnapshot, dbReadOnly)

   If Err Then
     On Error GoTo 0
     If MODOMESS$ <> "NOMESS" Then
        TTXX$ = "Imposible Ejecutar Consulta\   \'" + sqlx$ + "'"
        If COMALTER%(TTXX$ + "\\Continuar\Interrumpir") = 2 Then
           Call FINAL("")
        End If
     End If
     Exit Function
   End If
   '
    '
    '
    With MIRS
     If Not (.EOF And .EOF) Then
       CANTREGISTROSMDB& = XVALO(!CANTIDAD)
     End If
    End With

End Function


