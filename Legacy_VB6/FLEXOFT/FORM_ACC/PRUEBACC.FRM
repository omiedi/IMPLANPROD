VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   1155
      TabIndex        =   0
      Top             =   735
      Width           =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CREABASE(BASEDA$, TABLAX$, NCAMPO$, TCAMPO, Optional LCAMPO)
   '
   Static CTXX%
   Static wrkPredeterminado As Workspace
   Static LUACC$
   '
   If CTXX% < 1 Then
     Set wrkPredeterminado = DBEngine.Workspaces(0)
LUACC$ = "C:\MIS DOCUMENTOS\PRUEBA\"
     CTXX% = 1
   End If
   '
   ' crea la base de datos BASEDA$
   On Error Resume Next
   Set CREABA = wrkPredeterminado.CreateDatabase(LUACC$ + BASEDA$ + ".MDB", dbLangGeneral)
   If Err < 1 Or Err = 3204 Then     ' LA BASE DE DATOS YA EXISTE
       GoTo 10
     Else
       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Base de Datos " + LUACC$ + BASEDA$
       Close: End
   End If
   '
10 On Error Resume Next
   ' crea la tabla TABLAX$
   Set BASDAT = OpenDatabase(LUACC$ + BASEDA$ + ".MDB")
   Set CREATA = BASDAT.CreateTableDef(TABLAX$)
   CREATA.Fields.Append CREATA.CreateField(NCAMPO$, TCAMPO, LCAMPO)
   BASDAT.TableDefs.Append CREATA
   If Err < 1 Or Err = 3010 Then     ' LA tabla ya existe
       GoTo 20
     Else
       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Tabla en "
       Close: End
   End If
   '
20 Set CREACA = BASDAT.TableDefs(TABLAX$)
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
           MsgBox "Error " & Err & " al Crear Campo '" + NCAMPO$ + "' en Tabla '" + TABLAX$ + "'."
       End If
     End If
   End With
   '
End Sub
'
Private Sub Command1_Click()
   '
   'Dim wrkPredeterminado As Workspace
   'Set wrkPredeterminado = DBEngine.Workspaces(0)
     '
'   On Error Resume Next
'   ' crea la base de datos
'   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + "ARTICULOS.MDB", dbLangGeneral)
'   If Err < 1 Or Err = 3204 Then     ' LA BASE DE DATOS YA EXISTE
'       GoTo 10
'     Else
'       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Base de Datos XXX"
'       Close: End
'   End If
'   '
'10 On Error Resume Next
'   Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
'   '
'   ' crea la tabla "SECOPER"
'   Set CREATABLA = Articulos.CreateTableDef("SECOPER")
'   CREATABLA.Fields.Append CREATABLA.CreateField("CodIConj", dbInteger)
'   Articulos.TableDefs.Append CREATABLA
'   If Err < 1 Or Err = 3010 Then     ' LA tabla ya existe
'       GoTo 20
'     Else
'       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Tabla en "
'       Close: End
'   End If
'   '
    Call CREABASE("ARTICULOS", "SECOPER", "CodXConj", dbText, 24)
    Call CREABASE("ARTICULOS", "SECOPER", "NumeOper", dbInteger)
 Stop: End
   '
   ' agrega campos a "SECOPER"
20 Set CREACAMPOS = Articulos.TableDefs("SECOPER")
   On Error GoTo 0
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

Private Sub Command2_Click()
   '
   Dim wrkPredeterminado As Workspace
   Set wrkPredeterminado = DBEngine.Workspaces(0)
     '
   On Error Resume Next
   ' crea la base de datos
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + "ARTICULOS.MDB", dbLangGeneral)
   If Err < 1 Or Err = 3204 Then     ' LA BASE DE DATOS YA EXISTE
       GoTo 10
     Else
       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Base de Datos XXX"
       Close: End
   End If
   '
10 On Error Resume Next
   Set Articulos = OpenDatabase(LUACC$ + "ARTICULOS.MDB")
   '
   ' crea la tabla "SECOPER"
   Set CREATABLA = Articulos.CreateTableDef("SECOPER")
   CREATABLA.Fields.Append CREATABLA.CreateField("CodIConj", dbInteger)
   Articulos.TableDefs.Append CREATABLA
   If Err < 1 Or Err = 3010 Then     ' LA tabla ya existe
       GoTo 20
     Else
       MsgBox "Error (" + Trim$(Str$(Err)) + ") - Imposible Crear Tabla en "
       Close: End
   End If
   '
   '
   ' agrega campos a "SECOPER"
20 Set CREACAMPOS = Articulos.TableDefs("SECOPER")
   On Error GoTo 0
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


