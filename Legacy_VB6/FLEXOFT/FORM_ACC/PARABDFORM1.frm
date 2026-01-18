VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PARABDFORM1 
   Caption         =   "Bases de Datos y Tablas"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "PARABDFORM1.frx":0000
      Height          =   4215
      Left            =   1050
      OleObjectBlob   =   "PARABDFORM1.frx":0010
      TabIndex        =   10
      Top             =   1260
      Width           =   8100
   End
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "PARABDFORM1.frx":09C2
      Left            =   9240
      List            =   "PARABDFORM1.frx":09C4
      TabIndex        =   9
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      Picture         =   "PARABDFORM1.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4270
      Width           =   1065
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\BADASTRU\BADASTRU.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TRACAMPOS"
      Top             =   4200
      Visible         =   0   'False
      Width           =   5685
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6300
      TabIndex        =   6
      Top             =   735
      Width           =   2850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar"
      Height          =   330
      Left            =   9240
      TabIndex        =   5
      Top             =   735
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   330
      Left            =   3990
      TabIndex        =   4
      Top             =   735
      Width           =   960
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   735
      Width           =   2850
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4935
      TabIndex        =   7
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B.Datos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -315
      TabIndex        =   3
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F:\BADASTRU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      TabIndex        =   1
      Top             =   105
      Width           =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -315
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
End
Attribute VB_Name = "PARABDFORM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BADASTRU As Database
Dim Batablas As Recordset
Dim Bacampos As Recordset
Dim TRAcampos As Recordset
Dim J As Recordset

Private Sub Combo1_Click()
   Combo2.Clear
   TTXX$ = Combo1.Text
   Set FILTABLAS = BADASTRU.OpenRecordset("SELECT * FROM bacampos WHERE BADATOS = '" & TTXX$ & "';")
   '
   PRITA% = 0
   With FILTABLAS
     On Error Resume Next
     .MoveFirst
     While Not .EOF
       NUETA$ = !Tabla
       For U& = 1 To Combo2.ListCount
         If UCase$(Combo2.List(U& - 1)) = UCase$(NUETA$) Then GoTo 19 ' MOVENEXT
       Next U&
       Combo2.AddItem NUETA$
       If PRITA% = 0 Then
         Combo2.Text = NUETA$
         PRITA% = 1
         Call Combo2_Click
       End If
       '
19     .MoveNext
     Wend
   End With
   '
End Sub

Private Sub Combo2_Click()
   '
   If Combo1.Text = "" Then Exit Sub
   If Combo2.Text = "" Then Exit Sub
   '
   On Error Resume Next
   TRAcampos.MoveLast
   While TRAcampos.RecordCount > 0
     TRAcampos.MoveLast
     TRAcampos.Delete
   Wend
   'RECORDSET
   Set J = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS WHERE BADATOS = '" & Combo1.Text & "' AND TABLA = '" & Combo2.Text & "'", dbOpenSnapshot)
   J.MoveLast
   CAREGIS& = J.RecordCount
   '
   J.MoveFirst
   RENGLOI& = 1
8  TRAcampos.AddNew
   For COLIU% = 1 To J.Fields.Count
      TRAcampos.Fields(COLIU% - 1).Value = J.Fields(COLIU% - 1).Value
   Next COLIU
   TRAcampos.Update
   '
   If Not J.EOF Then
     If RENGLOI& < CAREGIS& Then
       RENGLOI& = RENGLOI& + 1
       J.MoveNext
       GoTo 8
     End If
   End If
   '
   Set J = Nothing
   Data1.Refresh
   DBGrid1.Refresh
   '
' Determino no visible las columnas b.datos y tabla de la grilla
   DBGrid1.Columns(0).Visible = False
   DBGrid1.Columns(1).Visible = False
   DBGrid1.Columns(6).Visible = False
   DBGrid1.Columns(3).Width = 500
   DBGrid1.Columns(5).Width = 4700
   '
   DBGrid1.Visible = True
   '
End Sub

Private Sub Command1_Click()
   '
   NUEBASE$ = InputBox$("Nueva Base de Datos", "Nombre de la Base de Datos")
   If NUEBASE$ <> "" Then
     For U& = 1 To Combo1.ListCount
       If UCase$(Combo1.List(U& - 1)) = UCase$(NUEBASE$) Then
         MsgBox "Ya Existe Base con ese Nombre"
         Combo1.Text = NUEBASE$
         Exit Sub
       End If
     Next U&
     '
     Combo1.AddItem NUEBASE$
     Combo1.Text = NUEBASE$
   End If
   '
End Sub

Private Sub Command2_Click()
   '
   NUETABLA$ = InputBox$("Nueva Tabla", "Nombre de la Tabla")
   If NUETABLA$ <> "" Then
     For U& = 1 To Combo2.ListCount
       If UCase$(Combo2.List(U& - 1)) = UCase$(NUETABLA$) Then
         MsgBox "Ya Existe Tabla con ese Nombre"
         Combo2.Text = NUETABLA$
         Exit Sub
       End If
     Next U&
     '
     Combo2.AddItem NUETABLA$
     Combo2.Text = NUETABLA$
     Call Combo2_Click
   End If
   '
End Sub

Private Sub Command3_Click()
   '
   If Combo1 = "" Then Exit Sub
   If Combo2 = "" Then Exit Sub
   '
   'Chequeo que las no haya campos en blanco en la grilla
   For RENGLON = 1 To Data1.Recordset.RecordCount
     
      DBGrid1.Row = RENGLON - 1
      
      For COLUMNA = 3 To 6   ' DBGrid1.Columns.Count
        DBGrid1.Col = COLUMNA - 1
        If Trim$(DBGrid1.Text) = "" Then
            MsgBox ("Todos los campos deben completarse")
            Exit Sub
        End If
      Next COLUMNA
      '
   Next RENGLON
   '
   SINO = MsgBox("QUIERE GUARDAR", vbYesNo)
   If SINO = vbYes Then
     '
     DIVERI$ = Format$(Now, "ddmmyyhhmmss")
     With Batablas
        On Error Resume Next
        .MoveFirst
        While Not .EOF
          If UCase$(!BaDatos) = UCase$(Combo1.Text) Then
            If UCase$(!Tabla) = UCase$(Combo2.Text) Then
              .Edit
              !DIVERIF = DIVERI$
              .Update
              GoTo 5
            End If
          End If
          .MoveNext
        Wend
        .AddNew
        !BaDatos = Combo1.Text
        !Tabla = Combo2.Text
        !DIVERIF = DIVERI$
        .Update
        '
5    End With
     '
     Set REGISTROSABORRAR = BADASTRU.OpenRecordset("SELECT * FrOM BACAMPOS WHERE BADATOS= '" & Combo1.Text & "' AND TABLA = '" & Combo2.Text & "'", dbOpenDynaset)
     On Error Resume Next
     With REGISTROSABORRAR
       .MoveFirst
       If Not Err Then
         Do While Not .EOF
          .Edit
          !MARBORRA = 1
          .Update
          .MoveNext
         Loop
       End If
     End With
     '
     For RENGLON = 1 To Data1.Recordset.RecordCount
        DBGrid1.Row = RENGLON - 1
        Bacampos.AddNew
        For COLUMNA = 3 To 6   ' DBGrid1.Columns.Count
          DBGrid1.Col = COLUMNA - 1
          Bacampos.Fields(COLUMNA - 1).Value = DBGrid1.Text
        Next COLUMNA
        Bacampos.Fields(0).Value = Combo1.Text
        Bacampos.Fields(1).Value = Combo2.Text
        Bacampos.Fields(6).Value = 0
        Bacampos.Update
     Next RENGLON
     '
     Set REGISTROSABORRAR = BADASTRU.OpenRecordset("SELECT * FrOM BACAMPOS WHERE (BADATOS= '" & Combo1.Text & "' AND TABLA = '" & Combo2.Text & "' and MARBORRA = 1) OR TRIM$(BADATOS) = ''", dbOpenDynaset)
     On Error Resume Next
     With REGISTROSABORRAR
       .MoveLast
       If Not Err Then
         Do While Not .BOF
           .Delete
           .MovePrevious
         Loop
       End If
     End With
     '
   End If
   '
   Data1.Recordset.MoveLast
   '
End Sub



Private Sub Form_Load()
   '
   'Cargo la lista  del list
   With List1
    .List(0) = "1    boolean"
    .List(1) = "3    Integer"
    .List(2) = "4    Long"
    .List(3) = "5    Currency"
    .List(4) = "6    Single"
    .List(5) = "7    Double "
    .List(6) = "8    Date"
    .List(7) = "10  Text"
    .List(8) = "12  Memo"
   End With
   '
   Call CREABADAT
   '
   Combo1.Clear
   With Batablas
     On Error Resume Next
     .MoveFirst
     While Not .EOF
       Combo1.AddItem !BaDatos
       .MoveNext
     Wend
   End With
   '
   On Error Resume Next
   TRAcampos.MoveLast
   While TRAcampos.RecordCount > 0
     TRAcampos.MoveLast
     TRAcampos.Delete
   Wend
   Data1.Refresh
   '
   DBGrid1.Columns(0).Visible = False
   DBGrid1.Columns(1).Visible = False
   DBGrid1.Columns(6).Visible = False
   DBGrid1.Columns(3).Width = 500
   DBGrid1.Columns(5).Width = 4700
   '
End Sub

Sub CENTRAFORM(NOFORMU As Form)
   '
   If Screen.Height > NOFORMU.Height + 600 Then
     NOFORMU.Top = (Screen.Height - NOFORMU.Height) / 2
   End If
   If Screen.Width > NOFORMU.Width Then
     NOFORMU.Left = (Screen.Width - NOFORMU.Width) / 2
   End If
   DoEvents
   '
End Sub

Sub CREABADAT()
   '
   Static YABIER%
   '
   Screen.MousePointer = 11
   '
   LUACC$ = Label2 + "\"
   '
   If YABIER% > 0 Then GoTo 90
   '
   Dim wrkPredeterminado As Workspace
   Set wrkPredeterminado = DBEngine.Workspaces(0)
   '
   On Error Resume Next
   ' crea el directorio - carpeta
   MkDir Label2
   '
   ' crea la base de datos
   Set CREABASE = wrkPredeterminado.CreateDatabase(LUACC$ + "BADASTRU.MDB", dbLangGeneral)
   Set BADASTRU = OpenDatabase(LUACC$ + "BADASTRU.MDB")
   '
   ' crea la tabla "BATABLAS"
   Set CREATABLA = BADASTRU.CreateTableDef("BATABLAS")
   CREATABLA.Fields.Append CREATABLA.CreateField("BaDatos", dbText, 32)
   BADASTRU.TableDefs.Append CREATABLA
   '
   ' agrega campos a "BATABLAS"
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("Tabla", dbText, 32)
     .Append CREATABLA.CreateField("DiVerif", dbText, 12)
   End With
   '
   ' crea la tabla "BACAMPOS"
   Set CREATABLA = BADASTRU.CreateTableDef("BACAMPOS")
   CREATABLA.Fields.Append CREATABLA.CreateField("BaDatos", dbText, 32)
   BADASTRU.TableDefs.Append CREATABLA
   '
   ' agrega campos a "BACAMPOS"
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("Tabla", dbText, 32)
     .Append CREATABLA.CreateField("Campo", dbText, 32)
     .Append CREATABLA.CreateField("Tipo", dbText, 32)
     .Append CREATABLA.CreateField("Longitud", dbInteger)
     .Append CREATABLA.CreateField("Observa", dbMemo)
     .Append CREATABLA.CreateField("MarBorra", dbInteger)
   End With
   '
   ' crea la tabla "TRACAMPOS"
   Set CREATABLA = BADASTRU.CreateTableDef("TRACAMPOS")
   CREATABLA.Fields.Append CREATABLA.CreateField("BaDatos", dbText, 32)
   BADASTRU.TableDefs.Append CREATABLA
   '
   ' agrega campos a "TRACAMPOS"
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("Tabla", dbText, 32)
     .Append CREATABLA.CreateField("Campo", dbText, 32)
     .Append CREATABLA.CreateField("Tipo", dbText, 32)
     .Append CREATABLA.CreateField("Longitud", dbInteger)
     .Append CREATABLA.CreateField("Observa", dbMemo)
     .Append CREATABLA.CreateField("MarBorra", dbInteger)
   End With
   '
90 On Error GoTo 0
   Set Batablas = BADASTRU.OpenRecordset("BATABLAS")
   Set Bacampos = BADASTRU.OpenRecordset("BACAMPOS")
   Set TRAcampos = BADASTRU.OpenRecordset("TRACAMPOS")
   On Error GoTo 0
   YABIER% = 1
   Screen.MousePointer = 1
   '
End Sub

