VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PARABDFORM 
   Caption         =   "Bases de Datos y Tablas"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Select Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   12
      Top             =   4515
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insertar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   4040
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "PARABDFORM.frx":0000
      Height          =   4335
      Left            =   1050
      OleObjectBlob   =   "PARABDFORM.frx":0014
      TabIndex        =   10
      Top             =   1260
      Width           =   8100
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "PARABDFORM.frx":09E7
      Left            =   9240
      List            =   "PARABDFORM.frx":09E9
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
      Height          =   615
      Left            =   9240
      Picture         =   "PARABDFORM.frx":09EB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5000
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "F:\FLEXOFT\NOVEDAD\BADASTRU.MDB"
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
      Caption         =   "F:\FLEXOFT\NOVEDAD"
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
Attribute VB_Name = "PARABDFORM"
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
   If Combo1.Text = "" Or Combo2.Text = "" Then
     Data1.DatabaseName = ""
     Exit Sub
   End If
   '
   On Error Resume Next
   TRAcampos.MoveLast
   While TRAcampos.EOF = False  'TRAcampos.RecordCount > 0
     TRAcampos.MoveLast
     TRAcampos.Delete
   Wend
   'RECORDSET
   Set J = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS WHERE BADATOS = '" & Combo1.Text & "' AND TABLA = '" & Combo2.Text & "' ORDER BY NuOrden ASC;", dbOpenSnapshot)
   'ORDER BY NuOrden ASC;
   J.MoveLast
   CAREGIS& = J.RecordCount
   '
   J.MoveFirst
   RENGLOI& = 1
8  TRAcampos.AddNew   ' YA CONVERTIDO
            On Error Resume Next
   TRAcampos!CodiEmpr = Codiemp%
            On Error GoTo 0
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
 'Determino no visible las columnas b.datos y tabla de la grilla
   DBGrid1.Columns(0).Visible = False
   DBGrid1.Columns(1).Visible = False
   DBGrid1.Columns(6).Visible = False
   DBGrid1.Columns(7).Visible = False
   DBGrid1.Columns(3).Width = 500
   DBGrid1.Columns(5).Width = 4300
   DBGrid1.Columns(8).Width = 350
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
   ' Controla errores de tipo de dato
   On Error GoTo CONTROLERROR
   'Chequeo que las no haya campos en blanco en la grilla
   Data1.Recordset.MoveFirst
   Do While Data1.Recordset.EOF = False
     For I = 3 To 5
       If IsNull(Data1.Recordset.Fields(I)) Then
         MsgBox ("Todos los datos deben completarse")
         Exit Sub
       End If
       If I = 3 Then
         If Val(Data1.Recordset.Fields(I)) = 10 Then
           If Val(Data1.Recordset.Fields(I + 1)) < 1 Then
             MsgBox "Los Campos Texto (Tipo=10)" + Chr$(10) + Chr$(13) + "Deben Tener Longitud Mayor que Cero."
             Exit Sub
           End If
         End If
       End If
     Next I
     Data1.Recordset.MoveNext
   Loop
   '

   SINO = MsgBox("¿ Quiere Guardar los Datos ?", vbYesNo)
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
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = Codiemp%
        On Error GoTo 0
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
     With Bacampos
     For RENGLON = 1 To Data1.Recordset.RecordCount
        Data1.Recordset.AbsolutePosition = RENGLON - 1
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = Codiemp%
        On Error GoTo 0
        .Fields(0).Value = Combo1.Text
        .Fields(1).Value = Combo2.Text
            For COLUMNA = 2 To 5
                .Fields(COLUMNA).Value = Data1.Recordset.Fields(COLUMNA).Value
            Next COLUMNA
        .Fields(6).Value = 0
        .Fields(7).Value = RENGLON
        .Fields(8).Value = Data1.Recordset.Fields(8).Value
        .Update
     Next RENGLON
     End With
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
     On Error GoTo 0
     'Call GENABREBASE
     Call GENABREBASE_SQL
     Call GENABREBASE_ACC
     MsgBox "'ABREBASE.BAS' Re-Generado Existosamente."
     '
   End If
   '
   Data1.Recordset.MoveLast
   DBGrid1.Refresh
   '
   Exit Sub
   '
CONTROLERROR:
   MsgBox ("Verifique el tipo de dato introducido")
   On Error GoTo 0
End Sub



Private Sub Command4_Click()
  On Error GoTo 6 'Si no hay registros activos va a etiqueta 6
  Dim VALORI(256) As Variant
  BMK = DBGrid1.FirstRow
  BKR = DBGrid1.Row: BKC = DBGrid1.Col
  APOSI& = Data1.Recordset.AbsolutePosition
  UREGIS& = Data1.Recordset.RecordCount - 1
  Data1.Recordset.AddNew   ' YA CONVERTIDO
                 On Error Resume Next
  Data1.Recordset!CodiEmpr = Codiemp%
                 On Error GoTo 0
  Data1.Recordset.Update
  For U& = UREGIS& To APOSI& Step -1
       Data1.Recordset.AbsolutePosition = U&
       Data1.Recordset.Edit
     For KU& = 0 To Data1.Recordset.Fields.Count - 1
       VALORI(KU&) = Data1.Recordset.Fields(KU&).Value
     Next KU&
       Data1.Recordset.AbsolutePosition = U& + 1
       Data1.Recordset.Edit
     For KU& = 0 To Data1.Recordset.Fields.Count - 1
       Data1.Recordset.Fields(KU&).Value = VALORI(KU&)
     Next KU&
       Data1.Recordset.Update
  Next U&
  Data1.Recordset.AbsolutePosition = APOSI&
  DBGrid1.FirstRow = BMK
  DBGrid1.Row = BKR: DBGrid1.Col = BKC
  DBGrid1.Refresh
6:
End Sub

Private Sub Command5_Click()

On Error GoTo 5 'Si no hay registros activos va a etiqueta 5
With Data1
I = MsgBox("¿Esta seguro que desea eliminar el registro activo?", vbYesNo, "Eliminar registro")
    If I = vbYes Then
        pos& = .Recordset.AbsolutePosition
        .Recordset.Delete
    Else
        GoTo 5
    End If
    .Recordset.AbsolutePosition = pos&
    DBGrid1.Refresh
End With
5:
End Sub

Private Sub Command6_Click()
   '
   TTXX$ = ""
   For RENGLON = 1 To Data1.Recordset.RecordCount
     Data1.Recordset.AbsolutePosition = RENGLON - 1
     TTXX$ = TTXX$ + "!"
     TTXX$ = TTXX$ + Data1.Recordset.Fields(2).Value
     TTXX$ = TTXX$ + " = " + Chr$(10)
   Next RENGLON
   Clipboard.Clear
   Clipboard.SetText (TTXX$)
   DBGrid1.Refresh
   '
End Sub


'\\\OT-06-0106-WAR-23/03/06///
Sub GENABREBASE()
   '
   Dim VECINS$()
   N1% = FreeFile
   ILET$ = "D"
10 RUT1$ = ILET$ + ":\Archivos de Programa\DevStudio\VB\FLEXOFT\BIBL_ACC\"
   RUT1$ = RUT1$ + "ABREBASE.bas"
   '
   On Error Resume Next
   Open RUT1$ For Input As N1%
   If Err Then
     If ILET$ = "D" Then
       ILET$ = "C"
       GoTo 10
     End If
     MsgBox "Imposible Abrir Archivo: " + RUT1$
     On Error GoTo 0
     Exit Sub
   End If
   '
   ' CARGA ARCHIVO EN MEMORIA
   CAREG% = 0
   Do While Not EOF(N1%)
     Line Input #N1%, AXK$
     If UCase$(Trim$(AXK$)) = "SUB GENERBASE(NOMBASE$)" Then GoTo 20
     If CAREG% >= UBound(VECINS$) Then
       ReDim Preserve VECINS$(CAREG% + 128)
     End If
     CAREG% = CAREG% + 1
     VECINS$(CAREG%) = AXK$
   Loop
   '
20 ReDim Preserve VECINS$(CAREG% + 128)
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "SUB GENERBASE(NOMBASE$)"
   '
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   '"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   RUTIANT$ = RUTIACT$"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   RUTIACT$ = " + Chr$(34) + "GENERBASE(" + Chr$(34) + " + NOMBASE$ + " + Chr$(34) + ")" + Chr$(34)
   '
   'AGREGA INSTRUCCIONES SEGUN RECORDSET
   Set J = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS ORDER BY BaDatos ASC, Tabla ASC, NuOrden ASC ", dbOpenSnapshot)
   With J
     On Error Resume Next
     .MoveFirst
     If Err Then
       MsgBox "ERROR: Imposible Agregar Instrucciones 'CREACAMPO')"
       Exit Sub
     End If
     On Error GoTo 0
     BA0$ = "": TA0$ = ""
     Do While Not .EOF
       BA1$ = .Fields("BaDatos").Value
       TA1$ = .Fields("Tabla").Value
       CA1$ = .Fields("Campo").Value
       TICA1% = .Fields("Tipo").Value
       LOCA1% = .Fields("Longitud").Value
       STRADD1$ = "CALL CREACAMPO(" + Chr$(34) + BA1$ + Chr$(34) + ", " + Chr$(34)
       STRADD1$ = STRADD1$ + TA1$ + Chr$(34) + ", " + Chr$(34) + CA1$ + Chr$(34)
       STRADD1$ = STRADD1$ + ", " + CStr(TICA1%) + ", " + CStr(LOCA1%) + ")"
       '
       If BA1$ <> BA0$ Then
            If BA0$ <> "" Then
              CAREG% = CAREG% + 1
              VECINS$(CAREG%) = "   END IF"
              CAREG% = CAREG% + 1
              VECINS$(CAREG%) = "   '"
              CAREG% = CAREG% + 1
              VECINS$(CAREG%) = "   '********************************************************************************************"
            End If
            CAREG% = CAREG% + 1
            VECINS$(CAREG%) = "   '"
            CAREG% = CAREG% + 1
            VECINS$(CAREG%) = "   IF TRIM$(UCASE$(NOMBASE$)) = " + Chr$(34) + Trim$(UCase$(BA1$)) + Chr$(34) + " THEN"
          ElseIf TA1$ <> TA0$ Then
            CAREG% = CAREG% + 1
            VECINS$(CAREG%) = "     '"
       End If
       BA0$ = BA1$: TA0$ = TA1$
       '
       If CAREG% >= UBound(VECINS$) - 10 Then
         ReDim Preserve VECINS$(CAREG% + 128)
       End If
       CAREG% = CAREG% + 1
       VECINS$(CAREG%) = "     " + STRADD1$
     .MoveNext
     Loop
   End With
   '
   ReDim Preserve VECINS$(CAREG% + 128)
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   END IF"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   '"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   RUTIACT$ = RUTIANT$"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   '"
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "END SUB"
   '
   Close #N1%
   '
   ' VUELVE A ESCRIBIR EL ARCHIVO
   Open RUT1$ For Output As N1%
   For IREG% = 1 To CAREG%
     Print #N1%, VECINS$(IREG%)
   Next IREG%
   Close #N1%
   '
End Sub
'\\\OT-06-0106-WAR-FIN///

Sub GENABREBASE_SQL()
   '
   Dim VECINS$(), VECREA$()
   N1% = FreeFile
   ILET$ = "D"
10 RUT1$ = ILET$ + ":\Archivos de Programa\DevStudio\VB\FLEXOFT\BIBL_SQL\"
   RUT1$ = RUT1$ + "ABREBASE_SQL.bas"
   '
   On Error Resume Next
   Open RUT1$ For Input As N1%
   If Err Then
     If ILET$ = "D" Then
       ILET$ = "C"
       GoTo 10
     End If
     MsgBox "Imposible Abrir Archivo: " + RUT1$
     On Error GoTo 0
     Exit Sub
   End If
   '
   ' CARGA ARCHIVO EN MEMORIA
   CAREG% = 0: CAARCREA% = 0
   Do While Not EOF(N1%)
     Line Input #N1%, AXK$
     If UCase$(Trim$(AXK$)) = "SUB GENERBASE(NOMBASE$)" Then GoTo 20
     If CAREG% >= UBound(VECINS$) Then
       ReDim Preserve VECINS$(CAREG% + 128)
     End If
     CAREG% = CAREG% + 1
     VECINS$(CAREG%) = AXK$
   Loop
   '
20 ReDim Preserve VECINS$(CAREG% + 128)
   ReDim Preserve VECREA$(128)
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "SUB GENERBASE(NOMBASE$)"
   '
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   '"
'   CAREG% = CAREG% + 1
'   VECINS$(CAREG%) = "   RUTIANT$ = RUTIACT$"
'   CAREG% = CAREG% + 1
'   VECINS$(CAREG%) = "   RUTIACT$ = " + Chr$(34) + "GENERBASE(" + Chr$(34) + " + NOMBASE$ + " + Chr$(34) + ")" + Chr$(34)
   '
   'AGREGA INSTRUCCIONES SEGUN RECORDSET
   Set J = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS ORDER BY BaDatos ASC, Tabla ASC, NuOrden ASC ", dbOpenSnapshot)
   With J
     On Error Resume Next
     .MoveFirst
     If Err Then
       MsgBox "ERROR: Imposible Agregar Instrucciones 'CREACAMPO')"
       Exit Sub
     End If
     On Error GoTo 0
     BA0$ = "": TA0$ = ""
     Do While Not .EOF
       BA1$ = .Fields("BaDatos").Value
       TA1$ = .Fields("Tabla").Value
       CA1$ = .Fields("Campo").Value
       TICA1% = .Fields("Tipo").Value
       LOCA1% = .Fields("Longitud").Value
       IDEX1% = 0
       On Error Resume Next
       IDEX1% = .Fields("Idx").Value
       On Error GoTo 0
       If IDEX1% > 0 Then IDEX1% = 1
       STRADD1$ = "CALL CREACAMPO(" + Chr$(34) + BA1$ + Chr$(34) + ", " + Chr$(34)
       STRADD1$ = STRADD1$ + TA1$ + Chr$(34) + ", " + Chr$(34) + CA1$ + Chr$(34)
       STRADD1$ = STRADD1$ + ", " + CStr(TICA1%) + ", " + CStr(LOCA1%)
       If IDEX1% > 0 Then STRADD1$ = STRADD1$ + ", 1 "
       STRADD1$ = STRADD1$ + ")"
       '
       If BA1$ <> BA0$ Then
             VECINS$(CAREG% + 1) = "   IF TRIM$(UCASE$(NOMBASE$)) = " + Chr$(34) + Trim$(UCase$(BA1$)) + Chr$(34) + " THEN CALL CREATABLA_" + BA1$ + ": EXIT SUB"
             CAREG% = CAREG% + 1
             '
             If TA0$ <> "" Then
                VECREA$(CARCREA% + 1) = "   '"
                VECREA$(CARCREA% + 2) = "END SUB"
                CARCREA% = CARCREA% + 2
             End If
             '
             VECREA$(CARCREA% + 1) = "'"
             VECREA$(CARCREA% + 2) = "SUB CREATABLA_" + BA1$
             CARCREA% = CARCREA% + 2
             '
             BA0$ = BA1$
             TA0$ = ""
       End If
       '
       If TA1$ <> TA0$ Then
           CARCREA% = CARCREA% + 1
           VECREA$(CARCREA%) = "   '"
           TA0$ = TA1$
       End If
       BA0$ = BA1$: TA0$ = TA1$
       '
       If CARCREA% >= UBound(VECREA$) - 10 Then
          ReDim Preserve VECREA$(CARCREA% + 128)
       End If
       If CAREG% >= UBound(VECINS$) - 10 Then
          ReDim Preserve VECINS$(CAREG% + 128)
       End If
       CARCREA% = CARCREA% + 1
       VECREA$(CARCREA%) = "   " + STRADD1$
     .MoveNext
     Loop
   End With
   '
   VECINS$(CAREG% + 1) = "   '"
   VECINS$(CAREG% + 2) = "END SUB"
   CAREG% = CAREG% + 2
   '
   VECREA$(CARCREA% + 1) = "   '"
   VECREA$(CARCREA% + 2) = "END SUB"
   CARCREA% = CARCREA% + 2
   '
   Close #N1%
   '
   ' VUELVE A ESCRIBIR EL ARCHIVO
   Open RUT1$ For Output As N1%
   For IREG% = 1 To CAREG%
     Print #N1%, VECINS$(IREG%)
   Next IREG%
   For IREG% = 1 To CARCREA%
     Print #N1%, VECREA$(IREG%)
   Next IREG%
   Close #N1%
   'MsgBox "'ABREBASE.BAS' Re-Generado Existosamente."
   '
End Sub
'\\\OT-06-0106-WAR-FIN///

Sub GENABREBASE_ACC()
   '
   Dim VECINS$(), VECREA$()
   N1% = FreeFile
   ILET$ = "D"
10 RUT1$ = ILET$ + ":\Archivos de Programa\DevStudio\VB\FLEXOFT\BIBL_ACC\"
   RUT1$ = RUT1$ + "ABREBASE.bas"
   '
   On Error Resume Next
   Open RUT1$ For Input As N1%
   If Err Then
     If ILET$ = "D" Then
       ILET$ = "C"
       GoTo 10
     End If
     MsgBox "Imposible Abrir Archivo: " + RUT1$
     On Error GoTo 0
     Exit Sub
   End If
   '
   ' CARGA ARCHIVO EN MEMORIA
   CAREG% = 0: CAARCREA% = 0
   Do While Not EOF(N1%)
     Line Input #N1%, AXK$
     If UCase$(Trim$(AXK$)) = "SUB GENERBASE(NOMBASE$)" Then GoTo 20
     If CAREG% >= UBound(VECINS$) Then
       ReDim Preserve VECINS$(CAREG% + 128)
     End If
     CAREG% = CAREG% + 1
     VECINS$(CAREG%) = AXK$
   Loop
   '
20 ReDim Preserve VECINS$(CAREG% + 128)
   ReDim Preserve VECREA$(128)
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "SUB GENERBASE(NOMBASE$)"
   '
   CAREG% = CAREG% + 1
   VECINS$(CAREG%) = "   '"
'   CAREG% = CAREG% + 1
'   VECINS$(CAREG%) = "   RUTIANT$ = RUTIACT$"
'   CAREG% = CAREG% + 1
'   VECINS$(CAREG%) = "   RUTIACT$ = " + Chr$(34) + "GENERBASE(" + Chr$(34) + " + NOMBASE$ + " + Chr$(34) + ")" + Chr$(34)
   '
   'AGREGA INSTRUCCIONES SEGUN RECORDSET
   Set J = BADASTRU.OpenRecordset("SELECT * FROM BACAMPOS ORDER BY BaDatos ASC, Tabla ASC, NuOrden ASC ", dbOpenSnapshot)
   With J
     On Error Resume Next
     .MoveFirst
     If Err Then
       MsgBox "ERROR: Imposible Agregar Instrucciones 'CREACAMPO')"
       Exit Sub
     End If
     On Error GoTo 0
     BA0$ = "": TA0$ = ""
     Do While Not .EOF
       BA1$ = .Fields("BaDatos").Value
       TA1$ = .Fields("Tabla").Value
       CA1$ = .Fields("Campo").Value
       TICA1% = .Fields("Tipo").Value
       LOCA1% = .Fields("Longitud").Value
       IDEX1% = 0
       On Error Resume Next
       IDEX1% = .Fields("Idx").Value
       On Error GoTo 0
       If IDEX1% > 0 Then IDEX1% = 1
       STRADD1$ = "CALL CREACAMPO(" + Chr$(34) + BA1$ + Chr$(34) + ", " + Chr$(34)
       STRADD1$ = STRADD1$ + TA1$ + Chr$(34) + ", " + Chr$(34) + CA1$ + Chr$(34)
       STRADD1$ = STRADD1$ + ", " + CStr(TICA1%) + ", " + CStr(LOCA1%)
       'If IDEX1% > 0 Then STRADD1$ = STRADD1$ + ", 1 "
       STRADD1$ = STRADD1$ + ")"
       '
       If BA1$ <> BA0$ Then
             VECINS$(CAREG% + 1) = "   IF TRIM$(UCASE$(NOMBASE$)) = " + Chr$(34) + Trim$(UCase$(BA1$)) + Chr$(34) + " THEN CALL CREATABLA_" + BA1$ + ": EXIT SUB"
             CAREG% = CAREG% + 1
             '
             If TA0$ <> "" Then
                VECREA$(CARCREA% + 1) = "   '"
                VECREA$(CARCREA% + 2) = "END SUB"
                CARCREA% = CARCREA% + 2
             End If
             '
             VECREA$(CARCREA% + 1) = "'"
             VECREA$(CARCREA% + 2) = "SUB CREATABLA_" + BA1$
             CARCREA% = CARCREA% + 2
             '
             BA0$ = BA1$
             TA0$ = ""
       End If
       '
       If TA1$ <> TA0$ Then
           CARCREA% = CARCREA% + 1
           VECREA$(CARCREA%) = "   '"
           TA0$ = TA1$
       End If
       BA0$ = BA1$: TA0$ = TA1$
       '
       If CARCREA% >= UBound(VECREA$) - 10 Then
          ReDim Preserve VECREA$(CARCREA% + 128)
       End If
       If CAREG% >= UBound(VECINS$) - 10 Then
          ReDim Preserve VECINS$(CAREG% + 128)
       End If
       CARCREA% = CARCREA% + 1
       VECREA$(CARCREA%) = "   " + STRADD1$
     .MoveNext
     Loop
   End With
   '
   VECINS$(CAREG% + 1) = "   '"
   VECINS$(CAREG% + 2) = "END SUB"
   CAREG% = CAREG% + 2
   '
   VECREA$(CARCREA% + 1) = "   '"
   VECREA$(CARCREA% + 2) = "END SUB"
   CARCREA% = CARCREA% + 2
   '
   Close #N1%
   '
   ' VUELVE A ESCRIBIR EL ARCHIVO
   Open RUT1$ For Output As N1%
   For IREG% = 1 To CAREG%
     Print #N1%, VECINS$(IREG%)
   Next IREG%
   For IREG% = 1 To CARCREA%
     Print #N1%, VECREA$(IREG%)
   Next IREG%
   Close #N1%
   'MsgBox "'ABREBASE.BAS' Re-Generado Existosamente."
   '
End Sub
'\\\OT-06-0106-WAR-FIN///

Private Sub DBGrid1_DblClick()
   Clipboard.Clear
   Clipboard.SetText DBGrid1.Text
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If DBGrid1.Col = 2 Then
     Clipboard.Clear
     Clipboard.SetText DBGrid1.Text
   End If
End Sub

Private Sub Form_Load()
   '
   'Cargo la lista  del list
   With List1
    .List(0) = "1    boolean"
    .List(1) = "2    Binary"
    .List(2) = "3    Integer"
    .List(3) = "4    Long"
    .List(4) = "5    Currency"
    .List(5) = "6    Single"
    .List(6) = "7    Double "
    .List(7) = "8    Date"
    .List(8) = "10  Text"
    .List(9) = "11  Imagen"
    .List(10) = "12  Memo"
   End With
   '
   Call CREABADAT
   '
   Combo1.Clear
  'Tomo unicamente las base de datos de diferente nombre
   Set BAS = BADASTRU.OpenRecordset("select DISTINCT Badatos from BATABLAS", dbOpenDynaset)
   With BAS
    On Error Resume Next
      .MoveFirst
    While Not .EOF
        Combo1.AddItem !BaDatos
        .MoveNext
      Wend
    End With
   On Error Resume Next
   TRAcampos.MoveLast
   
   While TRAcampos.EOF = False 'TRAcampos.RecordCount > 0
     TRAcampos.Delete
     TRAcampos.MoveLast
   Wend
   Data1.Refresh
   '
   DBGrid1.Columns(0).Visible = False
   DBGrid1.Columns(1).Visible = False
   DBGrid1.Columns(6).Visible = False
   DBGrid1.Columns(3).Width = 500
   DBGrid1.Columns(5).Width = 4300
   DBGrid1.Columns(8).Width = 350
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
   Set CREATABLA = BADASTRU.TableDefs("BACAMPOS")
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("Tabla", dbText, 32)
     .Append CREATABLA.CreateField("Campo", dbText, 32)
     .Append CREATABLA.CreateField("Tipo", dbText, 32)
     .Append CREATABLA.CreateField("Longitud", dbInteger)
     .Append CREATABLA.CreateField("Observa", dbMemo)
     .Append CREATABLA.CreateField("MarBorra", dbInteger)
     .Append CREATABLA.CreateField("NuOrden", dbInteger)
     .Append CREATABLA.CreateField("Idx", dbInteger)
   End With
   '
   ' crea la tabla "TRACAMPOS"
   Set CREATABLA = BADASTRU.CreateTableDef("TRACAMPOS")
   CREATABLA.Fields.Append CREATABLA.CreateField("BaDatos", dbText, 32)
   BADASTRU.TableDefs.Append CREATABLA
   '
   ' agrega campos a "TRACAMPOS"
   Set CREATABLA = BADASTRU.TableDefs("TRACAMPOS")
   With CREATABLA.Fields
     .Append CREATABLA.CreateField("Tabla", dbText, 32)
     .Append CREATABLA.CreateField("Campo", dbText, 32)
     .Append CREATABLA.CreateField("Tipo", dbText, 32)
     .Append CREATABLA.CreateField("Longitud", dbInteger)
     .Append CREATABLA.CreateField("Observa", dbMemo)
     .Append CREATABLA.CreateField("MarBorra", dbInteger)
     .Append CREATABLA.CreateField("NuOrden", dbInteger)
     .Append CREATABLA.CreateField("Idx", dbInteger)
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

