Attribute VB_Name = "ABREBASESQL"
Dim miconn As ADODB.Connection
Dim mirs As ADODB.Recordset
Dim strsql As String
   

Private Sub Form_Load()
    cmdCreardatabase.Enabled = False
    cmdCrearTabla.Enabled = False
    cmdInsert.Enabled = False
    cmdAddnew.Enabled = False
    lswvertabla.Visible = False
    
End Sub

Private Sub cmdServidorSql_Click()
Screen.MousePointer = vbHourglass
    Dim oNames As Object
    Dim oSQLApp As Object
    Set oSQLApp = CreateObject("SQLDMO.Application")
    Set oNames = oSQLApp.ListAvailableSQLServers()  'corre la funcion que le trae los nombres de las instancias por defaul trabajo con Local
    List1.Clear
    If oNames.Count <> 0 Then
    For i = 1 To oNames.Count
        List1.AddItem oNames.Item(i)
    Next i
    
    '    txtServidorSeleccionado.Text = oNames.Item(1)
        cmdCreardatabase.Enabled = True
    Else
        MsgBox "No hay instancias de Servidores SQL en esta PC"
    End If
    
    Set oSQLApp = Nothing
    Set oNames = Nothing
Screen.MousePointer = vbNormal
End Sub

Private Sub cmdCreardatabase_Click()
If Len(TRIM(txtServidorSeleccionado.TEXT)) = 0 Then
    MsgBox "Nombre de servidor no valido", vbCritical, "Seleccione Servidor"
Exit Sub
End If

Set miconn = New ADODB.Connection
Set mirs = New ADODB.Recordset
miconn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=master;Data Source=" & txtServidorSeleccionado.TEXT & ";Connect Timeout=999"
miconn.Open
miconn.Execute ("Use master")
Set mirs = miconn.Execute("Select name from sysdatabases where name ='" & TRIM(txtBASE.TEXT) & "'")
If Not (mirs.EOF And mirs.BOF) Then
    miconn.Execute ("drop database " & TRIM(txtBASE.TEXT))
End If
    Screen.MousePointer = vbHourglass
    miconn.Execute ("create database " & TRIM(txtBASE.TEXT)) ' crea la base en el directorio por default del SQLServer con los parametros por default
    VerBases
    Screen.MousePointer = vbNormal
    cmdCrearTabla.Enabled = True
End Sub
Private Sub cmdCrearTabla_Click()
  Dim mirs As ADODB.Recordset
    Set mirs = New ADODB.Recordset
    miconn.Execute ("use " & TRIM(txtBASE.TEXT))
    Screen.MousePointer = vbHourglass
    miconn.Execute ("Create table " & TRIM(txttable.TEXT) & "(table1_id SmallInt IDENTITY(1,1)PRIMARY KEY CLUSTERED,Codigo numeric(9) NOT NULL,Descripcion varchar(50)NOT NULL,fecha as getdate(),usuario As SYSTEM_USER)")
    cmdInsert.Enabled = True
    cmdAddnew.Enabled = True
    Screen.MousePointer = vbNormal
    
End Sub
Private Sub cmdInsert_Click()
 Insertaregistro
 Vertabla
 lswvertabla.Visible = True

End Sub
Private Sub cmdAddnew_Click()
Addnewregistro
Vertabla
lswvertabla.Visible = True

End Sub

Private Sub cmdExit_Click()
Set mirs = Nothing
Set miconn = Nothing

Unload Me
End Sub

Function VerBases()
    lswverbases.ListItems.Clear
    strsql = ("Select name,crdate,filename from sysdatabases")
    Set mirs = miconn.Execute(strsql)
    If Not (mirs.EOF And mirs.BOF) Then
       Do While Not mirs.EOF
          Set l = lswverbases.ListItems.Add(, , mirs("name"))
             With l.ListSubItems
                 .Add , , mirs("crdate")
                 .Add , , mirs("filename")
             End With
            mirs.MoveNext
        Loop
     End If
End Function
Function Insertaregistro()
    strsql = ("insert into table1 (codigo,descripcion) values (4567,'Bicicleta')")
    Set mirs = miconn.Execute(strsql)

End Function
Function Addnewregistro()
Dim myconn As ADODB.Connection
Dim rsttable1 As ADODB.Recordset
Set myconn = New ADODB.Connection
myconn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=" & txtBASE.TEXT & ";Data Source=" & txtServidorSeleccionado.TEXT & ";Connect Timeout=999"
myconn.Open
Set rsttable1 = New ADODB.Recordset
rsttable1.CursorType = adOpenKeyset
rsttable1.LockType = adLockOptimistic
rsttable1.Open "Table1", myconn, , , adCmdTable
rsttable1.AddNew
rsttable1!codigo = 1234
rsttable1!descripcion = "Triciclo"
rsttable1.Update

End Function
Function Vertabla()
lswvertabla.ListItems.Clear
    strsql = ("Select * from ") & TRIM(txttable.TEXT)
    Set mirs = miconn.Execute(strsql)
    If Not (mirs.EOF And mirs.BOF) Then
       Do While Not mirs.EOF
          Set l = lswvertabla.ListItems.Add(, , mirs("table1_id"))
             With l.ListSubItems
                 .Add , , mirs("codigo")
                 .Add , , mirs("descripcion")
                 .Add , , mirs("fecha")
                 .Add , , mirs("usuario")
             End With
            mirs.MoveNext
        Loop
     End If
End Function

Private Sub List1_Click()
txtServidorSeleccionado.TEXT = List1.TEXT

End Sub

