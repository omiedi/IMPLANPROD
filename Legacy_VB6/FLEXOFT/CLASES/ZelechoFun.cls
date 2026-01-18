VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZelechoFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub ArmarZelechoArmado()
   If TRIM$(UCase$(VENZELECHO.NombreZelechoArmado)) = "DEUDOR12" Then
        VENZELECHO.CamposCFormato = "Cta.Cliente/F10;Razon Social/A48"
        VENZELECHO.Titulo = "Clientes Activos "
        VENZELECHO.Condix = "STAT_CLI >=0 "
        VENZELECHO.Campox = "Nume_cli,Raso_Cli"
        VENZELECHO.Tablax = "DEUDOR12"
        VENZELECHO.Anchox = 7000
        VENZELECHO.CampoOrdenamiento = "RASO_CLI ASC"
   End If
   '
   Call ArmaConsultaRecord(FRMZELECHO2.msf2, FRMZELECHO2.MSF1)
   Call CountRegMsf
   Call CargarInfo(FRMZELECHO2.msf2, FRMZELECHO2.MSF1)
   FRMZELECHO2.Caption = VENZELECHO.Titulo & " (" & VENZELECHO.CantiReg & " Registros)"
'   If VENZELECHO.Anchox > 6800 Then
'      FRMZELECHO2.Width = VENZELECHO.Anchox
'      FRMZELECHO2.MSF1 = VENZELECHO.Anchox - 150
'      FRMZELECHO2.MSF1 = VENZELECHO.Anchox - 150
'   End If
End Sub
Sub ArmaConsultaRecord(msf2 As MSFlexGrid, MSF1 As MSFlexGrid)
    Call CARGA_ENCABEZADO(msf2, VENZELECHO.CamposCFormato, FRMZELECHO2)
    Call CARGA_ENCABEZADO(MSF1, VENZELECHO.CamposCFormato, FRMZELECHO2)
    FRMZELECHO2.Caption = VENZELECHO.Titulo
    SQLX0$ = " SELECT " & VENZELECHO.Campox
    SQLX1$ = " SELECT count(*) "
    SQLX2$ = SQLX2$ + " FROM " & VENZELECHO.Tablax & " WITH(NOLOCK) "
    If TRIM$(VENZELECHO.Condix) <> "" Then SQLX2$ = SQLX2$ + " WHERE " & VENZELECHO.Condix
    VENZELECHO.SelectCount = SQLX1$ & SQLX2$
    SQLORDEN$ = "": If TRIM(VENZELECHO.CampoOrdenamiento) <> "" Then SQLORDEN$ = " ORDER BY " & VENZELECHO.CampoOrdenamiento
    VENZELECHO.SelectCompleta = SQLX0$ & SQLX2$ + SQLORDEN$
End Sub
Sub CountRegMsf()
   Set RS = READSET(VENZELECHO.SelectCount)
   VENZELECHO.CantiReg = XVALO(RS.Fields(0))
   RS.Close
   Set RS = Nothing
End Sub


Sub CargarInfo(msf2 As MSFlexGrid, MSF1 As MSFlexGrid)
    msf2.Rows = 1 ' lo limpio
    msf2.Rows = VENZELECHO.CantiReg + 2
    msf2.Cols = 3
    Set RS = READSET(VENZELECHO.SelectCompleta)
    With RS
      j& = 0
      Do While Not .EOF
         j& = j& + 1
         For i& = 1 To RS.Fields.Count
            msf2.TextMatrix(j&, i&) = RS.Fields(i& - 1)
         Next i&
         .MoveNext
      Loop
    End With
    
End Sub

Private Sub Class_Initialize()
'   If VENZELECHO.Inicializar = 1 Then
'      'SI SE LE PASA NOMBRE DE ZELECHO PREVIAMENTE ARMADO BUSCA POR EL NOMBRE
'      If TRIM$(VENZELECHO.NombreZelechoArmado) <> "" Then Call ArmarZelechoArmado
'   End If

End Sub
