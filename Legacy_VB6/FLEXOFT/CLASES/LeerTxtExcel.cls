VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeerTxtExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'devuelve un objeto Recordset con los datos de la hoja
Public Function Leer_Excel(ByVal PathXls As String, Hoja As String, Optional CAMPODEORDEN$) As ADODB.Recordset

      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset

      Dim cs As String
      Dim SQLX$
      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic
      If UCase$(Right(TRIM$(PathXls), 5)) = ".XLSX" Then
         cs = "DRIVER=Microsoft Excel Driver (*.xlsx);" & "DBQ=" & PathXls
      ElseIf UCase$(Right(TRIM$(PathXls), 4)) = ".XLS" Then
         cs = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & PathXls
      Else
         MsgBox "Formato de Base de Datos Seleccionado No Corresponde a un Archivo Excel " + vbCrLf + "Solo se Adminten Archivos con Extensión .XLS O XLSX"
      End If
      
      Hoja = "[" & Hoja & "$" & "]"
      SQLX$ = "SELECT * FROM " & Hoja
      If TRIM$(CAMPODEORDEN$) <> "" Then SQLX$ = SQLX$ + " ORDER BY " & CAMPODEORDEN$
      rs.Open SQLX$, cs
      Set Leer_Excel = rs
      
      Set rs = Nothing
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function

'devuelve un objeto Recordset con los datos del txt

Public Function LeerTxt(Directorio As String) As ADODB.Recordset
      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cn As ADODB.Connection
      Set cn = New ADODB.Connection
      
      
      cn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                         "DBQ=" & Directorio & ";", "", ""
       
      rs.Open "select * from [archivo#txt]", cn, adOpenStatic, adLockReadOnly, adCmdText
      Set LeerTxt = rs
      
      Set rs = Nothing
      Set cn = Nothing
      
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function

'Public Function Leer_Excel_agrupa(ByVal PathXls As String, Hoja As String) As ADODB.Recordset
'
'      On Error GoTo ErrorFunction
'      Dim SQLX$
'      Dim rs1 As ADODB.Recordset
'      Set rs1 = New ADODB.Recordset
'      Dim cs As String
'
'      rs1.CursorLocation = adUseClient
'      rs1.CursorType = adOpenKeyset
'      rs1.LockType = adLockBatchOptimistic
'
'      cs = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & PathXls
'
'      Hoja = "[" & Hoja & "$" & "]"
'
'      SQLX$ = " GROUP BY " & ProcesaExcel.lblcampo.Caption
'      SQLX$ = SQLX$ + " order by " & ProcesaExcel.lblcampo.Caption
'      rs1.Open "SELECT  " & ProcesaExcel.lblcampo.Caption & " as Campo1, COUNT(*) as Cantidad FROM " & Hoja & SQLX$, cs
'
'      Set Leer_Excel_agrupa = rs1
'      Set rs1 = Nothing
'      Exit Function
'ErrorFunction:
'      MsgBox Err.Description, vbCritical
'      Err.Clear
'End Function

Public Function cerrarconexion()
 End Function

