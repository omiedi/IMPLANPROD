Attribute VB_Name = "ExportaExcelSimpleGrid"
Sub CopyToExcel(InFlexGrid As MSFlexGrid, Nome$, _
      ByVal TextoAdicional$)
  Dim R%, c%, Buf$, LstRow%, LstCol%
  Dim FormatMoney As Boolean
  Dim MyExcel As Excel.Application
  Dim wbExcel As Excel.Workbook
  Dim shExcel As Excel.Worksheet
  On Error Resume Next

  Set MyExcel = GetObject(, "Excel.Application")
  If Err.Number <> 0 Then
 Set MyExcel = CreateObject("Excel.Application")
  End If
  Set wbExcel = MyExcel.Workbooks.Add
  Set shExcel = wbExcel.Worksheets.Add
  shExcel.Name = Nome$
  shExcel.Activate
  LstCol% = 0
  For c% = 0 To InFlexGrid.Cols - 1
 InFlexGrid.COL = c%
 LstRow% = 0
 shExcel.Columns(Chr(Asc("A") + c%)).ColumnWidth = InFlexGrid.ColWidth(c%) / 72
 For R% = 0 To InFlexGrid.Rows - 1
   InFlexGrid.Row = R%
   Err.Clear
   Buf$ = InFlexGrid.TextMatrix(R%, c%)
   If Buf$ <> "" Then
  FormatMoney = False
  If InStr(Buf$, vbCrLf) Then
    Buf$ = StrTran(Buf$, vbCrLf, vbLf)
    Do While Right(Buf$, 1) = vbLf
   Buf$ = Left(Buf$, Len(Buf$) - 1)
    Loop
    shExcel.Range(Chr(Asc("A") + c%)).WrapText = True
  ElseIf Format(CDbl(Buf$), csFormatMoneyZero) = Buf$ Then
    If Err.Number = 0 Then
   Buf$ = Str(CDbl(Buf$))
   FormatMoney = True
    End If
  End If
  If Buf$ <> "" Then
    If InFlexGrid.MergeRow(R%) Then
   For LstCol% = c% To 1 Step -1
     If InFlexGrid.TextMatrix(R%, LstCol% - 1) <> InFlexGrid.TextMatrix(R%, c%) Then
    Exit For
     End If
   Next
   If LstCol% <> c% Then
     shExcel.Range(Chr(Asc("A") + LstCol%) & (R% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).MergeCells = True
     shExcel.Range(Chr(Asc("A") + LstCol%) & (R% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).BorderAround
   End If
    End If
    If InFlexGrid.MergeCol(c%) And LstRow% <> R% Then
   If InFlexGrid.TextMatrix(LstRow%, c%) = InFlexGrid.TextMatrix(R%, c%) Then
     shExcel.Range(Chr(Asc("A") + c%) & (LstRow% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).MergeCells = True
     shExcel.Range(Chr(Asc("A") + c%) & (LstRow% + 1), _
         Chr(Asc("A") + c%) & (R% + 1)).BorderAround
   Else
     LstRow% = R%
   End If
    End If
    shExcel.Range(Chr(Asc("A") + c%) & _
        (R% + 1)).Font.COLOR = InFlexGrid.CellForeColor
    If R% < InFlexGrid.FixedRows Or c% < InFlexGrid.FixedCols Then
   shExcel.Range(Chr(Asc("A") + c%) & _
       (R% + 1)).Font.Bold = True
    shExcel.Range(Chr(Asc("A") + c%) & _
        (R% + 1)).Font.BackColor = 40
    End If
    shExcel.Range(Chr(Asc("A") + c%) & (R% + 1)).Value = Buf$
    If FormatMoney Then
   shExcel.Range(Chr(Asc("A") + c%) & _
       (R% + 1)).NumberFormat = "#,##0.00;#,##0.00;#,##0.00"
    End If
  End If
   End If
 Next
  Next
  If TextoAdicional$ <> "" Then
 ' shExcel.Rows(Str(r%+2)).Delete (xlShiftUp)
 Do While Right(TextoAdicional$, 1) = vbLf
   TextoAdicional$ = Left(TextoAdicional$, _
         Len(TextoAdicional$) - 1)
 Loop
 shExcel.Range("A" & (R% + 2)).Value = TextoAdicional$
  End If
  MyExcel.Visible = True
  Set shExcel = Nothing
  Set wbExcel = Nothing
  Set MyExcel = Nothing
End Sub
Public Function StrTran(Cadena As String, BUSCAR As String, Sustituir As String, Optional Veces As Variant) As String
   Dim Contador As Integer
 
Dim Resultado As String
   Dim Cambios As Integer
 

   Resultado = ""
   Cambios = 0
 
   For Contador = 1 To Len(Cadena)
   If Mid(Cadena, Contador, Len(BUSCAR)) = BUSCAR Then
  
Resultado = Resultado & Sustituir
   If Len(BUSCAR) > 1 Then
     
Contador = Contador + Len(BUSCAR) - 1
   End If
  

   ' si se especifica un nº de cambios determinados
   If Not IsMissing(Veces) Then
     
Cambios = Cambios + 1
   If Cambios = Veces Then
    
Resultado = Resultado & Mid(Cadena, Contador + 1)
    
Exit For
     
End If
End If
   If Len(BUSCAR) > 1 Then
     
Contador = Contador + Len(BUSCAR) - 1
   End If
 
Else
   Resultado = Resultado & Mid(Cadena, Contador, 1)
   End If
   Next
 
   StrTran = Resultado
End Function


