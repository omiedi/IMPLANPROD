VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oWord As Word.Application
'Variable de objeto para hacer referencia al documento
Private Documento As Word.Document

Private Sub Class_Initialize()
   
    'nuevo objeto de tipo Word Application
    Set oWord = New Word.Application
       
    'Añadimos un Nuevo documento de word
    Set Documento = oWord.Documents.Add
    Documento.PageSetup.Orientation = wdOrientLandscape

End Sub

Public Sub NuevaLinea(Texto$, Fuente$, Tamanio, Negrita, alineacion)

    Set Parrafo = Documento.Paragraphs.Add
    Set Rango = Parrafo.Range.Next
    Rango.ParagraphFormat.Alignment = alineacion
    Rango.Font.Name = Fuente$
    Rango.Font.Size = Tamanio
    Rango.Font.Bold = Negrita
    Rango.InsertAfter Texto$

End Sub

Public Sub AgregaTabla(ByRef Parrafo, Columnas$, CantidadFilas&, Optional Fuente$, Optional Tamanio, Optional Negrita)

    Dim TituloColumnas(1 To 30) As String
    Dim CantidadColumnas As Long
    
    Set Parrafo2 = Documento.Paragraphs.Add
    Set Rango = Parrafo2.Range.Next
    If TRIM$(Fuente$) <> "" Then Rango.Font.Name = Fuente$
    If Tamanio > 0 Then Rango.Font.Size = Tamanio
    Rango.Font.Bold = False
    If Negrita = True Then Rango.Font.Bold = True
    
    ' *** OBTIENE LOS TITULOS DE LAS COLUMNAS
    Columnas2$ = Columnas$
    car% = InStr(1, Columnas2$, ";")
    Do While car% > 0
        TituloCol$ = Mid$(Columnas2$, 1, car% - 1)
        Columnas2$ = Mid$(Columnas2$, car% + 1)
        CantidadColumnas = CantidadColumnas + 1
        
        TituloColumnas(CantidadColumnas) = TituloCol$
        car% = InStr(1, Columnas2$, ";")
    Loop
    If CantidadColumnas > 0 Then
        CantidadColumnas = CantidadColumnas + 1
        TituloColumnas(CantidadColumnas) = Columnas2$
    End If
    
    ' *** VALIDA LA CANTIDAD DE FILAS
    If CantidadFilas& < 1 Then CantidadFilas& = 1
    Set Parrafo2 = Documento.Tables.Add(Rango, CantidadFilas& + 1, CantidadColumnas, wdWord9TableBehavior, wdAutoFitContent)
    
    ' *** IMPRIME EL TITULO DE CADA COLUMNA
    For i% = 1 To CantidadColumnas
        Parrafo2.Cell(1, i%).Range.InsertAfter TituloColumnas(i%)
    Next i%
    
    Set Parrafo = Parrafo2
    
End Sub

Public Sub GuardarDocumento(RUTAWORD$)
    
    Documento.SaveAs (RUTAWORD$ & ".DOC")
    
End Sub

Public Sub CloseI()

    Documento.Close
    Set Documento = Nothing
    oWord.Quit
    Set oWord = Nothing
    
End Sub


