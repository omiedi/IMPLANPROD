Attribute VB_Name = "TODOWORD"

Sub AgregaLinea(documento, TEXTO$, FUENTE$, tamanio, negrita, alineacion)

    Set Parrafo = documento.Paragraphs.Add
    Set RANGO = Parrafo.Range.Next
'    Rango.FontName = FUENTE$
'    RANGO.ParagraphFormat.Alignment = alineacion
    RANGO.Font.Size = tamanio
    RANGO.Font.Bold = negrita
'    RANGO.InsertBefore TEXTO$
'    Rango.FontName = FUENTE$
End Sub

Sub CREARTXTDEWORD(Path_Doc As String, PATH_NAME_TXT As String, OK%)
''''  Call CREARTXTDEWORD2(Path_Doc, PATH_NAME_TXT, OK%)

On Error GoTo errSub
    OK% = 1
    'VALIDO SI EXISTE EN ESE CASO LO ELIMINO
    If EXISTE%(PATH_NAME_TXT) > 0 Then
       On Error Resume Next
       Kill PATH_NAME_TXT
       On Error GoTo 0
    End If
    '
    If EXISTE%(PATH_NAME_TXT) > 0 Then
       TXT$ = "NO SE PUDO ELIMINAR EL ARCHIVO:  " & PATH_NAME_TXT & vbCrLf
       TXT$ = TXT$ + "VERIFIQUE SI EL MISMO NO ESTA ABIERTO O PROTEGIDO "
       MsgBox TXT$
       OK% = 0
       Exit Sub
    End If
    '
    'CREO EL ARCHIVO DE ESTE MODO LO CREA VACIO
    Open PATH_NAME_TXT For Output As #5
    Close #5

    'Variable de objeto para poder abrir el Word
    Dim Obj_Word As Word.Application
    'Variable de tipo documento de Word
    Dim obj_Documento As Word.Document
    
    'Creamos la referencia para llamar a la aplicación
    Set Obj_Word = New Word.Application

    'Abrimos el archivo.doc que está en el app.path
    Set obj_Documento = Obj_Word.Documents.Open(Path_Doc)

    With obj_Documento
        'Guardamos el contenido del .Doc en formato txt, _
         es decir solo el texto sin tablas o cualquier _
         objeto que contenga el mismo
'         Dim oWord As Word.ApplicationClass
'Dim oDoc As Word.DocumentClass
'
'oWord = New Word.ApplicationClass()
'oDoc = New Word.DocumentClass()
'oDoc = oWord.Documents.Open("C:\FileToOpen.doc")
'oDoc.SaveAs("C:\MyFile.txt", wdFormatText) 'Here's where you can specify format
'oDoc.Close
'oWord.Quit
'oDoc = Nothing
'oWord = Nothing
        ActiveDocument.SaveAs FileName:=PATH_NAME_TXT, _
                                                FileFormat:=wdFormatText
                                                
    'VIERNES 28 ATENCION ACA PRODUJO ERROR QUE NO ENCONTRABA EL ARCHIVO
    'C:\Documents and Settings\OMAR\Escritorio\INMOBILIARIA\Contratos\PASEO DE LA LUNA CONTRATO DE COMODATO Y DEMAS.1TXT

    End With
    obj_Documento.Close
    Obj_Word.Quit
    Set obj_Documento = Nothing

    'Cerramos la aplicacion
'    Obj_Word.Quit
    'Descargamos las variables de la memoria
    Set Obj_Word = Nothing

Exit Sub

errSub:

MsgBox Err.Description, vbCritical
OK% = 0

On Local Error Resume Next

    Obj_Word.Quit
    Set Obj_Word = Nothing
    Set obj_Documento = Nothing
    On Error GoTo 0
End Sub
Sub CREARTXTDEWORD2(Path_Doc As String, PATH_NAME_TXT As String, OK%)
On Error GoTo errSub
    OK% = 1
    'VALIDO SI EXISTE EN ESE CASO LO ELIMINO
    If EXISTE%(PATH_NAME_TXT) > 0 Then
       On Error Resume Next
       Kill PATH_NAME_TXT
       On Error GoTo 0
    End If
    '
    If EXISTE%(PATH_NAME_TXT) > 0 Then
       TXT$ = "NO SE PUDO ELIMINAR EL ARCHIVO:  " & PATH_NAME_TXT & vbCrLf
       TXT$ = TXT$ + "VERIFIQUE SI EL MISMO NO ESTA ABIERTO O PROTEGIDO "
       MsgBox TXT$
       OK% = 0
       Exit Sub
    End If
    '
    'CREO EL ARCHIVO DE ESTE MODO LO CREA VACIO
    Open PATH_NAME_TXT For Output As #5
    Close #5

    'Variable de objeto para poder abrir el Word
    Dim Obj_Word As Object
    'Variable de tipo documento de Word
    Dim obj_Documento As Word.Document
    
    'Creamos la referencia para llamar a la aplicación
    Set Obj_Word = CreateObject("Word.Basic")

    'Abrimos el archivo.doc que está en el app.path
    Set obj_Documento = Obj_Word.Documents.Open(Path_Doc)

    With obj_Documento
        'Guardamos el contenido del .Doc en formato txt, _
         es decir solo el texto sin tablas o cualquier _
         objeto que contenga el mismo
'         Dim oWord As Word.ApplicationClass
'Dim oDoc As Word.DocumentClass
'
'oWord = New Word.ApplicationClass()
'oDoc = New Word.DocumentClass()
'oDoc = oWord.Documents.Open("C:\FileToOpen.doc")
'oDoc.SaveAs("C:\MyFile.txt", wdFormatText) 'Here's where you can specify format
'oDoc.Close
'oWord.Quit
'oDoc = Nothing
'oWord = Nothing
        ActiveDocument.SaveAs FileName:=PATH_NAME_TXT, _
                                                FileFormat:=wdFormatText
                                                
    'VIERNES 28 ATENCION ACA PRODUJO ERROR QUE NO ENCONTRABA EL ARCHIVO
    'C:\Documents and Settings\OMAR\Escritorio\INMOBILIARIA\Contratos\PASEO DE LA LUNA CONTRATO DE COMODATO Y DEMAS.1TXT

    End With
    obj_Documento.Close
    Obj_Word.Quit
    Set obj_Documento = Nothing

    'Cerramos la aplicacion
'    Obj_Word.Quit
    'Descargamos las variables de la memoria
    Set Obj_Word = Nothing

Exit Sub

errSub:

MsgBox Err.Description, vbCritical
OK% = 0

On Local Error Resume Next

    Obj_Word.Quit
    Set Obj_Word = Nothing
    Set obj_Documento = Nothing
    On Error GoTo 0
End Sub






Function Imprimir(PATH As String, _
                  Optional Visible_Word As Boolean = True) As Boolean


    ' variable de objeto para acceder al Word
    Dim Obj_Word As Object


    ' crea el objeto
    Set Obj_Word = CreateObject("Word.Application")


    ' Visible / No visible
    If Visible_Word Then
        Obj_Word.Visible = True
    Else
        Obj_Word.Visible = False
    End If

    'Abre el documento
    Obj_Word.Documents.Open PATH

    ' Imprime el documento activo con Printout
    If Visible_Word = False Then
      Obj_Word.ActiveDocument.PrintOut
      ' Cierra el documento
       Obj_Word.Quit

    End If
    ' Elimina la referencia
    Set Obj_Word = Nothing
    ' Cierra el documento
    Obj_Word.Quit

    ' Elimina la referencia
    Set Obj_Word = Nothing

    ' retorno
    If Err.Number = 0 Then
        Imprimir = True
    End If


Exit Function

Error_Function:

' error
MsgBox Err.Description
On Error Resume Next
Obj_Word.Quit

Set Obj_Word = Nothing

End Function
''''''
''''''' CommandButton que imprime
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
''''''Private Sub Command1_Click()
''''''    Dim ret As Boolean
''''''
''''''    ' le pasa el documento de word que se va a imprimir
''''''    ret = Imprimir("c:\archivo_word.doc", False)
''''''
''''''    If ret Then
''''''       MsgBox "Ok", vbInformation
''''''    End If
''''''
''''''End Sub
''''''

 Function CloseDoc()
        'ABRE UN DOCUMENTO WORD
        Dim wordobj As Object
        Set wordobj = CreateObject("Word.Basic")
        wordobj.FILEOPEN "C:\Archivos de programa\Microsoft Visual Studio\VB98\DATOS\CONTRATOS_GENERADOS\6-PASEO BOGOTA CONTRATO DE COMODATO Y DEMAS.doc"
        ' VISUALIZACION DEL MISMO OPCIONAL
        wordobj.AppShow
        ' Selects the entire document and makes it Bold.
        'EJEMPLO LO EDITA (SELECCIONA TODO)
        wordobj.EditSelectAll
        wordobj.Bold 1
        ' GRABA Y CIERRA
        wordobj.filesave
        wordobj.FileClose (1)
        wordobj.appclose
        Set wordobj = Nothing
  End Function

Function Buscar_Reemplazar_En_Word(Cadena_Origen As String, _
                                   Cadena_reemplazo As String, _
                                   Path_Word As String, Optional ABRIR%) As Boolean
    
    On Error GoTo Err_Buscar
    
    Dim Obj_Word As Object
    
    'Crea una nueva referencia para poder acceder al documento
    Set Obj_Word = CreateObject("Word.Application")

    ' aplicación de word no visible
    Obj_Word.Visible = True
    
    'Abre el archivo especificado mediante el método Open y algunas opciones
    Obj_Word.Documents.Open FileName:=Path_Word

    
    ' Coloca la selección al comienzo del documento
    Obj_Word.ActiveDocument.Goto What:=wdGoToPage, _
                                         Which:=wdGoToNext, _
                                         Name:="1"
    
    
   
    Obj_Word.Selection.Find.ClearFormatting
    Obj_Word.Selection.Find.Replacement.ClearFormatting
    
    ' Propiedades de la búsqueda y reemplazo
    With Obj_Word.Selection.Find
        .TEXT = Cadena_Origen
        .Replacement.TEXT = Cadena_reemplazo
        .Forward = True
        .Wrap = wdFindAsk
        '.Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        '.MatchWholeWord = True

        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
    End With
    
    ' Ejecuta la búsqueda y reemplazo
'    Set myRange = ActiveDocument.Content
'    myRange.Find.Execute FindText:="hi", ReplaceWith:="hello", REPLACE:=wdReplaceAll
   
    Obj_Word.ActiveWindow.Selection.Find.Execute REPLACE:=wdReplaceAll
    
    
    'Ok
    Buscar_Reemplazar_En_Word = True
    If ABRIR% = 1 Then
      Dim MiDoc As String
      MiDoc = Path_Word
      Obj_Word.Visible = True
    Else
      'Obj_Word.ActiveDocument.Save
      'Cierra el documento abierto ( TRue es para guardar los cambios )
       Obj_Word.ActiveDocument.Close True
    
       'Destruye las variables de objeto y cierra la aplicación de word
        Obj_Word.Quit False
         Set Obj_Word = Nothing
    End If
    Exit Function

'Error
Err_Buscar:
    MsgBox Err.Description, vbCritical
    
    On Error Resume Next
    'cierra la aplicación de word
    Obj_Word.Quit False
    'Elimina la referencia de la memoria
    Set Obj_Word = Nothing
End Function

Function REEMPLA_DATOS2$(DOMICILIX$, FechaDeHoy$, FechaDeRealizacion$, PATH As String)
    '|Junio de 2015|
    '|THORNE 382|
    '|miércoles 17-jun a las  08:45 hs |
    'REEMPLAZO LOS DATOS EN EL WORD DE DESTINO
    x = Buscar_Reemplazar_En_Word("|Junio de 2015|", FechaDeHoy$, PATH, 0)
    x = Buscar_Reemplazar_En_Word("|THORNE 382|", DOMICILIX$, PATH, 0)
    x = Buscar_Reemplazar_En_Word("| miércoles 17-jun a las  08:45 hs |", FechaDeRealizacion$, PATH, 0)
    '
    'REPITO POR QUE ESTA DOS VECES
    x = Buscar_Reemplazar_En_Word("|Junio de 2015|", FechaDeHoy$, PATH, 0)
    x = Buscar_Reemplazar_En_Word("|THORNE 382|", DOMICILIX$, PATH, 0)
    x = Buscar_Reemplazar_En_Word("| miércoles 17-jun a las  08:45 hs |", FechaDeRealizacion$, PATH, 1)

End Function


'Function REEMPLA_DATOS_RESERVA$(FORMU As Form, PATH As String)
'
'    NC% = XVALO(FORMU.Label4&)
'    NL& = XVALO(FORMU.Label9)
'    DESCLOCAL$ = TRIM$(FORMU.Label5)
'    NUMLOCAL$ = TRIM$(FORMU.Label13)
'    HOII% = HOY(HOS$)
'    DIADEHOY$ = FECHATEXLAR$(HOII%)
'    CONDI$ = "NUME_CLI = " & NC%
'    DOCUME$ = VALOADMIN("DEUDOR12", "DOCU_CLI", CONDI$, "NOMESS")
'    TELECLI$ = TRIM$(VALOADMIN("DEUDOR12", "Tele_Cli", CONDI$, "NOMESS"))
'    TELECLI$ = TELECLI$ + " " + TRIM$(VALOADMIN("DEUDOR12", "CELU_CLI", CONDI$, "NOMESS"))
'    IMPORESRVA$ = TRIM$(FORMU.Text1)
'    IMPORESERVA_LETRAS$ = Num2Text(XVALO(IMPORESRVA$))
'    DOMICILIO_LOCAL$ = TRIM$(FORMU.Label6)
'    IMPO_ALQUILER$ = TRIM$(FORMU.Label11)
'    COMISION$ = TRIM$(FORMU.Label11)
'    LLAVE$ = TRIM$(FORMU.Label12)
'    VENCI_Reserva$ = FORMU.DTPicker1.Value
'    x = GALERIA_DE_LOCAL&(XVALO(NUMLOCAL$), DESCRI_GALERIA$)
'
'    'REEMPLAZO LOS DATOS EN EL WORD DE DESTINO
'    x = Buscar_Reemplazar_En_Word("|NOMBRE CLIENTE|", RASOCLI$(NC%), PATH)
'    FORMU.ProgressBar1.Value = 40
'    x = Buscar_Reemplazar_En_Word("|DOMICILIO CLIENTE|", DOMICLI$(NC%), PATH)
'    FORMU.ProgressBar1.Value = 50
'    x = Buscar_Reemplazar_En_Word("|NUMERO LOCAL|", NUMLOCAL$, PATH)
'    FORMU.ProgressBar1.Value = 60
'    x = Buscar_Reemplazar_En_Word("|FECHA DE INICIO|", FEINI$, PATH)
'    FORMU.ProgressBar1.Value = 70
'    x = Buscar_Reemplazar_En_Word("|FIN DE CONTRATO|", FEFINCONT$, PATH)
'    FORMU.ProgressBar1.Value = 80
'    x = Buscar_Reemplazar_En_Word("|FECHA DEL DIA|", DIADEHOY$, PATH)
'    x = Buscar_Reemplazar_En_Word("|DNI CLIENTE|", DOCUME$, PATH)
'    FORMU.ProgressBar1.Value = 90
'    x = Buscar_Reemplazar_En_Word("|TELE CLIENTE|", TELECLI$, PATH)
'    FORMU.ProgressBar1.Value = 100
'    x = Buscar_Reemplazar_En_Word("|IMPORESERVA|", IMPORESRVA$, PATH)
'    FORMU.ProgressBar1.Value = 120
'    x = Buscar_Reemplazar_En_Word("|IMPORESERVA LETRAS|", IMPORESERVA_LETRAS$, PATH)
'    FORMU.ProgressBar1.Value = 130
'    x = Buscar_Reemplazar_En_Word("|DOMICILIO LOCAL|", DOMICILIO_LOCAL$, PATH)
'    FORMU.ProgressBar1.Value = 140
'    x = Buscar_Reemplazar_En_Word("|IMPO ALQUILER|", IMPO_ALQUILER$, PATH)
'    FORMU.ProgressBar1.Value = 150
'    x = Buscar_Reemplazar_En_Word("|COMISION|", COMISION$, PATH)
'    FORMU.ProgressBar1.Value = 160
'    x = Buscar_Reemplazar_En_Word("|LLAVE|", LLAVE$, PATH)
'    FORMU.ProgressBar1.Value = 170
'    x = Buscar_Reemplazar_En_Word("|VENCIRESERVA|", VENCI_Reserva$, PATH)
'    FORMU.ProgressBar1.Value = 180
'    x = Buscar_Reemplazar_En_Word("|TELE CLIENTE|", TELECLI$, PATH)
'    FORMU.ProgressBar1.Value = 190
'    x = Buscar_Reemplazar_En_Word("|GALERIA|", DESCRI_GALERIA$, PATH)
'    FORMU.ProgressBar1.Value = 200
'    '
'
'End Function


