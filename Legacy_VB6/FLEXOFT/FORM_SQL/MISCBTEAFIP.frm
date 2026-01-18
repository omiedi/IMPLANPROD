VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MISCBTESAFIP 
   BackColor       =   &H00C7D9CC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comparar con Portal IVA"
   ClientHeight    =   8250
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   13140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid mfgComprobantesSistema 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6588
      _Version        =   393216
      BackColorFixed  =   16501405
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid mfgComprobantesAfip 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   5530
      _Version        =   393216
      BackColorFixed  =   16501405
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C7D9CC&
      Caption         =   "Comprobantes en Afip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C7D9CC&
      Caption         =   "Comprobantes del Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuRecuperarComprobantesAfip 
         Caption         =   "Recuperar Comprobantes Afip y Comparar"
      End
   End
   Begin VB.Menu mnuConsultasyListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuListadoDferenciasConAfip 
         Caption         =   "Listado de Diferencias con Afip"
      End
      Begin VB.Menu mnuListadoDiferenciasConSistema 
         Caption         =   "Listado de Diferencias con Sistema"
      End
   End
End
Attribute VB_Name = "MISCBTESAFIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PeriodoDesde As Integer
Public PeriodoHasta As Integer

Private Sub Command1_Click()
Command1.Enabled = False
Call IntercaleFilasVacias(mfgComprobantesSistema)
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Call EliminarFilasVacias(mfgComprobantesSistema)
Command2.Enabled = True
End Sub

Private Sub Form_Load()

    Campos$ = "F.Emisión/D8;Comprobante/A18;Cuit/A13;Proveedor/A40;Imp. Total/F14.2;Observaciones/A24"
    Call CARGA_ENCABEZADO(MISCBTESAFIP.mfgComprobantesAfip, Campos$, FRMZELECHO)
    
    Call ABRECABAN

End Sub

Private Sub mnuListadoDferenciasConAfip_Click()

    Dim rg As New ReporteGrilla
    
    rg.TITULO = "LIBRO IVA DIGITAL COMPRAS - DIFERENCIA CON AFIP"
    Campos$ = "F.Emisión/D8;F.Contable/D8;F.Acred/D0;F.Real/D0;Comprobante/A18;Cuit/A13;Proveedor/A20;Imp. Total/F12.2;Observaciones/A24"
    rg.EncabezadoDeColumnas = Campos$
    rg.CodigoReporte = "LIVDIDIF"
    Set rg.Grilla = mfgComprobantesSistema
    rg.MENU1.Visible = True
    rg.MENU1.Caption = "Imprimir"
    
    Call rg.Mostrar
    
End Sub

Private Sub mnuListadoDiferenciasConSistema_Click()

    Dim rg As New ReporteGrilla
    
    rg.TITULO = "LIBRO IVA DIGITAL COMPRAS - DIFERENCIA CON SISTEMA"
    rg.EncabezadoDeColumnas = "F.Emisión/D10;Comprobante/A18;Cuit/A13;Proveedor/A40;Imp. Total/F14.2;Observaciones/A24"
    rg.CodigoReporte = "LIVDIDIF"
    Set rg.Grilla = mfgComprobantesAfip
    rg.MENU1.Visible = True
    rg.MENU1.Caption = "Imprimir"
    
    Call rg.Mostrar
    
End Sub

Private Sub mnuRecuperarComprobantesAfip_Click()
    
    
     OpcionInforme% = COMALTER%("Tipo de Informe\\Portal IVA\Mis Comprobantes")
     If OpcionInforme% < 1 Or OpcionInforme% > 2 Then GoTo 99
     MueveColumna% = 0
     If OpcionInforme% = 2 Then MueveColumna% = 1
     
     On Error GoTo ErrorInstanciaExcel
     Dim objexcek As Excel.Application
     Dim Libro  As Excel.Workbook
     Dim Hoja As Excel.Worksheet
     Set objexcel = New Excel.Application
     On Error GoTo 0
     
     CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
     CommonDialog1.DialogTitle = "Archivo Origen de Datos"
     CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLSX)|*.XLSX|Hojas de Cálculo Excel(*.XLS)|*.XLS"
     CommonDialog1.FilterIndex = 1
     On Error GoTo 99
     CommonDialog1.ShowOpen
     On Error GoTo 0
     '
     DoEvents
     Screen.MousePointer = 11
     '
     Ruta$ = UCase$(CommonDialog1.FileName)
     Set Libro = objexcel.Workbooks.Open(Ruta$)
     Set Hoja = Libro.Sheets(1)
     '
     On Error GoTo ErrorLecturaExcel
     ExtraPeriodo% = 0
     URECORD& = UltimoRegistro(Hoja, 1)
     mfgComprobantesAfip.Rows = 1
     For FILA& = 3 To URECORD&
         '
         FechaAfip = CDate(TRIM$(Hoja.Cells(FILA&, 1)))
         FechaAfipFormatoSistema% = CVINT(FechaAfip)
         If FechaAfipFormatoSistema% >= PeriodoDesde% And FechaAfipFormatoSistema <= PeriodoHasta% Then
            FechaComprobante$ = FECHATEX(CVINT(FechaAfip))
            Tipo$ = TRIM$(Hoja.Cells(FILA&, 2))
            PtoVenta& = CLng(Hoja.Cells(FILA&, 3))
            Numero& = CLng(Hoja.Cells(FILA&, 4))
            Cuit$ = TRIM$(Hoja.Cells(FILA&, 7 + MueveColumna%))
            razonSocial$ = TRIM$(Hoja.Cells(FILA&, 8 + MueveColumna%))
                razonSocial$ = REPLACAR$(razonSocial$, "'", "")
            Moneda$ = TRIM$(Hoja.Cells(FILA&, 10 + MueveColumna%))
            TIPOCAMBIO$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 9 + MueveColumna%), ",", "."))
            ImpoNeto$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 11 + MueveColumna%), ",", "."))
            ImpoNoGravado$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 12 + MueveColumna%), ",", "."))
            ImpoExento$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 13 + MueveColumna%), ",", "."))
            ImporteIva$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 14 + MueveColumna%), ",", "."))
            ImporteTotal$ = TRIM$(REPLACAR$(Hoja.Cells(FILA&, 15 + MueveColumna%), ",", "."))
            
            mfgComprobantesAfip.Rows = mfgComprobantesAfip.Rows + 1
            fil& = mfgComprobantesAfip.Rows - 1
            mfgComprobantesAfip.TextMatrix(fil&, 1) = FechaComprobante$
               ComprobanteConFormato$ = FormateoComprobanteAfip(Tipo$, PtoVenta&, Numero&)
            mfgComprobantesAfip.TextMatrix(fil&, 2) = ComprobanteConFormato$
            mfgComprobantesAfip.TextMatrix(fil&, 3) = Cuit$
            mfgComprobantesAfip.TextMatrix(fil&, 4) = razonSocial$
               ImpTotal# = XVALO(REPLACAR$(ImporteTotal$, ",", "."))
            mfgComprobantesAfip.TextMatrix(fil&, 5) = FORMATNUM(ImpTotal#, "F14.2")
            
            CondicionSql$ = "codicom_lar='" & ComprobanteConFormato$ & "' and cuit='" & Cuit$ & "'"
            If CantRegistros("ivadigco", CondicionSql$) = 0 Then
               
               ConsultaSql$ = "insert into ivadigco (codiempr, fechemisi, codicom_lar, cuit, razonsocial, moneda, tipocambio, imponeto, imponogravado, impoexento, impoiva, impototal) " _
               & " values" _
               & " ('" & CODIEMPR% & "'," _
               & " '" & FETEXCOR1$(FECHANUM(FechaComprobante$)) & "'," _
               & " '" & ComprobanteConFormato$ & "'," _
               & " '" & Cuit$ & "'," _
               & " '" & Left$(razonSocial$, 48) & "'," _
               & " '" & Moneda$ & "'," _
               & " '" & TIPOCAMBIO$ & "'," _
               & " '" & ImpoNeto$ & "'," _
               & " '" & ImpoNoGravado$ & "'," _
               & " '" & ImpoExento$ & "'," _
               & " '" & ImporteIva$ & "'," _
               & " '" & ImporteTotal$ & "')"
               
               FLEXCONN.Execute ConsultaSql$
            End If
         Else
            ExtraPeriodo% = 1
         End If
    Next FILA&
    On Error GoTo 0
    '
    On Error Resume Next
    Set Hoja = Nothing
    Libro.Close
    objexcel.Quit
    On Error GoTo 0
    
    Call CompararComprobantes
    
    If ExtraPeriodo% = 1 Then
        Call COMUNI("Se Han Detectado Comprobantes De AFIP\Fuera del Período Elegido")
    End If
    
99  Screen.MousePointer = 1
    Exit Sub

ErrorInstanciaExcel:
    On Error GoTo 0
    Call COMUNI("Error al Acceder a Microsoft Excel")
    Exit Sub
    
ErrorLecturaExcel:
    Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    Set Hoja = Nothing
    Libro.Close
    objexcel.Quit
    On Error GoTo 0
    
    Call COMUNI("Error al Leer a Microsoft Excel.\ \Verifique Haber Elegido la Modalidad\que Corresponda al Archivo Excel.")
    Exit Sub
    
End Sub

Function FormateoComprobanteAfip(Tipo As String, PuntoVenta As Long, Numero As Long) As String

    n% = InStr(1, Tipo, "-")
    CodigoCbte% = XVALO(Mid$(Tipo, 1, n% - 1))
    
    Select Case CodigoCbte%
        Case 1: COMPROBANTE$ = "FC-A"
        Case 2: COMPROBANTE$ = "ND-A"
        Case 3: COMPROBANTE$ = "NC-A"
        Case 6: COMPROBANTE$ = "FC-B"
        Case 7: COMPROBANTE$ = "ND-B"
        Case 8: COMPROBANTE$ = "NC-B"
        Case 11: COMPROBANTE$ = "FC-C"
        Case 12: COMPROBANTE$ = "ND-C"
        Case 13: COMPROBANTE$ = "NC-C"
        Case 19: COMPROBANTE$ = "FC-E"
        Case 20: COMPROBANTE$ = "ND-E"
        Case 21: COMPROBANTE$ = "NC-E"
        Case 51: COMPROBANTE$ = "FC-M"
        Case 52: COMPROBANTE$ = "ND-M"
        Case 53: COMPROBANTE$ = "NC-M"
        Case 81: COMPROBANTE$ = "TF-A"
        Case 82: COMPROBANTE$ = "TF-B"
    End Select
    
    PtoVenta$ = Right("0000" & CStr(PuntoVenta), 4)
    
    NroCbte$ = Right("00000000" & CStr(Numero), 8)
    
    COMPROBANTE$ = COMPROBANTE$ & "-" & PtoVenta$ & "-" & NroCbte$
    
    FormateoComprobanteAfip = COMPROBANTE$
    
End Function

Function UltimoRegistro(Hoja As Excel.Worksheet, COLUMNA As Integer) As Double
   Dim valor As Variant
   Dim CantFilasVacias&
   CantFilasVacias& = 0
   UltFila& = 0
   For FILA& = 2 To Hoja.Rows.Count
      If CantFilasVacias >= 100 Then
         GoTo 10
      End If
         valor = TRIM$(Hoja.Cells(FILA&, COLUMNA))
         If valor = "" Then 'Si no hay valor
            If UltFila& = 0 Then 'Si no hay registro de ultima fila se registra, si no se mantiene la almacenada
               UltFila& = FILA& - 1
            End If
            CantFilasVacias& = CantFilasVacias& + 1
           Else
            CantFilasVacias& = 0
            UltFila& = 0
          End If
   Next FILA&
10 UltimoRegistro = UltFila&
End Function


Sub CompararComprobantes()
   
    ' *** BLANQUEA AMBAS GRILLAS
    For si = 1 To mfgComprobantesSistema.Rows - 1
         Call COLOREAR_GRILLA_SOLAFILA(mfgComprobantesSistema, RGB(255, 255, 255), si, 0, 1, 0)
    Next si
    
    For af = 1 To mfgComprobantesAfip.Rows - 1
         Call COLOREAR_GRILLA_SOLAFILA(mfgComprobantesAfip, RGB(255, 255, 255), af, 0, 1, 0)
    Next af
    
    Dim CoincideComprobante As Boolean
    Dim CoincideCuit As Boolean

    ' *** COMPARA COMPROBANTES DEL SISTEMA HACIA AFIP
    
    For si = 1 To mfgComprobantesSistema.Rows - 1
    
        CbteSis$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 5))
        CuitSis$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 6))
        ImpTotalSis$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 8))
        
        encontro% = 0
        CoincideComprobante = False
        CoincideCuit = False
        For af = 1 To mfgComprobantesAfip.Rows - 1
        
            CbteAfip$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 2))
            CuitAfip$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 3))
            ImpTotalAfip$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 5))
            
            If CbteSis$ = CbteAfip$ Then
                CoincideComprobante = True
                If CuitSis$ = CuitAfip$ Then
                    CoincideCuit = True
                    If ImpTotalSis$ = ImpTotalAfip$ Then
                        encontro% = 1
                    End If
                End If
            End If
        Next af
        
        If encontro% = 0 Then
            CondicionSql$ = "codicom_lar='" & CbteSis$ & "' and cuit='" & CuitSis$ & "' and abs(impototal-" & ImpTotalSis$ & ")<0.005 and fechemisi<'" & FETEXCOR1$(PeriodoDesde%) & "'"
            FechaAfip% = CVINT(VALOBADA("ivadigco", "fechemisi", CondicionSql$))
            If FechaAfip% > 0 Then
                mfgComprobantesSistema.TextMatrix(si, 9) = "En Afip: " & FECHATEX$(FechaAfip%)
                encontro% = 1
            End If
        End If
        
        If encontro% = 0 Then
            Call COLOREAR_GRILLA_SOLAFILA(mfgComprobantesSistema, RGB(201, 184, 143), si, 0, 0, 0)
            If CoincideComprobante = False Then
                mfgComprobantesSistema.TextMatrix(si, 9) = "Falta Cbte en Afip"
            ElseIf CoincideCuit = False Then
                mfgComprobantesSistema.TextMatrix(si, 9) = "No Coincide CUIT en Cbte en Afip"
            Else
                mfgComprobantesSistema.TextMatrix(si, 9) = "No Coincide Importe en Cbte en Afip"
            End If
        End If
        
    Next si
    
    
    ' *** COMPARA COMPROBANTES DE AFIP HACIA EL SISTEMA

    For af = 1 To mfgComprobantesAfip.Rows - 1
    
        CbteSis$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 2))
        CuitSis$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 3))
        ImpTotalSis$ = TRIM$(mfgComprobantesAfip.TextMatrix(af, 5))
        
        encontro% = 0
        CoincideComprobante = False
        CoincideCuit = False
        For si = 1 To mfgComprobantesSistema.Rows - 1
        
            CbteAfip$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 5))
            CuitAfip$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 6))
            ImpTotalAfip$ = TRIM$(mfgComprobantesSistema.TextMatrix(si, 8))
            
            If CbteSis$ = CbteAfip$ Then
                CoincideComprobante = True
                If CuitSis$ = CuitAfip$ Then
                    CoincideCuit = True
                    If ImpTotalSis$ = ImpTotalAfip$ Then
                        encontro% = 1
                    End If
                End If
            End If
        Next si
               
        If encontro% = 0 Then
            Call COLOREAR_GRILLA_SOLAFILA(mfgComprobantesAfip, RGB(201, 184, 143), af, 0, 0, 0)
            
            If CoincideComprobante = False Then
                mfgComprobantesAfip.TextMatrix(af, 6) = "Falta Cbte en Sistema"
            ElseIf CoincideCuit = False Then
                mfgComprobantesAfip.TextMatrix(af, 6) = "No Coincide CUIT en Cbte en Sistema"
            Else
                mfgComprobantesAfip.TextMatrix(af, 6) = "No Coincide Importe en Cbte en Sistema"
            End If
        End If
        
    Next af
    
End Sub
