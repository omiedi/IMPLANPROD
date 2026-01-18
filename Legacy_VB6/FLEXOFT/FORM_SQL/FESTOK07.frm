VERSION 5.00
Begin VB.Form FESTOK07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Estadístico de Consumos"
   ClientHeight    =   4890
   ClientLeft      =   390
   ClientTop       =   585
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   120
      Index           =   0
      Left            =   8820
      ScaleHeight     =   60
      ScaleWidth      =   375
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   900
      Left            =   5040
      ScaleHeight     =   840
      ScaleWidth      =   3990
      TabIndex        =   31
      Top             =   210
      Width           =   4050
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         Height          =   650
         Left            =   2500
         Picture         =   "FESTOK07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   70
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   450
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   60
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         Height          =   650
         Left            =   3250
         Picture         =   "FESTOK07.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   70
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Height          =   650
         Left            =   1720
         Picture         =   "FESTOK07.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Re-Calcula Resumen Estadístico de Consumo"
         Top             =   70
         Width           =   645
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   60
         Picture         =   "FESTOK07.frx":0896
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   360
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Listados de Activación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   5025
      TabIndex        =   24
      Top             =   2520
      Width           =   4020
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   945
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   25
         Top             =   1185
         Width           =   2430
         Begin VB.CommandButton Command4 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTOK07.frx":27B8
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Activación de Recepción"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Recepciones Pen- dientes de Aprobar"
            Height          =   750
            Left            =   630
            TabIndex        =   27
            Top             =   75
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   28
         Top             =   660
         Width           =   2430
         Begin VB.CommandButton Command5 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK07.frx":3082
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Activación Compras Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Activa- ción de Compras"
            Height          =   540
            Left            =   675
            TabIndex        =   30
            Top             =   45
            Width           =   1710
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Análisis Gráfico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      TabIndex        =   20
      Top             =   1230
      Width           =   4020
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2790
         TabIndex        =   21
         Top             =   360
         Width           =   2850
         Begin VB.CommandButton Command17 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK07.frx":394C
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Gráfico Estadístico de Consumo Mensual por Pieza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Estadísticas de Consumo por Artículo"
            Height          =   750
            Left            =   630
            TabIndex        =   23
            Top             =   105
            Width           =   2115
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   90
      TabIndex        =   2
      Top             =   2520
      Width           =   4845
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1470
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   9
         Top             =   1470
         Width           =   2745
         Begin VB.CommandButton Command18 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK07.frx":3C56
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Análisis de Disponibilidad"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a Disponibili- dad General de Stock"
            Height          =   540
            Left            =   735
            TabIndex        =   11
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1155
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   3
         Top             =   945
         Width           =   2640
         Begin VB.CommandButton Command15 
            Height          =   600
            Left            =   -20
            Picture         =   "FESTOK07.frx":3F60
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Historial de Plazoz de Entrega"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Análisis de Plazos de Entrega por Artículo"
            Height          =   750
            Left            =   630
            TabIndex        =   5
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   6
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command19 
            Height          =   600
            Left            =   0
            Picture         =   "FESTOK07.frx":4C5A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Cálculo Stock Mínimo Recomendado"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Mínimo Reco- mendado por Pieza"
            Height          =   540
            Left            =   675
            TabIndex        =   8
            Top             =   45
            Width           =   1815
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ultimo Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4845
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Meses"
         Height          =   1065
         Left            =   3465
         TabIndex        =   17
         Top             =   300
         Width           =   1170
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   240
            TabIndex        =   18
            Top             =   315
            Width           =   675
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Período"
         Height          =   1065
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Width           =   3165
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   16
            Top             =   630
            Width           =   1935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            Height          =   225
            Left            =   210
            TabIndex        =   15
            Top             =   690
            Width           =   750
         End
         Begin VB.Label label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   14
            Top             =   315
            Width           =   1935
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            Height          =   225
            Left            =   210
            TabIndex        =   13
            Top             =   380
            Width           =   750
         End
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1275
         TabIndex        =   38
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente Calc.:"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   37
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Cálculo:"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   1545
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   1485
         Width           =   3375
      End
   End
   Begin VB.Menu mnuarchvi 
      Caption         =   "Archivo"
      Begin VB.Menu mnuNuevaAutCompSuc 
         Caption         =   "Nueva Autorización de Pedido de Reposición de Sucursal"
      End
      Begin VB.Menu mnureportemanual 
         Caption         =   "Generación de Reporte Control Estadístico de Consumos vs. O.Compra ult. 18 meses"
      End
      Begin VB.Menu MNU_AUTOMATICO_EXCEL_CONSUMO_ESTADISTICO 
         Caption         =   "ATUOMATICO Generación de Reporte Control Estadístico de Consumos vs. O.Compra ult. 18 meses  "
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Edición 
      Caption         =   "Edición"
      Begin VB.Menu editStockmym 
         Caption         =   "Modificar Stocks Mínimos y Máximos"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
End
Attribute VB_Name = "FESTOK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim DESDEMES%, HASTAMES%
Dim CAMESES%, DESMES$, HASMES$

Sub CARGAR_TABLA_StockSugXSuc()
   '
   Call CARGA_ENCABEZADO(EDITMINMAX.GRID, "NCLIENTE/F0;CODINT/F0;Código/F20;Descripción/A40;St.Min.Act./F9.1;St.Max.Act./F10.1;St.Min.Suc./F10.1;St.Max.Suc/F10.1;St.Min.Sug/F10.1;St.Max.Sug./F10.1;F.Envio/D8;F.Emsion/F0;Usuario/F0;FechaCalculo/F0", EDITMINMAX, , , ANTOT)
   SQLX$ = "SELECT * FROM StockSugXSuc WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE CODINT > 0 "
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       NC = XVALO(!NUCLIEN)
       CI1% = XVALO(!CODINT)
       CODI$ = SAFETEXT$(!COD_EXTE)
       DESCR$ = SAFETEXT$(!Descripcion)
       STOCKMINACT = XVALO(!StockMinimoActual)
       STOCKMAXACT = XVALO(!StockMaximoActual)
       STOCKMINACTSUC = XVALO(!StockMinimoActualEnSuc)
       STOCKMAXACTSUC = XVALO(!StockMaximoActualEnSuc)
       STOCKMINSUG = XVALO(!StockMinimoSugerido)
       STOCKMAXSUG = XVALO(!StockMaximoSugerido)
       FEULTENV% = CVINT(!FechaUltEnvio)
       FEMI% = CVINT(!FechaEmision)
       US% = XVALO(!USUARIO)
       FECHCALC% = SAFETEXT$(!FECHACALCULO)
       '
       i& = i& + 1
       GRID.Rows = i& + 1
       GRID.TextMatrix(i&, 1) = CI1%
       GRID.TextMatrix(i&, 2) = NC
       GRID.TextMatrix(i&, 3) = CODI$
       GRID.TextMatrix(i&, 4) = DESCR$
       GRID.TextMatrix(i&, 5) = FORMATNUM$(STOCKMINACT, "F10.1")
       GRID.TextMatrix(i&, 6) = FORMATNUM$(STOCKMAXACT, "F10.1")
       GRID.TextMatrix(i&, 7) = FORMATNUM$(STOCKMINACTSUC, "F10.1")
       GRID.TextMatrix(i&, 8) = FORMATNUM$(STOCKMAXACTSUC, "F10.1")
       GRID.TextMatrix(i&, 9) = FORMATNUM$(STOCKMINSUG, "F10.1")
       GRID.TextMatrix(i&, 10) = FORMATNUM$(STOCKMAXSUG, "F10.1")
       GRID.TextMatrix(i&, 11) = FECHATEX$(FEULTENV%)
       GRID.TextMatrix(i&, 12) = FECHATEX$(FEMI%)
       GRID.TextMatrix(i&, 13) = US%
       GRID.TextMatrix(i&, 14) = FECHATEX$(FECHCALC%)
       .MoveNext
     Loop
   End With
   RS.Close: Set RS = Nothing
End Sub

Sub CrearTablaStockSugXSuc()

   Call CREACAMPO("", "StockSugXSuc", "NUCLIEN", 4, 0, 1)
   Call CREACAMPO("", "StockSugXSuc", "Codint", 4, 0, 1)
   Call CREACAMPO("", "StockSugXSuc", "COD_EXTE", 10, 24)
   Call CREACAMPO("", "StockSugXSuc", "Descripcion", 10, 64)
   Call CREACAMPO("", "StockSugXSuc", "StockMinimoActual", 7, 0)
   Call CREACAMPO("", "StockSugXSuc", "StockMaximoActual", 7, 0)

   Call CREACAMPO("", "StockSugXSuc", "StockMinimoActualEnSuc", 7, 0)
   Call CREACAMPO("", "StockSugXSuc", "StockMaximoActualEnSuc", 7, 0)
   Call CREACAMPO("", "StockSugXSuc", "StockMinimoSugerido", 7, 0)
   Call CREACAMPO("", "StockSugXSuc", "StockMaximoSugerido", 7, 0)

   Call CREACAMPO("", "StockSugXSuc", "FechaUltEnvio", 8, 0)
   Call CREACAMPO("", "StockSugXSuc", "FechaEmision", 8, 0)
   Call CREACAMPO("", "StockSugXSuc", "USUARIO", 3, 0)
   Call CREACAMPO("", "StockSugXSuc", "FechaCalculo", 8, 0)
   Call CREACAMPO("", "StockSugXSuc", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("", "StockSugXSuc", "FechaModifItem", 8, 0)
   
End Sub
Sub CreatablaExcepcion()
   Call CREACAMPO("", "Excepcion", "Nume_Ocom", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "NUCLIEN", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "CODINT", 4, 0, 1)
   Call CREACAMPO("", "Excepcion", "CODIGO", 10, 32)
   Call CREACAMPO("", "Excepcion", "FECHA", 8, 0)
   Call CREACAMPO("", "Excepcion", "USUARIO", 3, 0)
   Call CREACAMPO("", "Excepcion", "CantiExcepcion", 7, 0)
   
End Sub

Sub GENERA_ARCHIVO_EXCEL_ALERTA()
   'SI TIENE INFO MANDA EL ARCHIVO CON ENCABEZADO STOCKS SUB-MINMOS SEGUN CONTROL ESTADISTICOS DE CONSUMO
   'SI NO TIENE INFORMACION ENVIAR MAIL INFORMANDO NO SE DETECTARON STOCKS SUB - MINIMOS SEGUN CONSUMOS ESTADISTICOS.
   NombreDeArchivo$ = CONTROL("ESTOCK07", "RUTASUMI")
   If EXISTE%(NombreDeArchivo$) > 0 Then Kill NombreDeArchivo$
   
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   i& = 1
   XLSHEET.Cells(i&, 1) = "Cód.Int."
   XLSHEET.Cells(i&, 2) = "Código"
   XLSHEET.Cells(i&, 3) = "Sugerido"
   XLSHEET.Cells(i&, 4) = "Stock Act."
   XLSHEET.Cells(i&, 5) = "Faltan"
   '
   SQLX$ = "SELECT COD_INTE,MiniSuger, StoFeProceso, FeProceso FROM CONTMINI WITH(NOLOCK) WHERE ALERTA > 0 "
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       i& = i& + 1
       XLSHEET.Cells(i&, 1) = XVALO(!COD_INTE)

       XLSHEET.Cells(i&, 2) = CODEXT$(XVALO(!COD_INTE))
       XLSHEET.Cells(i&, 2).NumberFormat = "@"
       XLSHEET.Cells(i&, 3) = XVALO(!MiniSuger)
       XLSHEET.Cells(i&, 3).NumberFormat = "#,##0.0"
       XLSHEET.Cells(i&, 4) = XVALO(!StoFeProceso)
       XLSHEET.Cells(i&, 4).NumberFormat = "#,##0.0"
       XLSHEET.Cells(i&, 5) = XVALO(!MiniSuger) - XVALO(!StoFeProceso)
       XLSHEET.Cells(i&, 5).NumberFormat = "#,##0.0"
       .MoveNext
     Loop
   End With
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   'Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   'MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   'Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
End Sub



Sub GENERA_STOCK_MINIMOS()
     
     '
     '1.-BORRAR LA TABLA NUEVA
     Call BORRAREGISTROS("CONTMINI", "COD_INTE > 0", REGAF&, "NOMESS")
     '2.-LEO EL ARCHIVO GENERADO CUANDO HACE EL REPORTE DE LOS ITEMS QUE EL SISTEMA CONSIDERA DEBE COMPRARSE
     Dim STOCKMIN(), COMPRA_PEND(), RECEP_PEN()
     FE% = HOY(HO$)
     J& = 0
     ReDim Preserve STOCKMIN(NUMAS%)
     ReDim Preserve COMPRA_PEND(NUMAS%)
     ReDim Preserve RECEP_PEN(NUMAS%)
     FIN& = ULTREG&("STMINREC")
     '
     For i& = 1 To FIN&
        TXT$ = REGLEIDO$("STMINREC", i&)
        CIXI% = CVI(Mid$(TXT$, 1, 2))
        STOCKMIN(CIXI%) = CVS(Mid$(TXT$, 9, 4))
        COMPRA_PEND(CIXI%) = COMPRA_PEND_DE_ENTR(0, CIXI%)
        RECEP_PEN(CIXI%) = CANT_RECEP_PENDIENTE_APROB(0, CIXI%)
        
     Next i&

     'GENERA UN RECORDSET CON LOS CODIGOS QUE SE COMPRARON EN LOS ULTIMOS 18 MESES
     FEDIASANTES% = DIASANTES(HOY(HO$), 540)  'FALTARIA REALIZAR CONSULTA DEL MASTER CON LOS QUE NO EXISTEN EN OCOMDETA DESDE 180 DIAS ATRAS A LA FECHA
  SQLX$ = "SELECT COD_INTE FROM OCOMDETA WITH(NOLOCK)"
       FEDES = FETEXCOR1$(FEDIASANTES%)
     SQLX$ = "SELECT COD_INTE FROM OCOMDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE STAT_OCOM < 9 AND Femi_Ocom >= '" & FEDES & "'"
     SQLX$ = SQLX$ + " GROUP BY COD_INTE ORDER BY COD_INTE"
     Set RS1 = READSET(SQLX$)
     If RS1.EOF = True Then RS1.Close: Set RS1 = Nothing: Exit Sub
     'ABRO LA TABLA CONTMINI
     SQLX$ = "SELECT * FROM CONTMINI WHERE ID > 0 ORDER BY ID DESC"
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

     With RS1
       Do While Not .EOF
          CIXI% = XVALO(!COD_INTE)
          STMISUG = STOCKMIN(CIXI%)
          If STMISUG > 0 Then ' SI TIENEN STOCK MINMOS SUERIDOS
            RS.AddNew
            J& = J& + 1
            RS!ID = J&
            RS!COD_INTE = CIXI%
            RS!COD_EXTE = ""
            RS!MiniSuger = STMISUG
            RS!FeProceso = CVDAT(FE%)
            STOCKAXT = STOCKACT(CIXI%)
            RS!StoFeProceso = STOCKAXT
            
            RS!COMP_PEND = COMPRA_PEND(CIXI%)
            RS!RECE_PEND = RECEP_PEN(CIXI%)
            
            If STOCKAXT < STMISUG Then
              RS!ALERTA = 1
            Else
              RS!ALERTA = 0
            End If
            
            RS.Update
          End If
         .MoveNext
       Loop
     End With
     RS.Close
     RS1.Close
     Set RS = Nothing
     Set RS1 = Nothing
     
    'PROBAR SI SE CONSIDERA CUANDO HAY COMPRAS, O RECEPCIONES PENDIENTES.

End Sub
Sub GENERAR_STMINREC()
   FIN& = ULTREG&("STATSTOK")
   Screen.MousePointer = 1
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("STATSTOK", U&)
     RESTAT$ = XX$
     ANESTO07.MESAN = Left$(XX$, 15)
     TTXX$ = REGBLAN$("STMINREC")
     CIXI% = CODINT%(ANESTO07.MESAN)
     MINIREC = Val(ANESTO07.Text1)
     LREPREC = Val(ANESTO07.Text2)
     '
     Call REPLA(TTXX$, MKI$(CIXI%), 1, 2)
     Call REPLA(TTXX$, UNIMED$(CIXI%), 3, 2)
     Call REPLA(TTXX$, MKS$(Val(ANESTO07.Label13)), 5, 4)
     Call REPLA(TTXX$, MKS$(MINIREC), 9, 4)
     Call REPLA(TTXX$, MKS$(Val(ANESTO07.Label16)), 13, 4)
     Call REPLA(TTXX$, MKS$(LREPREC), 17, 4)
     '
     If CVS(Mid$(TTXX$, 9, 4)) > CVS(Mid$(TTXX$, 5, 5)) + 0.05 Then
       Call REPLA(TTXX$, "Ajustar St.Minimo", 21, 28)
     End If
     '
     JJ& = JJ& + 1
     Call GRAREG("STMINREC", TTXX$, JJ&)
   Next U&
   Call SETULTREG("STMINREC", JJ&)

End Sub

Sub PROCESO_CALCURESU()

    '
'   Call ARMAMES
'   If DESDEFE% < 1 Then Exit Sub
   '
   '\\\OT-08-0862-OD-19/08/08///
   DESDE = FETEXCOR1$(DESDEFE%)
   HASTA = FETEXCOR1$(HASTAFE%)
   '\\\OT-08-0862-OD-FIN///
   '
   '
    
   'ESTE DATO SE CONFIGURA EN MODULO DE REMITO DE TRASLADO DE ARCON
   'DONDE SE ASOCIA EL CLIENTE Y EL DEPOSITO EJEMPLO CLIENTE NEUQUEN Y DEPOSITO ALMACEN NEUQUEN
'   CLIENTE1 = Control("RETRASLA", "SUC1CLIE")
'   CLIENTE2 = Control("RETRASLA", "SUC2CLIE")
'   If XVALO(CLIENTE1) + XVALO(CLIENTE2) < 1 Then GoTo 10
'   DEPORELA% = 0: CLIENTESEL = 0
'   OPX% = COMALTER%("Modalidad de Cálculo: \\Todos \" & Left$(TRIM$(RASOCLI$(XVALO(CLIENTE1))), 16) & "\" & Left$(TRIM$(RASOCLI$(XVALO(CLIENTE2))), 16) & "")
'   If OPX% < 1 Or OPX% > 3 Then Exit Sub
'   If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE1): DEPORELA% = XVALO(Control("RETRASLA", "SUC1DEPO"))
'   If OPX% = 3 Then CLIENTESEL = XVALO(CLIENTE2): DEPORELA% = XVALO(Control("RETRASLA", "SUC2DEPO"))
   '
10 Screen.MousePointer = 11
   Dim TOPED(23, 32767)
   Dim TOCON(23, 32767)
   Dim ACCLI%(32767)
   Dim IVEN(2, 48)
   Dim ACUCO%(2)
   '\\\OT-08-0862-OD-19/08/08///
   ' CANTIDAD DE REGISTROS SEGUN CONSULTA PARA BARRA TRAZADORA
   SQLX$ = "SELECT COUNT(*) AS CANTIREG FROM PEDDETA"
   SQLX$ = SQLX$ + " WHERE (STATUS < 8 OR STATUS IS NULL)" ' MENOS ANULADAS Y CUMPLIDAS
   SQLX$ = SQLX$ + " AND FeSolEnt >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FeSolEnt <= '" & HASTA & "'"
'   If CLIENTESEL > 0 Then SQLX$ = SQLX$ + " AND NroClie = " & CLIENTESEL
   
   Dim CANTEMP As ADODB.Recordset
   Set CANTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CANTEMP
     FIN& = XVALO(!CANTIREG)
   End With
   CANTEMP.Close
   Set CANTEMP = Nothing
   U& = 0
   SQLX$ = "SELECT * FROM PEDDETA"
   SQLX$ = SQLX$ + " WHERE (STATUS < 9 OR STATUS IS NULL)"
   SQLX$ = SQLX$ + " AND FeSolEnt >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FeSolEnt <= '" & HASTA & "'"
   If CLIENTESEL > 0 Then SQLX$ = SQLX$ + " AND NroClie = " & CLIENTESEL
   Dim PEDDETEMP As ADODB.Recordset
   Set PEDDETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With PEDDETEMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        FENTRESOL% = CVINT(!FeSolEnt)
        FEX = FENTRESOL%
        FESOL$ = FECHATEX(FEX)
        ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
        CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
        CI% = XVALO(!CodiInt)
        If CI% > 0 And CI% <= NUMAS% Then
           caped = XVALO(!CANPED)
           TOPED(CODIPE% - DESDEMES%, CI%) = TOPED(CODIPE% - DESDEMES%, CI%) + caped
           ACCLI%(CI%) = 1
        End If
        .MoveNext
      Loop
   End With
   PEDDETEMP.Close
   Set PEDDETEMP = Nothing
   FIN& = 0
   
   '
'   FIN& = ULTREG&("PEDDETA")
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     X$ = REGLEIDO$("PEDDETA", U&)
'     '        NC% = CVI(Mid$(X$, 7, 2))
'     If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
'       FENTRESOL% = CVI(Mid$(X$, 25, 2))
'       If FENTRESOL% >= DESDEFE% Then
'         If FENTRESOL% <= HASTAFE% Then
'           FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
'           ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
'           CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
'           CI% = CVI(Mid$(X$, 33, 2))
'           If CI% > 0 And CI% <= NUMAS% Then
'             CAPED = CVS(Mid$(X$, 79, 4))
'             TOPED(CODIPE% - DESDEMES%, CI%) = TOPED(CODIPE% - DESDEMES%, CI%) + CAPED
'             ACCLI%(CI%) = 1
'           End If
'         End If
'       End If
'     End If
'   Next U&
'   Call CIERRARCH("PEDDETA")

   ' CANTIDAD DE REGISTROS SEGUN CONSULTA PARA BARRA TRAZADORA

   FIN& = 0
   SQLX$ = "SELECT COUNT(*) AS CANTIREG FROM MOVISTO"
   If DEPORELA% > 0 Then
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'RC' AND CANTINGRE > 0 " ' SOLO BR
   Else
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'OF'" ' SOLO OF
   End If
   'SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF'"
   SQLX$ = SQLX$ + " AND FechMov >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FechMov <= '" & HASTA & "'"
   If DEPORELA% > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPORELA% ' FILTRO POR DEPOSITO EN EL CASO QUE HAYA ELEGIDO.
   '
   Dim CANTMOVMP As ADODB.Recordset
   Set CANTMOVMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CANTMOVMP
     FIN& = XVALO(!CANTIREG)
   End With
   CANTMOVMP.Close
   Set CANTMOVMP = Nothing
   U& = 0
   '
   NMS% = NUMAS%
   SQLX$ = "SELECT * FROM MOVISTO"
   If DEPORELA% > 0 Then
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'RC' AND CANTINGRE > 0 " ' SOLO BR
   Else
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'OF'" ' SOLO OF
   End If
   SQLX$ = SQLX$ + " AND FechMov >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FechMov <= '" & HASTA & "'"
   If DEPORELA% > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPORELA% ' FILTRO POR DEPOSITO EN EL CASO QUE HAYA ELEGIDO.
   Dim MOVISTEMP As ADODB.Recordset
   Set MOVISTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVISTEMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '
        If DEPORELA% > 0 Then ' PARA QUE CONSIDERE LOS INGRESOS A SUCURSAL
          caped = XVALO(!CantIngre)
        Else
          caped = XVALO(!Cantsalid) 'ESTO ES CODIGO ORIGINAL
        End If
        '
        FENTRESOL% = CVINT(!fechMov)
        CI% = XVALO(!COD_INTE)
         FEX = FENTRESOL%
         FESOL$ = FECHATEX$(FEX)
         ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
         CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
         If CI% > 0 And CI% <= NMS% Then
            TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + caped
            ACCLI%(CI%) = 1
         End If
       .MoveNext
      Loop
   End With
   MOVISTEMP.Close
   Set MOVISTEMP = Nothing
   U& = 0
   '\\\OT-08-0862-OD-FIN///

   '
'   FIN& = ULTREG&("MOVISTO")
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     X$ = REGLEIDO$("MOVISTO", U&)
'     '        NC% = CVI(Mid$(X$, 7, 2))
'     FENTRESOL% = CVI(Mid$(X$, 1, 2))
'     If FENTRESOL% >= DESDEFE% Then
'       If FENTRESOL% <= HASTAFE% Then
'         CAPED = CVD(Mid$(X$, 96, 8))
'         If CAPED > 0.01 Then
'           If Mid$(X$, 23, 2) = "OF" Then
'             FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
'             ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
'             CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
'             CI% = CVI(Mid$(X$, 3, 2))
'             If CI% > 0 And CI% <= NUMAS% Then
'               TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + CAPED
'               ACCLI%(CI%) = 1
'             End If
'           End If
'         End If
'       End If
'     End If
'   Next U&
'   Call CIERRARCH("MOVISTO")
   '
   REBLA$ = REGBLAN$("STATSTOK")
   JJ& = 0
   FIN& = NUMAS%
   
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     'X$ = REGLEIDO$("INVERT", U& + 1)
     CI% = U&    '  CVI(Mid$(X$, 17, 2))
     If CI% > 0 Then
       If ACCLI%(CI%) > 0 Then
         Y$ = REBLA$
         Call REPLA(Y$, AJUSTI$(CODEXT$(CI%), 16), 1, 16)
         Call REPLA(Y$, DESCRIT0$(CI%), 17, 48)
         Y0$ = "": Z0$ = ""
         For UKI% = 0 To 23
           Y0$ = Y0$ + MKS$(TOPED(UKI%, CI%))
           Z0$ = Z0$ + MKS$(TOCON(UKI%, CI%))
         Next UKI%
         '
          
         Y1$ = "": Z1$ = ""
         SUCOVEN = 0: SUCOFAB = 0
         SUVENI = 0: SUVEND = 0
         SUFABI = 0: SUFABD = 0
         SUTOTI = 0: SUTOTD = 0
         MAXIVEN = 0: MAXIPRO = 0: MAXITOT = 0
         MINIVEN = 9999999: MINIFAB = 9999999: MINITOT = 9999999
         For UJ% = 1 To CAMESES%
           '
           COVEN = CVS(Mid$(Y0$, 4 * UJ% - 3, 4))
           SUCOVEN = SUCOVEN + COVEN
           If UJ% <= CAMESES% / 2 Then SUVENI = SUVENI + COVEN Else SUVEND = SUVEND + COVEN
           'If COVEN > MAXIVEN Then MAXIVEN = COVEN: MINIVEN = COVEN
           'If COVEN < MINIVEN Then MINIVEN = COVEN
           If COVEN > MAXIVEN Then
                MAXIVEN = COVEN
                If MINIVEN = 0 Then MINIVEN = COVEN
           End If
           If COVEN < MINIVEN And COVEN > 0 Then MINIVEN = COVEN
           '
           COFAB = CVS(Mid$(Z0$, 4 * UJ% - 3, 4))
           SUCOFAB = SUCOFAB + COFAB
           If UJ% <= CAMESES% / 2 Then SUFABI = SUFABI + COFAB Else SUFABD = SUFABD + COFAB
           'If COFAB > MAXIPRO Then MAXIPRO = COFAB
           'If COFAB < MINIFAB Then MINIFAB = COFAB
           If COFAB > MAXIPRO Then
                MAXIPRO = COFAB
                If MINIFAB = 0 Then MINIFAB = COFAB
           End If
           If COFAB < MINIFAB And COFAB > 0 Then MINIFAB = COFAB
           '
           If COVEN + COFAB > MAXITOT Then MAXITOT = COVEN + COFAB
           If COVEN + COFAB < MINITOT Then MINITOT = COVEN + COFAB
           If UJ% <= CAMESES% / 2 Then SUTOTI = SUTOTI + COVEN + COFAB Else SUTOTD = SUTOTD + COVEN + COFAB
         Next UJ%
         '
         ESTIM = Int((4 * SUVEND - 2 * SUVENI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOVEN / CAMESES% Then ESTIM = SUCOVEN / CAMESES%
         ESTIVE = ESTIM
         Y1$ = MKS$(MINIVEN) + MKS$(MAXIVEN) + MKS$(SUCOVEN / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = Int((4 * SUFABD - 2 * SUFABI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOFAB / CAMESES% Then ESTIM = SUCOFAB / CAMESES%
         ESTIFA = ESTIM
         Z1$ = MKS$(MINIFAB) + MKS$(MAXIPRO) + MKS$(SUCOFAB / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = Int((4 * SUTOTD - 2 * SUTOTI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < ESTIVE + ESTIFA Then ESTIM = ESTIVE + ESTIFA
         T1$ = MKS$(MINITOT) + MKS$(MAXITOT) + MKS$((SUCOVEN + SUCOFAB) / CAMESES%) + MKS$(ESTIM)
         '
         Call REPLA(Y$, T1$, 65, 16)
         Call REPLA(Y$, Y0$, 129, 96)
         Call REPLA(Y$, Y1$, 225, 16)
         Call REPLA(Y$, Z0$, 257, 96)
         Call REPLA(Y$, Z1$, 353, 16)
         JJ& = JJ& + 1
'         Debug.Print Y$
         Call GRAREG("STATSTOK", Y$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("STATSTOK", JJ&)
   Call FILESORT("STATSTOK", "", 1, 1)
   '
   TTXX$ = Space$(40)
   Call REPLA(TTXX$, DESMES$, 1, 16)
   Call REPLA(TTXX$, HASMES$, 17, 16)
   Call REPLA(TTXX$, TRIM$(Str$(CAMESES%)), 33, 2)
   Call GRACONTROL("ESTACONS", "PERICALC", TTXX$)
   FESTOK07.Label9 = TRIM$(HASMES$)
   '
   TTXX$ = Format(FileDateTime(LUDAT$ + "STATSTOK.DAT"), "dd/mm/yy hh:mm")
   FEX = FECHANUM(Left$(TTXX$, 8))
   FESTOK07.Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
   FESTOK07.label1 = TRIM$(DESMES$)
   FESTOK07.Label9 = TRIM$(HASMES$)
   FESTOK07.Label12 = TRIM$(Str$(CAMESES%))
   'GUARDA LA INFORMACIN DEL ULTIMO CALCULO
'   If CLIENTESEL > 0 Then
'      Label11 = RASOCLI$(CLIENTESEL)
'      ClienteCalculado$ = CLIENTESEL
'   Else
'      Label11 = "TODOS"
'      ClienteCalculado$ = ""
'   End If
   Call GRACONTROL("FESTOCK7", "CLIECALC", "")
   '
90 Screen.MousePointer = 1
   
   '
End Sub


'Dim CORRESP$, ABREMES$, CORRESQ$
'Dim DESFE$, HASFE$
'Dim CODIMES%
'
'Dim GRCC%(32767)
'Dim GRCV%(32767)
'Dim COREGIS#(32767)
'
'Dim MONECO%      ' 1=Pesos, 2=Dolar
'
Private Sub Command15_Click()
   Command15.Enabled = False
   Call COPYSTRU("COPLAENT", "LIPLAENT")
   Call FILESORT("COPLAENT", "LIPLAENT", 1, 3, "ASC")
   Call FINAL("*LIPLAENT")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   If XVALO(CONTROL("FESTOCK7", "CLIECALC")) > 0 Then
       ANESTO07.Label4(4) = "Venta"
       ANESTO07.Label3(4) = "Fábrica"
   Else
       ANESTO07.Label4(4) = "Fábrica"
       ANESTO07.Label3(4) = "Venta"
   End If
   
   ANESTO07.Show 1
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   DISPSTK07.Show 1
   Command18.Enabled = True
End Sub

 Sub Command19_Click()
   Command19.Enabled = False
   RETECO$ = ANESTO07.MESAN
   ATENTI% = 1
   FIN& = ULTREG&("STATSTOK")
   Screen.MousePointer = 1
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("STATSTOK", U&)
     RESTAT$ = XX$
     ANESTO07.MESAN = Left$(XX$, 15)
     TTXX$ = REGBLAN$("STMINREC")
     CIXI% = CODINT%(ANESTO07.MESAN)
     MINIREC = Val(ANESTO07.Text1)
     LREPREC = Val(ANESTO07.Text2)
     '
     If CIXI% = 66 Then
     AA = BB
     End If
     Call REPLA(TTXX$, MKI$(CIXI%), 1, 2)
     Call REPLA(TTXX$, UNIMED$(CIXI%), 3, 2)
     Call REPLA(TTXX$, MKS$(Val(ANESTO07.Label13)), 5, 4)
     Call REPLA(TTXX$, MKS$(MINIREC), 9, 4)
     Call REPLA(TTXX$, MKS$(Val(ANESTO07.Label16)), 13, 4)
     Call REPLA(TTXX$, MKS$(LREPREC), 17, 4)
     '
     If CVS(Mid$(TTXX$, 9, 4)) > CVS(Mid$(TTXX$, 5, 5)) + 0.05 Then
       Call REPLA(TTXX$, "Ajustar St.Minimo", 21, 28)
     End If
     '
     JJ& = JJ& + 1
     Call GRAREG("STMINREC", TTXX$, JJ&)
   Next U&
   Call SETULTREG("STMINREC", JJ&)
   ATENTI% = 0
   ANESTO07.MESAN = RETECO$
   '
   Call FINAL("*STMINREC")
   Command19.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   ESTOSETU.Show 1
   Command26.Enabled = True
End Sub



Sub PROCESO_ARMAMES(Optional MESHASTA$)
   '
   If TRIM$(MESHASTA$) = "" Then
      MESANANOX = TRIM$(CONTROL("ESTOCK07", "MESANOSE"))
   Else
       MESANANOX = MESHASTA$
   End If
   CAMESES% = XVALO(CONTROL("ESTOCK07", "CANTMESE"))

   If TRIM$(MESANANOX) = "" Then
     Exit Sub
   End If
   If CAMESES% < 1 Then
     Exit Sub
   End If

   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(MESANANOX)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Exit Sub
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   HASTAMES% = 12 * ANIO% + MESO%: HASMES$ = ABREMESAN$(HASTAMES%, 1)
   '
   'CAMESES% = Val(OPCALS07.label1)
   '
   For UK% = 1 To CAMESES% - 1
     MESO% = MESO% - 1
     If MESO% < 1 Then
       MESO% = MESO% + 12
       ANIO% = ANIO% - 1
     End If
   Next UK%
   DESDEMES% = 12 * ANIO% + MESO%: DESMES$ = ABREMESAN$(DESDEMES%, 1)
   MESOX$ = Right$("00" + TRIM$(Str$(MESO%)), 2)
   ANIOX$ = Right$("00" + TRIM$(Str$(ANIO% Mod 100)), 2)
   DESDEFE% = FECHANUM("01/" + MESOX$ + "/" + ANIOX$)
   '
End Sub


Private Sub Command4_Click()
   Command4.Enabled = False
   Screen.MousePointer = 11
   '
   JJ& = 0: KK& = 0
   For U& = 1 To ULTREG&("BOLREPEN")
     X$ = REGLEIDO$("BOLREPEN", U&)
     NBR& = CVS(Left$(X$, 4))
     If NBR& > 0 Then
       KK& = KK& + 1
       Call GRAREG("BOLREPEN", X$, KK&)
       '
       APRO# = CVD(Mid$(X$, 163, 8))
       RECHA# = CVD(Mid$(X$, 171, 8))
       FALTA# = CVD(Mid$(X$, 179, 8))
       If Abs(APRO#) < 0.005 Then
         If Abs(RECHA#) < 0.005 Then
           If Abs(FALTA#) < 0.005 Then
             Call REPLA(X$, String$(128, 0), 155, 128)
             FEBOX = CVI(Mid$(X$, 81, 2))
             SATRA = (DIENTRE(FEBOX, HOY(HOS$))) / 7
             Call REPLA(X$, MKS$(SATRA), 155, 4)
             '
             JJ& = JJ& + 1
             Call GRAREG("ACTIREPA", X$, JJ&)
           End If
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("BOLREPEN", KK&)
   Call SETULTREG("ACTIREPA", JJ&)
   Call CIERRARCH("IBOLPEN")
   '
   Call FINAL("*ACTIREPA")
   '
   Screen.MousePointer = 1
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   '
   Screen.MousePointer = 11
   HOI1 = HOY(HOS$)
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 8 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         xxy$ = XXX$ + String$(32, 0)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         Call REPLA(xxy$, MKS$(CAPEN), 129, 4)
         FENTRESO = CVI(Mid$(XXX$, 95, 2))
         If FENTRESO <= HOI1 Then
           DIATRA = DIENTRE%(FENTRESO, HOI1)
           Call REPLA(xxy$, MKS$(DIATRA / 7), 149, 4)
           JJ& = JJ& + 1
           Call GRAREG("OPENDETA", xxy$, JJ&)
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   '
   Call FILESORT("OPENDETA", "", 3, 3, "ASC")
   Call FINAL("*LACTICOM")
   '
   Screen.MousePointer = 1
   '
   Command5.Enabled = True
End Sub

Private Sub Command7_Click()

End Sub

Private Sub editStockmym_Click()
    'Call CARGAR_TABLA_StockSugXSuc
    EDITMINMAX.TXTFECHACALCULO.TEXT = Me.Label3
    EDITMINMAX.TXTCLIENTECALCULADO.TEXT = Me.Label11
    EDITMINMAX.TXTCTADECALCULO.TEXT = XVALO(CONTROL("FESTOCK7", "CLIECALC"))
    EDITMINMAX.Show 1
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me)
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call ABRECONEXION
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '

    If EXISTE%(LUDAT$ + "STATSTOK.DAT") > 0 Then
      TTXX$ = Format(FileDateTime(LUDAT$ + "STATSTOK.DAT"), "dd/mm/yy hh:mm")
      FEX = FECHANUM(Left$(TTXX$, 8))
      Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
      CLIENTECALCULO = XVALO(CONTROL("FESTOCK7", "CLIECALC"))
      If CLIENTECALCULO = 0 Then
        Label11 = "TODOS"
      Else
        Label11 = rasocli$(XVALO(CONTROL("FESTOCK7", "CLIECALC")))
      End If
      TTXX$ = CONTROL$("ESTACONS", "PERICALC")
      label1 = TRIM$(Left$(TTXX$, 16))
      Label9 = TRIM$(Mid$(TTXX$, 17, 16))
      Label12 = TRIM$(Mid$(TTXX$, 33, 2))
      If Val(Label12) < 6 Or Val(Label12) > 24 Or (Val(Label12) Mod 2) <> 0 Then Label12 = ""
    End If
    '
    If Left$(APLINVO$, 6) = "CONEST" Then ' VIENE DE STOCKS
       CI1% = Mid$(APLINVO$, 7)
       Y$ = FILTERGETRECORD$("STATSTOK", 1, CODEXT$(CI1%)) ' TRAE EL RESTO DE LOS DATOS QUE ESTAN EN STATSTOCK
       RESTAT$ = Y$                                        ' POR QUE SI NO NO MUESTRA LOS DATOS DE STOCKS Y GRAFICOS
       ANESTO07.MESAN = CODEXT$(CI1%)
       Call Command17_Click
       Call FINAL("")
    End If
    '
    Call CrearTablaStockSugXSuc
    Call CreatablaExcepcion
    
'    ACA GENERAR PROCESO Y LUEGO CERRAR
    If Left$(APLINVO$, 6) = "CONMIN" Then
       'MENU NO VISIBLE
       Call MNU_AUTOMATICO_EXCEL_CONSUMO_ESTADISTICO_Click
       Call FINAL("")
    End If


    Show
    Refresh
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image2_DblClick()
   OPCALCOS.Show 1
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   OPCALS07.label1 = "8"
   OPCALS07.Show 1
   DoEvents
   If Val(OPCALS07.label1) < 1 Then GoTo 99
   '
   Call ARMAMES
   Call CALCURESU
   '
99 Command2.Enabled = True
   '
End Sub

Sub ARMAMES()
   '**** SI SE MODIFICA ESTA RUTINA MODIFICAR TAMBIEN 'PROCESO_ARMAMES' ****
   '===========================================================================
   
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(OPCALS07.MESAN)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Call MENSERR(24, "Falta Indicar Período Mensual")
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   HASTAMES% = 12 * ANIO% + MESO%: HASMES$ = ABREMESAN$(HASTAMES%, 1)
   '
   CAMESES% = Val(OPCALS07.label1)
   '
   For UK% = 1 To CAMESES% - 1
     MESO% = MESO% - 1
     If MESO% < 1 Then
       MESO% = MESO% + 12
       ANIO% = ANIO% - 1
     End If
   Next UK%
   DESDEMES% = 12 * ANIO% + MESO%: DESMES$ = ABREMESAN$(DESDEMES%, 1)
   MESOX$ = Right$("00" + TRIM$(Str$(MESO%)), 2)
   ANIOX$ = Right$("00" + TRIM$(Str$(ANIO% Mod 100)), 2)
   DESDEFE% = FECHANUM("01/" + MESOX$ + "/" + ANIOX$)
   '
End Sub

Sub CALCURESU()
    '
   '*****
   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   '\\\OT-08-0862-OD-19/08/08///
   DESDE = FETEXCOR1$(DESDEFE%)
   HASTA = FETEXCOR1$(HASTAFE%)
   '\\\OT-08-0862-OD-FIN///
   '
   '
    
   'ESTE DATO SE CONFIGURA EN MODULO DE REMITO DE TRASLADO DE ARCON
   'DONDE SE ASOCIA EL CLIENTE Y EL DEPOSITO EJEMPLO CLIENTE NEUQUEN Y DEPOSITO ALMACEN NEUQUEN
   CLIENTE1 = CONTROL("RETRASLA", "SUC1CLIE")
   CLIENTE2 = CONTROL("RETRASLA", "SUC2CLIE")
   If XVALO(CLIENTE1) + XVALO(CLIENTE2) < 1 Then GoTo 10
   DEPORELA% = 0: CLIENTESEL = 0
   OPX% = COMALTER%("Modalidad de Cálculo: \\Todos \" & Left$(TRIM$(rasocli$(XVALO(CLIENTE1))), 16) & "\" & Left$(TRIM$(rasocli$(XVALO(CLIENTE2))), 16) & "")
   If OPX% < 1 Or OPX% > 3 Then Exit Sub
   If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE1): DEPORELA% = XVALO(CONTROL("RETRASLA", "SUC1DEPO"))
   If OPX% = 3 Then CLIENTESEL = XVALO(CLIENTE2): DEPORELA% = XVALO(CONTROL("RETRASLA", "SUC2DEPO"))
   '
10 Screen.MousePointer = 11
   Dim TOPED(23, 32767)
   Dim TOCON(23, 32767)
   Dim ACCLI%(32767)
   Dim IVEN(2, 48)
   Dim ACUCO%(2)
   '\\\OT-08-0862-OD-19/08/08///
   ' CANTIDAD DE REGISTROS SEGUN CONSULTA PARA BARRA TRAZADORA
   SQLX$ = "SELECT COUNT(*) AS CANTIREG FROM PEDDETA"
   SQLX$ = SQLX$ + " WHERE (STATUS < 8 OR STATUS IS NULL)" ' MENOS ANULADAS Y CUMPLIDAS
   SQLX$ = SQLX$ + " AND FeSolEnt >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FeSolEnt <= '" & HASTA & "'"
   If CLIENTESEL > 0 Then SQLX$ = SQLX$ + " AND NroClie = " & CLIENTESEL
   
   Dim CANTEMP As ADODB.Recordset
   Set CANTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CANTEMP
     FIN& = XVALO(!CANTIREG)
   End With
   CANTEMP.Close
   Set CANTEMP = Nothing
   U& = 0
   SQLX$ = "SELECT * FROM PEDDETA"
   SQLX$ = SQLX$ + " WHERE (STATUS < 9 OR STATUS IS NULL)"
   SQLX$ = SQLX$ + " AND FeSolEnt >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FeSolEnt <= '" & HASTA & "'"
   If CLIENTESEL > 0 Then SQLX$ = SQLX$ + " AND NroClie = " & CLIENTESEL
   Dim PEDDETEMP As ADODB.Recordset
   Set PEDDETEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With PEDDETEMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        FENTRESOL% = CVINT(!FeSolEnt)
        FEX = FENTRESOL%
        FESOL$ = FECHATEX(FEX)
        ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
        CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
        CI% = XVALO(!CodiInt)
        If CI% > 0 And CI% <= NUMAS% Then
           caped = XVALO(!CANPED)
           TOPED(CODIPE% - DESDEMES%, CI%) = TOPED(CODIPE% - DESDEMES%, CI%) + caped
           ACCLI%(CI%) = 1
        End If
        .MoveNext
      Loop
   End With
   PEDDETEMP.Close
   Set PEDDETEMP = Nothing
   FIN& = 0
   
   '
'   FIN& = ULTREG&("PEDDETA")
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     X$ = REGLEIDO$("PEDDETA", U&)
'     '        NC% = CVI(Mid$(X$, 7, 2))
'     If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
'       FENTRESOL% = CVI(Mid$(X$, 25, 2))
'       If FENTRESOL% >= DESDEFE% Then
'         If FENTRESOL% <= HASTAFE% Then
'           FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
'           ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
'           CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
'           CI% = CVI(Mid$(X$, 33, 2))
'           If CI% > 0 And CI% <= NUMAS% Then
'             CAPED = CVS(Mid$(X$, 79, 4))
'             TOPED(CODIPE% - DESDEMES%, CI%) = TOPED(CODIPE% - DESDEMES%, CI%) + CAPED
'             ACCLI%(CI%) = 1
'           End If
'         End If
'       End If
'     End If
'   Next U&
'   Call CIERRARCH("PEDDETA")

   ' CANTIDAD DE REGISTROS SEGUN CONSULTA PARA BARRA TRAZADORA

   FIN& = 0
   SQLX$ = "SELECT COUNT(*) AS CANTIREG FROM MOVISTO"
   If DEPORELA% > 0 Then
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'RC' AND CANTINGRE > 0 " ' SOLO BR
   Else
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'OF'" ' SOLO OF
   End If
   'SQLX$ = SQLX$ + " WHERE CodiMovi = 'OF'"
   SQLX$ = SQLX$ + " AND FechMov >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FechMov <= '" & HASTA & "'"
   If DEPORELA% > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPORELA% ' FILTRO POR DEPOSITO EN EL CASO QUE HAYA ELEGIDO.
   '
   Dim CANTMOVMP As ADODB.Recordset
   Set CANTMOVMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With CANTMOVMP
     FIN& = XVALO(!CANTIREG)
   End With
   CANTMOVMP.Close
   Set CANTMOVMP = Nothing
   U& = 0
   '
   NMS% = NUMAS%
   SQLX$ = "SELECT * FROM MOVISTO"
   If DEPORELA% > 0 Then
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'RC' AND CANTINGRE > 0 " ' SOLO BR
   Else
     SQLX$ = SQLX$ + " WHERE substring(DoPriLar,1,2) = 'OF'" ' SOLO OF
   End If
   SQLX$ = SQLX$ + " AND FechMov >= '" & DESDE & "'"
   SQLX$ = SQLX$ + " AND FechMov <= '" & HASTA & "'"
   If DEPORELA% > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPORELA% ' FILTRO POR DEPOSITO EN EL CASO QUE HAYA ELEGIDO.
   Dim MOVISTEMP As ADODB.Recordset
   Set MOVISTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MOVISTEMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        '
        If DEPORELA% > 0 Then ' PARA QUE CONSIDERE LOS INGRESOS A SUCURSAL
          caped = XVALO(!CantIngre)
        Else
          caped = XVALO(!Cantsalid) 'ESTO ES CODIGO ORIGINAL
        End If
        '
        FENTRESOL% = CVINT(!fechMov)
        CI% = XVALO(!COD_INTE)
         FEX = FENTRESOL%
         FESOL$ = FECHATEX$(FEX)
         ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
         CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
         If CI% > 0 And CI% <= NMS% Then
            TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + caped
            ACCLI%(CI%) = 1
         End If
       .MoveNext
      Loop
   End With
   MOVISTEMP.Close
   Set MOVISTEMP = Nothing
   U& = 0
   '\\\OT-08-0862-OD-FIN///

   '
'   FIN& = ULTREG&("MOVISTO")
'   For U& = 1 To FIN&
'     Call TRACE(20, U&, FIN&)
'     X$ = REGLEIDO$("MOVISTO", U&)
'     '        NC% = CVI(Mid$(X$, 7, 2))
'     FENTRESOL% = CVI(Mid$(X$, 1, 2))
'     If FENTRESOL% >= DESDEFE% Then
'       If FENTRESOL% <= HASTAFE% Then
'         CAPED = CVD(Mid$(X$, 96, 8))
'         If CAPED > 0.01 Then
'           If Mid$(X$, 23, 2) = "OF" Then
'             FEX = FENTRESOL%: FESOL$ = FECHATEX$(FEX)
'             ANIOX% = Val(Mid$(FESOL$, 7, 2)): If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
'             CODIPE% = 12 * ANIOX% + Val(Mid$(FESOL$, 4, 2))
'             CI% = CVI(Mid$(X$, 3, 2))
'             If CI% > 0 And CI% <= NUMAS% Then
'               TOCON(CODIPE% - DESDEMES%, CI%) = TOCON(CODIPE% - DESDEMES%, CI%) + CAPED
'               ACCLI%(CI%) = 1
'             End If
'           End If
'         End If
'       End If
'     End If
'   Next U&
'   Call CIERRARCH("MOVISTO")
   '
   REBLA$ = REGBLAN$("STATSTOK")
   JJ& = 0
   FIN& = NUMAS%
   
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     'X$ = REGLEIDO$("INVERT", U& + 1)
     CI% = U&    '  CVI(Mid$(X$, 17, 2))
     If CI% > 0 Then
       If ACCLI%(CI%) > 0 Then
         Y$ = REBLA$
         Call REPLA(Y$, AJUSTI$(CODEXT$(CI%), 16), 1, 16)
         Call REPLA(Y$, DESCRIT0$(CI%), 17, 48)
         Y0$ = "": Z0$ = "": A = 0
         For UKI% = 0 To 23
           Y0$ = Y0$ + MKS$(TOPED(UKI%, CI%))
           Z0$ = Z0$ + MKS$(TOCON(UKI%, CI%))
           A = A + TOCON(UKI%, CI%)
         Next UKI%
         '
         Y1$ = "": Z1$ = ""
         SUCOVEN = 0: SUCOFAB = 0
         SUVENI = 0: SUVEND = 0
         SUFABI = 0: SUFABD = 0
         SUTOTI = 0: SUTOTD = 0
         MAXIVEN = 0: MAXIPRO = 0: MAXITOT = 0
         MINIVEN = 9999999: MINIFAB = 9999999: MINITOT = 9999999
         For UJ% = 1 To CAMESES%
           '
           COVEN = CVS(Mid$(Y0$, 4 * UJ% - 3, 4))
           SUCOVEN = SUCOVEN + COVEN
           If UJ% <= CAMESES% / 2 Then SUVENI = SUVENI + COVEN Else SUVEND = SUVEND + COVEN
           'If COVEN > MAXIVEN Then MAXIVEN = COVEN: MINIVEN = COVEN
           'If COVEN < MINIVEN Then MINIVEN = COVEN
           If COVEN > MAXIVEN Then
                MAXIVEN = COVEN
                If MINIVEN = 0 Then MINIVEN = COVEN
           End If
           If COVEN < MINIVEN And COVEN > 0 Then MINIVEN = COVEN
           '
           COFAB = CVS(Mid$(Z0$, 4 * UJ% - 3, 4))
           SUCOFAB = SUCOFAB + COFAB
           If UJ% <= CAMESES% / 2 Then SUFABI = SUFABI + COFAB Else SUFABD = SUFABD + COFAB
           'If COFAB > MAXIPRO Then MAXIPRO = COFAB
           'If COFAB < MINIFAB Then MINIFAB = COFAB
           If COFAB > MAXIPRO Then
                MAXIPRO = COFAB
                If MINIFAB = 0 Then MINIFAB = COFAB
           End If
           If COFAB < MINIFAB And COFAB > 0 Then MINIFAB = COFAB
           '
           If COVEN + COFAB > MAXITOT Then MAXITOT = COVEN + COFAB
           If COVEN + COFAB < MINITOT Then MINITOT = COVEN + COFAB
           If UJ% <= CAMESES% / 2 Then SUTOTI = SUTOTI + COVEN + COFAB Else SUTOTD = SUTOTD + COVEN + COFAB
         Next UJ%
         '
         ESTIM = Int((4 * SUVEND - 2 * SUVENI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOVEN / CAMESES% Then ESTIM = SUCOVEN / CAMESES%
         ESTIVE = ESTIM
         Y1$ = MKS$(MINIVEN) + MKS$(MAXIVEN) + MKS$(SUCOVEN / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = Int((4 * SUFABD - 2 * SUFABI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < SUCOFAB / CAMESES% Then ESTIM = SUCOFAB / CAMESES%
         ESTIFA = ESTIM
         Z1$ = MKS$(MINIFAB) + MKS$(MAXIPRO) + MKS$(SUCOFAB / CAMESES%) + MKS$(ESTIM)
         '
         ESTIM = Int((4 * SUTOTD - 2 * SUTOTI) / CAMESES% + 0.5)
         If ESTIM < 0 Then ESTIM = 0
         If ESTIM < ESTIVE + ESTIFA Then ESTIM = ESTIVE + ESTIFA
         T1$ = MKS$(MINITOT) + MKS$(MAXITOT) + MKS$((SUCOVEN + SUCOFAB) / CAMESES%) + MKS$(ESTIM)
         '
         Call REPLA(Y$, T1$, 65, 16)
         Call REPLA(Y$, Y0$, 129, 96)
         Call REPLA(Y$, Y1$, 225, 16)
         Call REPLA(Y$, Z0$, 257, 96)
         Call REPLA(Y$, Z1$, 353, 16)
         JJ& = JJ& + 1
'         Debug.Print Y$
         Call GRAREG("STATSTOK", Y$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("STATSTOK", JJ&)
   Call FILESORT("STATSTOK", "", 1, 1)
   '
   TTXX$ = Space$(40)
   Call REPLA(TTXX$, DESMES$, 1, 16)
   Call REPLA(TTXX$, HASMES$, 17, 16)
   Call REPLA(TTXX$, TRIM$(Str$(CAMESES%)), 33, 2)
   Call GRACONTROL("ESTACONS", "PERICALC", TTXX$)
   Label9 = TRIM$(HASMES$)
   '
   TTXX$ = Format(FileDateTime(LUDAT$ + "STATSTOK.DAT"), "dd/mm/yy hh:mm")
   FEX = FECHANUM(Left$(TTXX$, 8))
   Label3 = DSEMIN$(FEX) + " " + Left$(TTXX$, 8) + " - " + Mid$(TTXX$, 10) + " hs"
   label1 = TRIM$(DESMES$)
   Label9 = TRIM$(HASMES$)
   Label12 = TRIM$(Str$(CAMESES%))
   'GUARDA LA INFORMACIN DEL ULTIMO CALCULO
   If CLIENTESEL > 0 Then
      Label11 = rasocli$(CLIENTESEL)
      ClienteCalculado$ = CLIENTESEL
   Else
      Label11 = "TODOS"
      ClienteCalculado$ = ""
   End If
   Call GRACONTROL("FESTOCK7", "CLIECALC", ClienteCalculado$)
   '
90 Screen.MousePointer = 1
   '
End Sub
Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   HAS% = HASTAFE%
   DES% = DESDEFE%
   '
   Screen.MousePointer = 11
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      np% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If np% > 0 Then
          If np% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= DES% Then
              If FF% <= HAS% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  Exit Sub
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("BOLRECEP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For K& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
          UKA& = K&
          If CAPEN(K&) > 0 Then
            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
            CAPEN(K&) = CAPEN(K&) - CAREC
            GoTo 19
          End If
        End If
      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   NCOCO& = 0
   For K& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
      NUOCO& = CVS(Left$(X$, 4))
      FEOCO% = CVI(Mid$(X$, 5, 4))
      CIXI% = CVI(Mid$(X$, 9, 2))
      NPRO% = CVI(Mid$(X$, 7, 2))
      '
      Z0$ = REGBLAN$("COPLAENT")
      Call REPLA(Z0$, MKI$(CIXI%), 1, 2)
      Call REPLA(Z0$, Left$(X$, 10), 3, 10)
      '
      FENTREPRO = FEOCO%
      If Len(PUBOL$(K&)) >= 4 Then
         For II& = 1 To Len(PUBOL$(K&)) Step 4
           Z$ = Z0$
           REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
           Y$ = REGLEIDO$("BOLRECEP", REBREC&)
           FERECI = CVI(Mid$(Y$, 81, 2))
           DIATRA% = 0
           If FENTREPRO < FERECI Then
             DIATRA% = DIENTRE%(FENTREPRO, FERECI)
           End If
           FERE% = FERECI
           Call REPLA(Z$, MKI$(FERE%), 13, 2)
           Call REPLA(Z$, MKI$(DIATRA%), 15, 2)
           Call REPLA(Z$, MKS$(DIATRA% / 30), 17, 4)
           '
           NCOCO& = NCOCO& + 1
           Call GRAREG("COPLAENT", Z$, NCOCO&)
         Next II&
      End If
      '
   Next K&
   '
   Call SETULTREG("COPLAENT", NCOCO&)
   Call CIERRARCH("COPLAENT")
   '
   Call FILESORT("COPLAENT", "", 4, 4, "ASC")
   Screen.MousePointer = 1
End Sub



Private Sub MNU_AUTOMATICO_EXCEL_CONSUMO_ESTADISTICO_Click()
   FE% = HOY(HO$)
   MESACTUALX$ = METEXTO$(FE%)

   Call PROCESO_ARMAMES(MESACTUALX$)
   Call PROCESO_CALCURESU
   Call GENERAR_STMINREC
   Call CREATABLA_CONTMINI
   Call GENERA_STOCK_MINIMOS
   Call GENERA_ARCHIVO_EXCEL_ALERTA
End Sub

Private Sub mnuNuevaAutCompSuc_Click()
   CLIENTE1 = CONTROL("RETRASLA", "SUC1CLIE")
   CLIENTE2 = CONTROL("RETRASLA", "SUC2CLIE")
   RASOCLIENTE1$ = TRIM$(rasocli$(XVALO(CLIENTE1)))
   RASOCLIENTE2$ = TRIM$(rasocli$(XVALO(CLIENTE2)))
   If XVALO(CLIENTE1) + XVALO(CLIENTE2) < 1 Then
      Call COMUNI("Falta Identificar Cliente/Sucursal\(Configuración en Setup de Remito de Traslado)")
      Exit Sub
   End If
   
   CLIENTESEL = 0
   OPX% = COMALTER%("Selección de Cliente/Sucursal: \\" & Left$(RASOCLIENTE1$, 16) & "\" & Left$(RASOCLIENTE2$, 16) & "")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   '
   'COMPLETA CASILLEROS DE ENCABEZADO
   If OPX% = 1 Then CLIENTESEL = XVALO(CLIENTE1): AUTOROCO.TXTCLIENTE = CLIENTE1
   If OPX% = 2 Then CLIENTESEL = XVALO(CLIENTE2): AUTOROCO.TXTCLIENTE = CLIENTE2
   
'   AUTOROCO.TXTORDENCOMPRA.Locked = False
'   AUTOROCO.TXTORDENCOMPRA.BackColor = &HFFFFFF
   AUTOROCO.Show 1
End Sub

Private Sub mnureportemanual_Click()
   MESCONFIG$ = CONTROL("ESTOCK07", "MESANOSE")
   FE% = HOY(HO$)
   MESACTUALX$ = METEXTO$(FE%)
   OPX% = COMALTER("Considerar Mes Hasta\\Mes Actual (" & MESACTUALX$ & ")\Mes Configurado (" & MESCONFIG$ & ")")
   MESHASTA$ = MESACTUALX$
   If OPX% = 2 Then MESHASTA$ = MESCONFIG$
   
   Call PROCESO_ARMAMES(MESHASTA$)
   Call PROCESO_CALCURESU
   Call GENERAR_STMINREC
   Call CREATABLA_CONTMINI
   Call GENERA_STOCK_MINIMOS
   OPX% = COMALTER%("Opciones de Trabajo\\Visualizar Reporte\Enviar Mail con Reporte en Excel")
   If OPX% = 1 Then
      Call FINAL("?ALERSUBM")
   ElseIf OPX% = 2 Then
      Call GENERA_ARCHIVO_EXCEL_ALERTA
   End If
End Sub


