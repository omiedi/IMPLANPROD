VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OCOMSUCU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14190
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10515
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Mensajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   19
      Top             =   7440
      Width           =   14175
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Modo Test"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   22
         Top             =   0
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFMEN 
         Height          =   2685
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior"
         Top             =   240
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   4736
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3480
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   10665
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   4560
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "OCOMSUCU.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "OCOMSUCU.frx":0234
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lista de Pedidos de Reposición de Sucursal "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14175
      Begin VB.CommandButton Command2 
         Caption         =   "Maximizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9600
         Picture         =   "OCOMSUCU.frx":0292
         TabIndex        =   14
         ToolTipText     =   "Emite Listado"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exp. a Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7920
         TabIndex        =   13
         Top             =   3480
         Width           =   1125
      End
      Begin VB.Frame Frame1 
         Height          =   600
         Left            =   120
         TabIndex        =   8
         Top             =   3360
         Width           =   7455
         Begin VB.ComboBox CMBFILTRO 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   160
            Width           =   3135
         End
         Begin VB.TextBox TXTBUSCAR 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   9
            Top             =   160
            Width           =   2535
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D8BD05&
            BackStyle       =   0  'Transparent
            Caption         =   "Campo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3360
            TabIndex        =   12
            Top             =   160
            Width           =   705
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D8BD05&
            BackStyle       =   0  'Transparent
            Caption         =   "Filtrar:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   160
            Width           =   585
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar Pedido Seleccionado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11160
         TabIndex        =   7
         ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
         Top             =   3480
         Width           =   1605
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   12840
         TabIndex        =   6
         ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
         Top             =   3480
         Width           =   1125
      End
      Begin MSFlexGridLib.MSFlexGrid GRID 
         Height          =   3165
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Si Tiene Pedido Generado Doble Click Para Modificar Pedido"
         Top             =   240
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF_2 
         Height          =   3165
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   240
         Visible         =   0   'False
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   8421631
         BackColorBkg    =   12632256
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle de Orden de Compra Seleccionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   17
      Top             =   4080
      Width           =   14175
      Begin MSFlexGridLib.MSFlexGrid MSFDET 
         Height          =   3045
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferio"
         Top             =   240
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   5371
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   1000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu S 
         Caption         =   "_"
      End
   End
End
Attribute VB_Name = "OCOMSUCU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long
Public APROBO%
Dim LecturaDatosOriginales As Byte
Dim GLOBALFRM_strContenidoCelda$


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub archivo_Click()

End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub cmdtext_Click()
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   
   Call ImprimirEnExcel(GRID)

99 Command1.Enabled = True

End Sub
Private Sub Command2_Click()
    '
    If Command2.Caption = "Maximizar" Then
        If LecturaDatosOriginales = 0 Then
            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
            GridTop = GRID.Top: GridLeft = GRID.Left: GridWidth = GRID.Width: GridHeight = GRID.Height
            Com1Top = Command1.Top: Com1Top = Command1.Top: Com1LEFT = Command1.Left
            Com2Top = Command2.Top: Com2Top = Command2.Top: Com2Left = Command2.Left

            FRAMETop = Frame1.Top: FRAMELeft = Frame1.Left: FRAMEWidth = Frame1.Width: FRAMEHeight = Frame1.Height
            LecturaDatosOriginales = 1
        End If
    
        QCOL% = GRID.Cols
        ANCHOCOL& = 0
        For COLUM% = 1 To QCOL% - 1
            ANCHOCOL& = ANCHOCOL& + GRID.ColWidth(COLUM%)
        Next COLUM%
        ANCHOSCR& = Screen.Width
        If ANCHOCOL& > ANCHOSCR& Then
            Me.Left = 0: Me.Width = ANCHOSCR&
            GRID.Width = ANCHOSCR& - 150
        Else
            LEFTFORM& = Int((ANCHOSCR& - ANCHOCOL&) / 2)
            Me.Left = LEFTFORM&: Me.Width = ANCHOCOL&
            GRID.Width = ANCHOCOL& - 150
        End If
        '
        QFILAS% = GRID.Rows
        altoFila& = 0
        For Fila% = 1 To QFILAS% - 1
            altoFila& = altoFila& + GRID.RowHeight(Fila%)
        Next Fila%
        ALTOSCR& = Screen.Height
        If altoFila& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            GRID.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (altoFila& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = altoFila& + 1600
            GRID.Height = altoFila& + 600
        End If
        '
        BORDEINF& = GRID.Height + 100
        Command2.Top = BORDEINF&
        Command1.Top = BORDEINF&
        Frame1.Top = BORDEINF&
        Command1.Left = Me.Width - 1800
        Command2.Caption = "Minimizar"
        
    Else
        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
        GRID.Top = GridTop: GRID.Left = GridLeft: GRID.Width = GridWidth: GRID.Height = GridHeight
        Frame1.Top = FRAMETop: Frame1.Left = FRAMELeft: Frame1.Width = FRAMEWidth: Frame1.Height = FRAMEHeight
        Command2.Top = Com2Top: Command2.Top = Com2Top: Command2.Left = Com2Left
        Command1.Top = Com1Top: Command1.Top = Com1Top: Command1.Left = Com1LEFT
        Command2.Caption = "Maximizar"
    End If
    MSF_2.Height = GRID.Height
    MSF_2.Width = GRID.Width
    MSF_2.Top = GRID.Top
    MSF_2.Left = GRID.Left
    
    '
End Sub
Sub GENEXCEL(NombreDeArchivo$, GRID As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #n1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.QUIT      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
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
   XLAPP.QUIT ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Command3_Click()
    Command3.Enabled = False
    'SETEO EL MSF QUE ESTE VISIBLE
    
    Dim MSF As msFlexGrid
    If GRID.Visible = True Then
        Set MSF = GRID
    Else
        Set MSF = MSF_2
    End If
    
    'RECORRO Y GRABO
    For i& = 1 To MSF.Rows - 1
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A1"
    
          If MSF.TextMatrix(i&, 0) = "*" Then
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A2"
          
              ENCA$ = "Sucursal/F10;R.Social/32;O.Compra/F9;Emis.O.C./D8;NC.Pedido/F12;R.Social Cliente de Sucursal/A32;P.Avellaneda/F12.0"
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A3"
             P_Avellaneda& = XVALO(MSF.TextMatrix(i&, 7))
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A4"
             If P_Avellaneda& > 0 Then
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A5"
             
               CONDI$ = "NROPED = " & P_Avellaneda&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A6"
               If CantRegistros&("PEDENCA", CONDI$, "NOMESS") > 0 Then
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A7"
                  Call COMUNI("Orden de Compra Seleccionada Ya Posee Pedido Generado")
                  GoTo 99
               End If
             End If
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A8"
             NUOC& = XVALO(MSF.TextMatrix(i&, 3))
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A9"
             
             NC_Pedido = XVALO(MSF.TextMatrix(i&, 5))
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A10"
             SUCURSAL = XVALO(MSF.TextMatrix(i&, 1))
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A11"
             Call GenerarPedido(SUCURSAL, NUOC&, P_Avellaneda&, NC_Pedido)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "A12"
             
             Exit For
           End If
    Next i&
    APROBO% = 1
         
99  Command3.Enabled = True
    Unload Me
End Sub

Sub GenerarPedido(SUCURSAL, NUOC&, P_Avellaneda&, NC_Pedido)
'Call GENERAR_ARCHIVO_GENERO_PEDIDO(SUCURSAL, 202, 1111)
'Exit Sub
EPAC$ = TRIM$(UCase$(EMPREAC$))
    'SI EL PEDIDO NO EXISTE
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B0"
    
    HOII% = HOY(HOS$)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B1"
    
        'ESTO ES POR QEU SI NO SE LE PASA EL CLIENTE DEL PEDIIOD DE NEUQUEN EL CLIENTE VA A SER NEUQUEN
'        If XVALO(NC_Pedido) < 1 Then
If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then
        NC_Pedido = NC_Pedido ' MANTIENE EL PARAMETRIZADO
Else
        'ESTO ES POR QEU SI NO SE LE PASA EL CLIENTE DEL PEDIIOD DE NEUQUEN EL CLIENTE VA A SER NEUQUEN
        NC_Pedido = SUCURSAL
End If

If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B2"
'        End If
        ORCOMP$ = "PR-" & NUOC&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B3"
        REFE$ = " - " + ORCOMP$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B4"
        REFE$ = REFE$ + " - " + FECHATEX$(HOII%)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B5"
        LOMA% = 48 - Len(REFE$)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B6"
        REFE$ = Left$(TRIM$(rasocli$(NC_Pedido)), LOMA%) + REFE$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B7"
        '

        Dim rs1 As ADODB.Recordset
        Set rs1 = New ADODB.Recordset
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B8"
        SQLX$ = "SELECT * FROM  PEDENCA"
        SQLX$ = SQLX$ + " WHERE NroPED = " & P_Avellaneda&
        rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B9"
        With rs1
          If .EOF Then 'SI NO EXISTE EL PEDIDO DEBE GENERARLOS
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B10"
            .AddNew
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B11"
            On Error Resume Next
            !CODIEMPR = CodiEmp%
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B12"
            On Error GoTo 0
            NroPedix& = NUPEDIDO& ' FUNCION
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B13"
            !NROPED = NroPedix&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B14"
            !REFERENCIA = Left$(REFE$, 48)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B15"
            !FECHEMIS = CVDAT(HOII%)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B16"
            !NROCLIE = NC_Pedido
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B17"
            !RaSoClie = rasocli$(NC_Pedido)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B18"
            !NROOCOM = SAFETEXT$(NUOC&)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B19"
             !DOMIENT = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B20"
            !LOCAENT = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B21"
            !TranspEnt = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B22"
            !NTranspo = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B23"
            !NSucent = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B24"
            !Observa = ""
            
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B25"
            !Status = 0              ' Status  de Aprobación
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B26"
            !ListPre = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B27"
            !CondPag = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B28"
            !VENDED = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B29"
            !TDescApli = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B30"
            .Update
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B31"
          End If
        End With
        '
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B32"
        rs1.Close
        Set rs1 = Nothing
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B33"
        ANOTA$ = ""
        TTXYZ$ = Space$(76)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B34"
        Call REPLA(TTXYZ$, HOS$, 1, 8)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B35"
        Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B36"
        Call REPLA(TTXYZ$, USERTER$, 69, 24)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B37"
        
        ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
        IKITEM% = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B38"
        SQLX$ = "SELECT * FROM OC_SUCURSAL WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE SUCURSAL = " & SUCURSAL
        SQLX$ = SQLX$ + " AND O_COMPRA_SUC =" & NUOC&
        Set RS2 = READSET(SQLX$)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B39"
        With RS2
          Do While Not .EOF
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B40"
            CI1% = XVALO(!CODINT)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B41"
            CANTI = XVALO(!Cantidad)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B42"
            FECHENT% = CVINT(!FECHA_SOLICI_COMPRA)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B43"
            Dim rs3 As ADODB.Recordset
            Set rs3 = New ADODB.Recordset

            SQLX$ = "SELECT * FROM  PEDDETA  "
            SQLX$ = SQLX$ + " WHERE  NroPED = " & P_Avellaneda&
            rs3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B44"

            With rs3
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B45"
                .AddNew
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B46"
                On Error Resume Next
                !CODIEMPR = CodiEmp%
                On Error GoTo 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B47"
                IKITEM% = IKITEM% + 1
                IK$ = TRIM$(str$(IKITEM%))
                CAVERI$ = CAVER$ + "-" + IK$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B48"
                !NROPED = NroPedix&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B49 " & NroPedix&
                !FECHEMIS = CVDAT(HOII%)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B50 "
                !NROCLIE = NC_Pedido
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B51 "
                   !RaSoClie = rasocli$(NC_Pedido)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B52 "
                   !NROOCOM = ORCOMP$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B53 "
                   !FeSolEnt = FETEXCOR1$(FECHENT%) ' LO GENERA GUANDO GRABA EL PEDENCA
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B54 " & NroPedix&
                   !Status = 0              ' Status  de Aprobación
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B55 " & NroPedix&
                   !ListPre = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B56 " & NroPedix&
                   !CondPag = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B57 " & NroPedix&
                   !VENDED = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B58 " & NroPedix&
                   !PorComVend = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B59 " & NroPedix&
                   !NuOrItem = IKITEM%
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B60 " & NroPedix&
                   
                   !CODIINT = CI1%
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B61 " & NroPedix&
                     CODIEXTERNO$ = CODEXT$(CI1%)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B62 " & NroPedix&
                   !CODIEXT = CODIEXTERNO$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B63 " & NroPedix&
                   !SubCod = ""
 If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B64 " & NroPedix&
                    Descr$ = DESCRIT0$(CI1%)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B65 " & NroPedix&
                   !Descrip = Descr$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B66 " & NroPedix&
                   !TLargo = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B67 " & NroPedix&
                   !TAncho = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B68 " & NroPedix&
                   !TEspesor = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B69 " & NroPedix&
                   !CBase = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B70 " & NroPedix&
                   !TTALLE = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B71 " & NroPedix&
                   !TCOLOR = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B72 " & NroPedix&
                   !VenRep = "V"   ' Venta o Reparacion
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B73 " & NroPedix&
                   !MoneEmis = 1
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B74 " & NroPedix&
                   PREUNI# = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B75 " & NroPedix&
                   !TIPOCAM = 1
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B76 " & NroPedix&
                   !PreList = PREUNI#
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B77 " & NroPedix&
                   !DescLista = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B78 " & NroPedix&
                   !DescCombi = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B79 " & NroPedix&
                   !DescRubro = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B80 " & NroPedix&
                   !DescTotal = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B81 " & NroPedix&
                   !CanAsig = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B82 " & NroPedix&
                   !CanPed = CANTI
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B83 " & NroPedix&
                   !PreUnid = PREUNI#
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B84 " & NroPedix&
                   !ImpoTot = PREUNI#
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B85 " & NroPedix&
                   !CANTENT = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B86 " & NroPedix&
                   !CanSaldo = 0
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B87 " & NroPedix&
                   !FeUltEnt = CVDAT(0)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B88 " & NroPedix&
                   !NroRemi = ""
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B89 " & NroPedix&
                   !TPedItem = CAVERI$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B90 " & NroPedix&
                   '
                   TTXYZ$ = Space$(76)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B91 " & NroPedix&
                   Call REPLA(TTXYZ$, "- " + CODEXT$(CI1%), 11, 18)
                   Call REPLA(TTXYZ$, DESCRIT0$(CI1%), 29, 38)
                   Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTI), 3, 10, 1, 2), 57, 10)
                   ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B92 " & NroPedix&
                   '
                 .Update
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B93 " & NroPedix&
           End With
           '
           rs3.Close
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B94 " & NroPedix&
           Set rs3 = Nothing
           .MoveNext
           Loop
        End With
     ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B95 " & NroPedix&
     SQLX$ = " SELECT * FROM PEDENCA "
     SQLX$ = SQLX$ + "WHERE  NroPed = " & NroPedix& & ""
     Dim PEDDDENCA As ADODB.Recordset
     Set PEDDDENCA = New ADODB.Recordset
     PEDDDENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With PEDDDENCA
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B96 " & NroPedix&
        If Not .EOF Then
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B97 " & NroPedix&
        
          If IsNull(!AnotaPed) = True Or SAFETEXT$(!AnotaPed) = "" Then
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B98 " & NroPedix&
            PEDDDENCA!AnotaPed = ANOTA$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B99 " & NroPedix&
          Else
            VarAnotaped$ = SAFETEXT$(!AnotaPed)
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B100 " & NroPedix&
            !AnotaPed = VarAnotaped$ + ANOTA$
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B101 " & NroPedix&
          End If
          .Update
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B102 " & NroPedix&
        End If
      End With
      PEDDDENCA.Close
      Set PEDDDENCA = Nothing
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B103 " & NroPedix&
      '
      'GRABAR EL NUMERO DE PEDIDO Y EL CLIENTE DEL PEDIDO EN LA SOLICITUD DE COMPRA
      SQLX$ = "SELECT * FROM OC_SUCURSAL "
      SQLX$ = SQLX$ + " WHERE SUCURSAL = " & SUCURSAL
      SQLX$ = SQLX$ + " AND O_COMPRA_SUC =" & NUOC&
      Dim RS4 As ADODB.Recordset
      Set RS4 = New ADODB.Recordset
      RS4.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B104 " & NroPedix&
      With RS4
        If Not .EOF Then
           Do While Not .EOF
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B105 " & NroPedix&
           
            !CODIEMPR = CodiEmp%
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B106 " & NroPedix&
            !NPED_GENERADO_EN_AVELLANEDA = NroPedix&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B107 " & NroPedix&
             .Update
             .MoveNext
            Loop
        End If
      End With
      RS4.Close
      Set RS4 = Nothing
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B108 " & NroPedix&
      '
      Call LIBARCH("*.*")
      Call BAJALDISCO
      '
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B109 " & NroPedix&
     OKX% = 1
     '
98   Call LIBARCH("PEDENCA")
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B110 " & NroPedix&
     Call GENERAR_ARCHIVO_GENERO_PEDIDO(NC_Pedido, NroPedix&, NUOC&, "GENEPEDI")
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B111 " & NroPedix&
     Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0ORDER BY NROPED DESC", "NOMESS")
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B112 " & NroPedix&
     Screen.MousePointer = 1
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B113 " & NroPedix&
     
     '
     'ESTO ES PARA QUE RETORNE CON EL VALOR DEL PEDIDO REAL CON EL QUE SE GENERO
     P_Avellaneda& = NroPedix&
If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "B114 " & NroPedix&
     
    
End Sub
Sub GENERAR_ARCHIVO_GENERO_PEDIDO(NCLIE, NROPED&, NUOC&, Optional ACCION$)
   Ruta$ = ""
   
   NC_CONFIG1 = XVALO(Control("RETRASLA", "SUC1CLIE"))
   NC_CONFIG2 = XVALO(Control("RETRASLA", "SUC2CLIE"))
   'VALIDO SI ES CLIENTE DE SUCURSAL
   If NC_CONFIG1 <> NCLIE And NC_CONFIG2 <> NCLIE Then Exit Sub
   If NC_CONFIG1 = NCLIE Then
        Ruta$ = Control("RETRASLA", "SUC1RUTA")
    ElseIf NC_CONFIG2 = NCLIE Then
        Ruta$ = Control("RETRASLA", "SUC2RUTA")
   End If
   '
   Ruta2$ = RutaCarpetaLocalCompraInterna$ 'SI LA RUTA CONFIGURADA EN TABLA AUXILIAR (CTEPROCOMPRA) EXISTE
   If Ruta2$ <> "" Then Ruta$ = Ruta2$
   
   If Ruta$ = "" Then
       Call COMUNI("Atención !!!!! Sucursal Seleccinada No es Válida\Verifique Configuración de Remito a Sucursales\No se Generó Archivo de Envio de Stock Sugeridos")
      Exit Sub
   End If
   If FEMI% = 0 Then FEMI% = HOY(HO$)
   'CARGA LOS TITULOS
   A$ = Date
   A$ = REPLACAR(A$, "/", "")
   B$ = Time
   B$ = REPLACAR(B$, ":", "")
   B$ = REPLACAR(B$, ".", "")
'Call mostrarDatosx("3" & Ruta$ + "\ARCHIVO_GENERO_PEDIDO " + A$ + " " + B$ + ".TXT")

   NARCHI2% = FILEOPEN%(Ruta$ + "\ARCHIVO_GENERO_PEDIDO " + A$ + " " + B$ + ".TXT", 2, 256)    ' ABRE EL ARCHIVO DE ESCRITURA
    TX$ = AJUSTI$("NCLIENTE", 8): TX$ = TX$ + ";" 'NCLIENTE
    TX$ = TX$ + AJUSTI$("N.PEDIDO", 8): TX$ = TX$ + ";" 'NRO DE PEDIDO
    TX$ = TX$ + AJUSTI$("N.COMPRA", 8): TX$ = TX$ + ";"  'NUMERO DE ORDEN DE COMPRA
    TX$ = TX$ + AJUSTI$("EMISION", 8): TX$ = TX$ + ";" 'FECHA DE EMISION
    TX$ = TX$ + AJUSTI$("CI1%", 8): TX$ = TX$ + ";" 'CODIGO ING
    TX$ = TX$ + AJUSTI$("CODIGOEXT", 24): TX$ = TX$ + ";" 'CODIGO EXT
    TX$ = TX$ + AJUSTI$("DESCRIPCION", 32): TX$ = TX$ + ";" 'DESCRIPCION
    TX$ = TX$ + AJUSTI$("CANTIPEDIDA", 10): TX$ = TX$ + ";" 'CANTIDAD PEDIDDA
    TX$ = TX$ + AJUSTI$("CANTIENTRE", 10): TX$ = TX$ + ";" 'CANTIDAD ENTREGADA
    TX$ = TX$ + AJUSTI$("ACCION", 10): TX$ = TX$ + ";" 'ACCION
    TX$ = TX$ + AJUSTI$("USUARIO", 8): TX$ = TX$ 'USUARIO
   Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
   FE% = HOY(HO$)
   NUSER1% = USNBR% Mod 100
   
   SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NROPED = " & NROPED&
   SQLX$ = SQLX$ + " AND  NroClie = " & NCLIE
   Set rs = READSET(SQLX$)
   With rs
        Do While Not .EOF
            TX$ = AJUSTI$(SAFETEXT$(NCLIE), 8): TX$ = TX$ + ";" 'NCLIENTE
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NROPED&), 8): TX$ = TX$ + ";" 'NRO DE PEDIDO
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUOC&), 8): TX$ = TX$ + ";"  'NUMERO DE ORDEN DE COMPRA
            TX$ = TX$ + AJUSTI$(FECHATEX$(FEMI%), 8): TX$ = TX$ + ";" 'FECHA DE EMISION
            TX$ = TX$ + AJUSTI$(XVALO(!CODIINT), 8): TX$ = TX$ + ";"  'CODIGO ING
            TX$ = TX$ + AJUSTI$(SAFETEXT$(!CODIEXT), 24): TX$ = TX$ + ";" 'CODIGO EXT
            TX$ = TX$ + AJUSTI$(SAFETEXT$(!Descrip), 32): TX$ = TX$ + ";" 'DESCRIPCION
            TX$ = TX$ + AJUSTI$(FORMATNUM$(XVALO(!CanPed), "F10.1"), 10): TX$ = TX$ + ";" 'CANTIDAD PEDIDDA
            TX$ = TX$ + AJUSTI$(FORMATNUM$(XVALO(!CANTENT), "F10.1"), 10): TX$ = TX$ + ";" 'CANTIDAD ENTREGADA
            TX$ = TX$ + AJUSTI$(SAFETEXT$(ACCION$), 10): TX$ = TX$ + ";" 'ACCION
            TX$ = TX$ + AJUSTI$(SAFETEXT$(NUSER1%), 8): TX$ = TX$ 'USUARIO
            If OCOMSUCU.Check1.Value = 1 Then OCOMSUCU.Check1.Caption = "C23 "
            Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
            
            .MoveNext
         Loop
             
   End With
   Close #NARCHI2%
   rs.Close
   
   Set rs = Nothing

End Sub


Function NUPEDIDO&()
   '
   NPP& = 1
   Call ABREPEDCLI
   NPXX& = 1 + XVALO(VALOBADA("PEDENCA", "MAX(NROPED)", "NROPED > 0", "NOMESS"))
   If NPXX& > NPP& Then NPP& = NPXX&
   NUPEDIDO& = NPP&
   '
End Function

Private Sub Command4_Click()
   APROBO% = 0
   Unload Me
End Sub

Private Sub Form_Load()
  Call CREAR_TABLA_DetaOVenta
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  APROBO% = 0
End Sub

Private Sub GRID_Click()

    Static MODO  As Boolean
    Reg& = GRID.MouseRow
    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    Else
'        Call COLOREAR_GRILLASQL_TODAS_MISMO_COLOR(GRID, vbWhite) 'PARA QUE VUELVA TODO A BLANCO Y LUEGO PINTE LA ULTIMA SELECCIONADA
        For i& = 1 To GRID.Rows - 1
           If TRIM$(GRID.TextMatrix(i&, 0)) = "*" Then
             
             Call COLOREAR_GRILLA_SOLAFILA(GRID, &HFFFFFF, i&, 0, 0, 0)
             GRID.TextMatrix(i&, 0) = ""
           End If
        Next i&
        '
        'SI YA TIENE PEDIDO GENERADO DESHABILITA EL BOTON DE GRABAR PARA QUE NO GRABE NUEVAMENTE
        If XVALO(GRID.TextMatrix(Reg&, 7)) > 0 Then
          Command3.Enabled = False
          Command3.Caption = "Pedido: " & XVALO(GRID.TextMatrix(Reg&, 7)) & " Yá Generado"
        Else
          Command3.Enabled = True
          Command3.Caption = "Grabar Pedido Seleccionado"
        End If
        Call COLOREAR_GRILLA_SOLAFILA(GRID, &HE0E0E0, Reg&, 0, 0, 0)
        Call MOSTRARDETALLE(GRID.MouseRow)
        Call MOSTRAR_MENSAJE(Reg&, XVALO(GRID.TextMatrix(Reg&, 7)))
    End If
End Sub


Private Sub GRID_DblClick()
   Reg& = GRID.MouseRow

   NPED& = XVALO(GRID.TextMatrix(Reg&, 7))
   If NPED& > 0 Then
      Call OPPED07.CAMBIAPED(NPED&)
   End If
End Sub


Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     Reg& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, Reg&, COL%)
  End If

End Sub

Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then

        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    Else
       Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.Row, 0, 0, 0)
    End If
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  '
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub MSFDET_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
     Reg& = MSFDET.MouseRow
     COL% = MSFDET.MouseCol
     Call MENU_CELDA(MSFDET, Reg&, COL%)
  End If
End Sub

Private Sub MSFMEN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then

     Reg& = MSFMEN.MouseRow
     COL% = MSFMEN.MouseCol
     Call MENU_CELDA(MSFMEN, Reg&, COL%)
    End If
   
End Sub
Sub MENU_CELDA(MSF As msFlexGrid, Reg&, COL%)
        If MSF.Rows <= 1 Then Exit Sub

        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, Reg&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(Reg&, COL%))  'CELDA SELECCIONADA
        PopupMenu menu
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, Reg&, COL%)  ' LA PINTA DE AMARILLA
End Sub

'Private Sub MSF2_Click()
'    Static MODO  As Boolean
'    If (MSF2.MouseRow = 0) Then
'
'        ' Ordena en forma ascendente
'        If MODO Then
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 2
'            MODO = False
'        ' Ordena en forma descendente
'        Else
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 1
'            MODO = True
'        End If
'    End If
'
'End Sub



Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    Reg& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            Reg& = Reg& + 1
            MSF_2.Rows = Reg& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(Reg&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
Sub MOSTRARDETALLE(Reg&)
  If Reg& = 0 Then Exit Sub
  '
  MSFDET.Rows = 1
  OCOMPRA& = XVALO(GRID.TextMatrix(Reg&, 3))
  SUCUR = XVALO(GRID.TextMatrix(Reg&, 1))
  CLIENTEPEDIDO = XVALO(GRID.TextMatrix(Reg&, 5))
  
  SQLX$ = "SELECT * FROM OC_SUCURSAL WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE O_COMPRA_SUC = " & OCOMPRA&
  SQLX$ = SQLX$ + " AND SUCURSAL = " & SUCUR
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
       CI1% = XVALO(!CODINT)
       DESC$ = DESCRIT0$(CI1%)
       CANTICOMPRA = XVALO(!Cantidad)
       FESOL% = CVINT(!FECHA_SOLICI_COMPRA)
       NUMPEDAVELLANEDA& = XVALO(!NPED_GENERADO_EN_AVELLANEDA)
       If NUMPEDAVELLANEDA& > 0 Then
          CONDI$ = " Nroped = " & NUMPEDAVELLANEDA&
          CONDI$ = CONDI$ + " and CodiInt = " & CI1%
          'CONDI$ = CONDI$ + " and FeSolEnt = '" & FETEXCOR1$(FESOL%) & "'" 'comentado p q' si no coincide la fecha no me trae los datos del pedido
          CANPEDIDA = XVALO(VALOBADA("PEDDETA", "CanPed", CONDI$, "NOMESS"))
          CANENTRE = XVALO(VALOBADA("PEDDETA", "CANTENT", CONDI$, "NOMESS"))
          SALDO = CANPEDIDA - CANENTRE: If SALDO < 0.05 Then SALDO = 0
       End If
       J& = J& + 1
       MSFDET.Rows = J& + 1
       MSFDET.TextMatrix(J&, 1) = CODEXT$(CI1%)
       MSFDET.TextMatrix(J&, 2) = DESC$
       MSFDET.TextMatrix(J&, 3) = TRIM$(FORMATNUM$(CANTICOMPRA, "F9.1"))
       MSFDET.TextMatrix(J&, 4) = FECHATEX$(FESOL%)
       MSFDET.TextMatrix(J&, 5) = NUMPEDAVELLANEDA&
       MSFDET.TextMatrix(J&, 6) = TRIM$(FORMATNUM$(CANPEDIDA, "F9.1"))
       MSFDET.TextMatrix(J&, 7) = TRIM$(FORMATNUM$(CANENTRE, "F9.1"))
       MSFDET.TextMatrix(J&, 8) = TRIM$(FORMATNUM$(SALDO, "F9.1"))
      .MoveNext
    Loop
  End With
  rs.Close
  Set rs = Nothing
  
  
End Sub
Sub MOSTRAR_MENSAJE(Reg&, NROPED&)
  If Reg& = 0 Then Exit Sub
  '
  Call CREAR_TABLA_MENSAJES_DE_SUCURSAL
  MSFMEN.Rows = 1
  OCOMPRA& = XVALO(GRID.TextMatrix(Reg&, 3))
  SUCUR = XVALO(GRID.TextMatrix(Reg&, 1))
  CLIENTEPEDIDO = XVALO(GRID.TextMatrix(Reg&, 5))
  
  SQLX$ = "SELECT * FROM MENSA_SUCURSAL WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE O_COMPRA_SUC = " & OCOMPRA&
  SQLX$ = SQLX$ + " AND SUCURSAL = " & SUCUR
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
       J& = J& + 1
       MSFMEN.Rows = J& + 1
       MENSAX$ = SAFETEXT$(!Mensaje)
       Label75 = MENSAX$
       MSFMEN.RowHeight(J&) = Label75.Height + 100

       MSFMEN.TextMatrix(J&, 1) = XVALO(!O_COMPRA_SUC)
       MSFMEN.TextMatrix(J&, 2) = MENSAX$
       MSFMEN.TextMatrix(J&, 3) = NROPED&
       
      .MoveNext
    Loop
  End With
  rs.Close
  Set rs = Nothing
  
  
End Sub
Sub mostrar_pedido(NPX&)
    '
    If NPX& < 1 Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    NUDOCU$ = TRIM$(str$(NPX&))
    TIDOCUM$ = "Pedido de Cliente"
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
    DOCUNU& = DOCID&(TIDOCUM$, str$(NPX&), NUDOCU$)
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, str$(NPX&), NUDOCU$)
    Screen.MousePointer = 1

End Sub
