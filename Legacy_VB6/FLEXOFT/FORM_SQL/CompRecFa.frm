VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CompRecFa 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12480
      TabIndex        =   15
      ToolTipText     =   "Doble Click Para Cambiar Configuración de Cantidad de Dias."
      Top             =   3435
      Width           =   1215
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
   Begin VB.Frame Frame2 
      Caption         =   "Detalle de Facturas Sistema Externo                                                        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14430
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13725
         TabIndex        =   16
         Top             =   3435
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Frame Frame1 
         Height          =   600
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Top             =   3240
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
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   9
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
            Left            =   80
            TabIndex        =   8
            Top             =   160
            Width           =   585
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRID 
         Height          =   3045
         Left            =   15
         TabIndex        =   10
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Click Para Ver Situación MRP al Momento de Generación de la Orden de Compra"
         Top             =   240
         Width           =   14355
         _ExtentX        =   25321
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
      Begin MSFlexGridLib.MSFlexGrid MSF_2 
         Height          =   3045
         Left            =   20
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   14350
         _ExtentX        =   25321
         _ExtentY        =   5371
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
      Begin VB.Label LHAS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   11760
         TabIndex        =   18
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LDEs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   10680
         TabIndex        =   17
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Configurar Qt   Dias Recepción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   9735
         TabIndex        =   14
         Top             =   3550
         Width           =   2625
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Items Facturados y Componentes Agrupados  segun Rango Seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   14415
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
         Left            =   10680
         TabIndex        =   28
         Top             =   3830
         Width           =   1485
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Componentes"
         Top             =   3830
         Width           =   1500
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Facturados"
         Top             =   3830
         Width           =   1460
      End
      Begin VB.Frame Frame1 
         Height          =   600
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   3720
         Width           =   7455
         Begin VB.TextBox TXTBUSCAR1 
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
            TabIndex        =   22
            Top             =   160
            Width           =   2535
         End
         Begin VB.ComboBox CMBFILTRO1 
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
            TabIndex        =   21
            Top             =   160
            Width           =   3135
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
            Index           =   2
            Left            =   80
            TabIndex        =   24
            Top             =   160
            Width           =   585
         End
         Begin VB.Label Label1 
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
            TabIndex        =   23
            Top             =   160
            Width           =   705
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ver Stocks Previo y posterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   12400
         TabIndex        =   19
         ToolTipText     =   "Recarga La Griila Inferior Agregando el Stock Previo al Rango Seleccionado y el Stocks Posterior al Mismo."
         Top             =   3850
         Width           =   1850
      End
      Begin MSFlexGridLib.MSFlexGrid MSFDET 
         Height          =   3525
         Left            =   15
         TabIndex        =   13
         ToolTipText     =   "Doble Click Para Ver Detalle de Consumo de Componentes"
         Top             =   240
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   6218
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   3525
         Left            =   15
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   14355
         _ExtentX        =   25321
         _ExtentY        =   6218
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
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   1000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
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
Attribute VB_Name = "CompRecFa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APROBO%
Dim GLOBALFRM_strContenidoCelda$
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private Hoja As Excel.Worksheet
Private Sub archivo_Click()

End Sub

Private Sub Check1_Click()
    '
    If Check1.Value = 1 Then
      campos$ = "Código/A20;Descripción/A32;Stock Prev./F10.1;Cant.Fact/F10.1;Cant.Recibida/F10.1;Diferencia/F10.1;Sock Fin/F10.1"
    Else
      campos$ = "Código/A20;Descripción/A32;Stock Prev./F0;Cant.Fact/F10.1;Cant.Recibida/F10.1;Diferencia/F10.1;Sock Fin/F0"
    End If
    '
    Call CARGA_ENCABEZADO(CompRecFa.MSFDET, campos$, CompRecFa)
    Call CARGA_ENCABEZADO(CompRecFa.MSF3, campos$, CompRecFa)
    Call CargarComboSegunGrid(MSF3, CMBFILTRO1, 1, campos$)
    Call SITSTOK07.CargarDatosSysflom(XVALO(LDEs), XVALO(LHAS))
    '
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
     TITULOX$ = "Productos Facturados y Consumos desde el " & FECHATEX$(XVALO(LDEs)) & " Hasta el " & FECHATEX$(XVALO(LHAS))
     If TRIM$(TXTBUSCAR1) <> "" Then
        TITULOX$ = TITULOX$ & "(Filtro: " & TXTBUSCAR1 & ")"
     End If
     If MSFDET.Visible = True Then
      Call GeneraExcelDeGrilla(AREXPO$, CompRecFa.MSFDET, TITULOX$)
'      Call GENEXCEL(AREXPO$, CompRecFa.MSFDET)
     Else
      'Call GENEXCEL(AREXPO$, CompRecFa.MSFDET, TITULOX$)
      Call GeneraExcelDeGrilla(AREXPO$, CompRecFa.MSFDET, TITULOX$)
     End If
   End If
   '

99 Command1.Enabled = True

End Sub

Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
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
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
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
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
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
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Command2_Click()
    Command2.Enabled = False
    If XVALO(MONUME$) >= 0 Then
      MONUME$ = Text1.TEXT
    Else
      Call COMUNI("Valor Ingresado No Válido - Debe Ser Numérico")
      GoTo 99
    End If
    Call GRACONTROL("COMPARAT", "RECEFACT", MONUME$)
    Call Text1_DblClick
99  Command2.Enabled = True
End Sub

Private Sub Form_Load()
 Text1 = CONTROL("COMPARAT", "RECEFACT")
 Call CargarComboSegunGrid(MSF3, CMBFILTRO1, 1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  APROBO% = 0
End Sub

Private Sub Frame4_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub GRID_Click()
   Call EVENTOCLICK(GRID)
End Sub
Sub EVENTOCLICK(GRID As MSFlexGrid)
     Static MODO  As Boolean
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
     End If
     
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     If REG& < 1 Then Exit Sub
     CI1 = XVALO(GRID.TextMatrix(REG&, 1))
End Sub


Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = GRID.MouseRow
     COL% = GRID.MouseCol
     Call MENU_CELDA(GRID, REG&, COL%)
  End If

End Sub

Private Sub Label2_Click()

End Sub

Private Sub MSF_2_Click()
    Call EVENTOCLICK(MSF_2)
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

Private Sub MSF3_Click()
   Call EVENTOCLICK(MSF3)
End Sub

Private Sub MSFDET_Click()
   Call EVENTOCLICK(MSFDET)
End Sub

Private Sub MSFDET_DblClick()
    REG& = MSFDET.MouseRow
    If REG& <= 1 Then Exit Sub
    If MSFDET.TextMatrix(REG&, 0) = "*" Then
       FRMZELECHO.TXTBUSCAR.TEXT = Me.MSFDET.TextMatrix(MSFDET.Row, 1)
       FRMZELECHO.TXTBUSCAR.Locked = True
       FRMZELECHO.NOCERRAR% = 1
       FRMZELECHO.Width = 14200
       FRMZELECHO.Caption = "Detalle de Composición de Consumos Según Productos Facturados"
       
       FRMZELECHO.Show 1
    End If
End Sub
Private Sub MSF3_DblClick()
    REG& = MSF3.MouseRow
    If REG& < 1 Then Exit Sub
    If MSF3.TextMatrix(REG&, 0) = "*" Then
       FRMZELECHO.TXTBUSCAR.TEXT = Me.MSF3.TextMatrix(MSF3.Row, 1)
       FRMZELECHO.TXTBUSCAR.Locked = True
       FRMZELECHO.NOCERRAR% = 1
       FRMZELECHO.Width = 14200
       FRMZELECHO.Caption = "Detalle de Composición de Consumos Según Productos Facturados"
       FRMZELECHO.Show 1
    End If
End Sub

Private Sub MSFDET_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSFDET.MouseRow
     COL% = MSFDET.MouseCol
     Call MENU_CELDA(MSFDET, REG&, COL%)
  End If
End Sub

Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        PopupMenu menu
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Private Sub Option1_Click()
End Sub

Private Sub Text1_DblClick()
   If Text1.BackColor = &HE0E0E0 Then
       Text1.BackColor = &HFFFFFF
       Command2.Visible = True
       Text1.Locked = False
   Else
       Text1.BackColor = &HE0E0E0
       Command2.Visible = False
       Text1.Locked = True
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Command2_Click
    End If
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
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For j& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = GRID.TextMatrix(i&, h&)
            Next h&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
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
Private Sub TXTBUSCAR1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR1.TEXT = ""
   End If
End Sub
 

Private Sub TXTBUSCAR1_Change()
    Static opt%
    opt% = Check1.Value = 1
   If opt% <> Check1.Value Then
        If Check1.Value = 1 Then
          campos$ = "Código/A20;Descripción/A32;Stock Prev./F10.1;Cant.Fact/F10.1;Cant.Recibida/F10.1;Diferencia/F10.1;Sock Fin/F10.1;FechaNum/D0"
        Else
          campos$ = "Código/A20;Descripción/A32;Stock Prev./F0;Cant.Fact/F10.1;Cant.Recibida/F10.1;Diferencia/F10.1;Sock Fin/F0;FechaNum/D0"
        End If
    End If
    
    If TRIM$(TXTBUSCAR1) = "" Then
      MSF3.Visible = False
      MSFDET.Visible = True
      Exit Sub
    End If
    '
    MSF3.Rows = 2
    For i& = 1 To MSF3.Cols - 1
       MSF3.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF3.Visible = True
    MSF3.ZOrder 0
    MSF3.Cols = MSFDET.Cols
    MSF3.Redraw = False
    MSFDET.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To MSFDET.Rows - 1
      For j& = 1 To MSFDET.Cols - 1
        If (InStr(UCase$(MSFDET.TextMatrix(i&, j&)), UCase$(TXTBUSCAR1))) Then
          If (TRIM$(CMBFILTRO1.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO1)) = TRIM$(UCase$(MSFDET.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF3.Rows = REG& + 1
            For h& = 0 To MSF3.Cols - 1
              MSF3.TextMatrix(REG&, h&) = MSFDET.TextMatrix(i&, h&)
            Next h&
          End If
          If (TRIM$(CMBFILTRO1.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next i&
    '
    MSF3.Redraw = True

End Sub
