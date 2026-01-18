VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IMPOLOGA 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16095
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   11
      Top             =   6360
      Width           =   12375
      Begin VB.CommandButton Command1 
         Caption         =   "Migrar Datos de Pantalla "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   10800
         TabIndex        =   20
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton Command22 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   16
         Top             =   480
         Width           =   220
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   7035
         TabIndex        =   14
         Top             =   480
         Width           =   7095
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   15
            TabIndex        =   15
            Top             =   0
            Width           =   15000
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   615
         Left            =   10080
         Picture         =   "IMPOLOGA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Establece Conexion con Origen de Datos"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   7680
         TabIndex        =   12
         Text            =   " "
         Top             =   460
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección de Tabla a Importar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   3375
      Begin VB.OptionButton Option5 
         Caption         =   "Sectores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tabla UM2PTH Costo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3015
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Tabla UM2PTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tabla UM2COSTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tabla UM2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "IMPOLOGA.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "IMPOLOGA.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid DbGrid1 
      Height          =   6165
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Click en el Detalle Para Seleccionar"
      Top             =   120
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   10874
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "IMPOLOGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public APROBO%
Dim LecturaDatosOriginales As Byte
Public HaySeleccion% '0 no hay ningun items seleccionado, 1 hay un item o mas seleccionados
Public MSFACTIVO As MSFlexGrid
Private objexcel As Excel.Application
Private Libro As Excel.Workbook
Private HOJA As Excel.Worksheet

Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub Combo1_Change()
   'Set objexcel = New Excel.Application
   'Set Libro = objexcel.Workbooks.Open(Label1.Caption)
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.List(KKI% - 1) = Libro.Sheets(KKI%).Name Then
        Set HOJA = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set HOJA = Libro.Sheets(1)
   DbGrid1.Rows = 1

End Sub

Private Sub Combo1_Click()
   'Set objexcel = New Excel.Application
   'Set Libro = objexcel.Workbooks.Open(Label1.Caption)
   CANOJ% = Libro.Sheets.Count
   For KKI% = 1 To CANOJ%
      If Combo1.TEXT = Libro.Sheets(KKI%).Name Then
        Set HOJA = Libro.Sheets(KKI%)
        Exit Sub
      End If
   Next KKI%
   Combo1.TEXT = Libro.Sheets(1).Name
   Set HOJA = Libro.Sheets(1)
   DbGrid1.Clear
   DbGrid1.Rows = 1
   '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   
   If Option1.Value = True Then
      TABLA$ = "TABLAUM2"
   ElseIf Option2.Value = True Then
      TABLA$ = "TABLAUM2COS"
   ElseIf Option3.Value = True Then
      TABLA$ = "UM2PTH"
   ElseIf Option4.Value = True Then
      TABLA$ = "UM2PTHCOS"
   ElseIf Option5.Value = True Then
      TABLA$ = "SECTORES"
   Else
      Call COMUNI("Falta Seleccionar Opción de Tabla a Importar")
      GoTo 99
   End If
   
   OPX% = COMALTER%("Se va a Importar La Información de Pantalla a la Tabla: " & TABLA$ & "\\Cancelar\Importar Sobre-Escribiendo la Información Existente")
   If OPX% < 2 Then GoTo 99
   If TABLA$ = "SECTORES" Then
      Call ACTUREGISTRO(TABLA$, "MARBORRA", "ID > 0 ", 1, REGAF&, "NOMESS")
   Else
      Call ACTUREGISTRO(TABLA$, "MARBORRA", "Lote_M2 > 0 ", 1, REGAF&, "NOMESS")
   End If
   
   FIN& = DbGrid1.Rows - 1
   JJ& = 0
   'REBLA$ = REGBLAN$(ARCHIPRE$)
   
   SQLX$ = "SELECT * FROM " & TABLA$
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   '
   With RS
   Dim valor$(24)
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
       For i& = 0 To DbGrid1.Cols - 1
          valor$(i&) = DbGrid1.TextMatrix(U&, i&)
       Next i&
       Select Case TABLA$
       Case "TABLAUM2"
           .AddNew
           !LOTE_M2 = valor$(0)
           !FR4D16SEMI35 = valor$(1)
           !FR2D16SEMI35 = valor$(2)
           !FR4D16Doble = valor$(3)
           !PoliesterSimple = valor$(4)
           !Polymida = valor$(5)
           !FR4S24SEMI35 = valor$(6)
           !FR4S32SEMI35 = valor$(7)
           !FR4S08SEMI35 = valor$(8)
           !FR4S10SEMI35 = valor$(9)
           !FR4S12SEMI35 = valor$(10)
           !FactorCosto = XVALO(valor$(11))
           !INTERPOLAR = "L"
           !MARBORRA = 0
           .Update

       
       Case "TABLAUM2COS"
           .AddNew

           !LOTE_M2 = valor$(0)
           !FR4SimpleCu = valor$(1)
           !FR4SimpleCel = valor$(2)
           !FR2SimpleCu = valor$(3)
           !FR4DobleCu = valor$(4)
           !FR4DobleCel = valor$(5)
           !PoliesterSimple = valor$(6)
           !FR4S24SEMI35 = valor$(7)
           !FR4S32SEMI35 = valor$(8)
           !Polymida = valor$(9)
           !FR4S08SEMI35 = valor$(10)
           !FR4S10SEMI35 = valor$(11)
           !FR4S12SEMI35 = valor$(12)
           !FactorCosto = valor$(13)
           !INTERPOLAR = "L"
           !MARBORRA = 0
           .Update

       
       Case "UM2PTH", "UM2PTHCOS"
           .AddNew
          !LOTE_M2 = valor$(0)
          !FR4D16PTH18 = valor$(1)
          !FR4D24PTH35 = valor$(2)
          !FR4D32PTH35 = valor$(3)
          !FR4D16PTH70 = valor$(4)
          !FR4D08PTH35 = valor$(5)
          !FR4D10PTH35 = valor$(6)
          !FR4D12PTH35 = valor$(7)
          !FR4D04PTH35 = valor$(8)
          !FactorCosto = valor$(9)
          !INTERPOLAR = "L"
          !MARBORRA = 0
          .Update
       Case "SECTORES"
          .AddNew
          !ID = XVALO(valor$(0))
          !SECTOR = Left$(valor$(1), 256)
          !SECUENPROD = XVALO(valor$(2))
          !MARBORRA = 0
          .Update
        End Select
         '
   Next U&
   End With

   '
   RS.Close
   Set RS = Nothing
   '
   CONDIBORRA$ = "MARBORRA = 1"
   Call BORRAREGISTROS(TABLA$, CONDIBORRA$, REGAF&, "NOMESS")
   Call COMUNI("Migración Finalizada")
99 Command1.Enabled = True

End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|(*.XLSX)|*.XLSX|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   Openforms = DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.FileName)
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command7_Click()
  Screen.MousePointer = 11
   DbGrid1.Rows = 1
   CANFIL& = 100   ' Hoja.Rows
   CANCOL& = DbGrid1.Cols
   JK& = 0
   CONTAERR& = 0
   ultimafila& = 0
   
   
   For i& = 1 To CANFIL&
     YAGRELI% = 0
     For J& = 1 To CANCOL&
       valor = HOJA.Cells(i&, J&)
       If TRIM$(CStr(valor)) <> "" Then
         If YAGRELI% < 1 Then
           YAGRELI% = 1
           JK& = JK& + 1
         End If
         If JK& > DbGrid1.Rows Then DbGrid1.Rows = DbGrid1.Rows + 1
         If JK& < 1 Then Call COMUNI("Imposible Continuar.\Falta Encabezado en Alguna de las Columnas"): Screen.MousePointer = 1: Exit Sub
         
         DbGrid1.Row = JK& - 1
       Else
         If YAGRELI% = 0 Then
            If ultimafila& < 1 Then
                ultimafila& = DbGrid1.Row + 1
                If ultimafila& >= DbGrid1.Rows Then DbGrid1.Rows = DbGrid1.Rows + 1
            End If
            DbGrid1.Row = ultimafila&
            YAGRELI% = 1
         End If
       End If
         If J& > DbGrid1.Cols Then DbGrid1.Cols = DbGrid1.Cols + 1
       DbGrid1.COL = J& - 1
       '
       If DbGrid1.Row = 4 Then
       A = B
       End If
       On Error Resume Next
       DbGrid1.TEXT = valor
       'CAPTURA DE ERROR DE DATOS NO VALIDOS DE PLANILLA EXCEL
       If Err > 0 Then
           CONTAERR& = CONTAERR& + 1
          'GUARDO EL DATO DE FILA Y COLUMNA DE LA PRIMER OCURRENCIA DE ERROR
          If CONTAERR& = 1 Then Fil& = i&: COL& = J&
          DbGrid1.TEXT = ""
       End If
       On Error GoTo 0
       
       '
       Label9.Font = DbGrid1.Font
       Label9.FontSize = DbGrid1.CellFontSize
       Label9.FontBold = DbGrid1.CellFontBold
       Label9.Caption = CStr(valor)
       If Label9.Width > DbGrid1.ColWidth(J& - 1) Then
         DbGrid1.ColWidth(J& - 1) = Label9.Width
       End If
     Next J&
     If i& > JK& + 256 Then
        GoTo 99
     End If
   Next i&
   '
99 Screen.MousePointer = 1
   '
   If ultimafila& > 0 Then DbGrid1.RemoveItem (ultimafila&)
   '
   'MENSAJE SI SE PRODUJO ERROR CAPTURADO DE DATO NO VALIDO
   If CONTAERR& > 0 Then
     Call MENSERR(24, "Se Han Detectado Uno o Más Datos No Válidos\ Primer Ocurrencia Fila: " & Fil& & " Columna: " & COL& & "\El o Los Mismos Fueron Cambiados por Texto Vacío \Puede Continuar Normalmente o Modificar La Planilla de Importación\ Y Realizar Nuevamente la Importación")
   End If
   '
   Call COMUNI("Origen de Datos Establecido.")
   '
End Sub

Private Sub Form_Load()
   Call Option1_Click
End Sub

Private Sub Label1_Change()
  DbGrid1.Rows = 1
  Combo1.Clear
  Combo1.TEXT = ""
  '
  If Label1.Caption <> "" Then
     'Dim objexcel As Excel.Application
     Set objexcel = New Excel.Application
     Set Libro = objexcel.Workbooks.Open(Label1.Caption)
     '
     CANOJ% = Libro.Sheets.Count
     For KKI% = 1 To CANOJ%
        Combo1.AddItem Libro.Sheets(KKI%).Name
        If KKI% = 1 Then
           Combo1.TEXT = Libro.Sheets(KKI%).Name
        End If
     Next KKI%
     Call Combo1_Click
     On Error Resume Next
     'Libro.Close
     'Set Libro = Nothing
     On Error GoTo 0
  End If
  
End Sub

Private Sub Option1_Click()
    Call CARGA_ENCABEZADO(Me.DbGrid1, "Lote(M2)/F12.2;FR4D16SEMI35/F12.2;FR2D16SEMI35/F12.2;FR4D16Doble/F12.2;PoliesterSimple/F12.2;Polymida/F12.2;FR4S24SEMI35/F12.2;FR4S32SEMI35/F12.2;FR4S08SEMI35/F12.2;FR4S10SEMI35/F12.2;FR4S12SEMI35/F12.2;FactorCosto/F12.2", FRMZELECHO)
End Sub

Private Sub Option2_Click()
    Call CARGA_ENCABEZADO(Me.DbGrid1, "Lote(M2)/F12.2;FR4SimpleCu/F12.2;FR4SimpleCel/F12.2;FR2SimpleCu/F12.2;FR4DobleCu/F12.2;FR4DobleCel/F12.2;PoliesterSimple/F12.2;FR4S24SEMI35/F12.2;FR4S32SEMI35/F12.2;Polymida/F12.2;FR4S08SEMI35/F12.2;FR4S10SEMI35/F12.2;FR4S12SEMI35/F12.2;FactorCosto/F12.2", FRMZELECHO)

End Sub

Private Sub Option3_Click()
    Call CARGA_ENCABEZADO(Me.DbGrid1, "Lote(M2)/F12.2;FR4D16PTH18/F12.2;FR4D24PTH35/F12.2;FR4D32PTH35/F12.2;FR4D16PTH70/F12.2;FR4D08PTH35/F12.2;FR4D10PTH35/F12.2;FR4D12PTH35/F12.2;FR4D04PTH35/F12.2;FactorCosto/F12.2", FRMZELECHO)

End Sub

Private Sub Option4_Click()
    Call CARGA_ENCABEZADO(Me.DbGrid1, "Lote(M2)/F12.2;FR4D16PTH18/F12.2;FR4D24PTH35/F12.2;FR4D32PTH35/F12.2;FR4D16PTH70/F12.2;FR4D08PTH35/F12.2;FR4D10PTH35/F12.2;FR4D12PTH35/F12.2;FR4D04PTH35/F12.2;FactorCosto/F12.2", FRMZELECHO)

End Sub
