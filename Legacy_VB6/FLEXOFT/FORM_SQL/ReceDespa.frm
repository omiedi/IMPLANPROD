VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ReceDespa 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5985
   ClientLeft      =   2.45790e5
   ClientTop       =   600
   ClientWidth     =   14160
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Height          =   700
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   13935
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
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
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Asignación/Modificación Datos Importación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10080
         Picture         =   "ReceDespa.frx":0000
         TabIndex        =   14
         ToolTipText     =   "Asignación/Modificación Datos de Importación"
         Top             =   280
         Width           =   3795
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   1560
         TabIndex        =   26
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "B. Recepción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   25
         Top             =   120
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cta."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   2880
         TabIndex        =   24
         Top             =   120
         Width           =   285
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   4200
         TabIndex        =   23
         Top             =   120
         Width           =   1740
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1320
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Skin1 
      Height          =   480
      Left            =   3960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   15
      Top             =   6000
      Width           =   1200
   End
   Begin VB.PictureBox SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      ScaleHeight     =   165
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   5115
      Left            =   120
      TabIndex        =   16
      Top             =   800
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   9022
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   4035
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   8205
      Begin VB.CommandButton Command3 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7755
         Picture         =   "ReceDespa.frx":030A
         TabIndex        =   11
         ToolTipText     =   "Enumerar Grilla"
         Top             =   180
         Width           =   345
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
         TabIndex        =   0
         Top             =   160
         Width           =   2535
      End
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
         TabIndex        =   1
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
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   160
         Width           =   585
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
         TabIndex        =   8
         Top             =   160
         Width           =   705
      End
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   1605
      TabIndex        =   12
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   9480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuMenu1 
         Caption         =   "MENU1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMenu2 
         Caption         =   "MENU2"
         Visible         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Begin VB.Menu mnuMenu3 
         Caption         =   "MENU3"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "ReceDespa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long





Private Sub BOTEXIT_Click()
  Unload Me
End Sub


Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub CMBFILTRO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Combo2_Change()

End Sub

Private Sub Command1_Click()
   Dim obj As New RECORXET
        CAMPOS$ = "N°.Recep./f9;N° Prov./F7;Proveedor/A24;Fecha/D8"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
        FRMZELECHO.Caption = ""
        
        Set RS1 = obj.RS_TODARECE()
        Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        
        FRMZELECHO.Caption = "LISTA DE RECEPCIONES"
        FRMZELECHO.Width = 12000
        FRMZELECHO.Show 1
        If XVALO(RESP_ZELE_VECT(1)) > 0 Then
          Text5 = XVALO(RESP_ZELE_VECT(1))
          Text1 = TRIM$(RESP_ZELE_VECT(4))
          Text3 = TRIM$(RESP_ZELE_VECT(2))
          Text4 = TRIM$(RESP_ZELE_VECT(3))
        End If
        



End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
  Dim MSF As MSFlexGrid
  Set MSF = GRID
  If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
  Call numerargrilla(MSF, 0)

End Sub


Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
    Unload VENTABNEW
    VENTANA.Inicializar = 0
    VENTANA.CampEditables = "1;2;3"
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.ColEcoCampo(3) = "CAMPO1;CAMPO2;DATASQL;LISTA = 'ADUANA'"
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = ""
    VENTANA.Inicializar = 1
    VENTANA.Campox = "Fecha/D8;Despacho/A48;Cód/f8;Aduana/A48"
    
    '****************************************************************
    'ASIGNO LOS DATOS SEGUN LA BRNUM AL OBJETO, Y LOS LEO APLICANDOLOS A LA GRILLA
    Dim obj As New Art_Importados
    Call obj.READ_DATOS_IMPO(XVALO(Text5))
    VENTABNEW.MSF.TextMatrix(1, 1) = obj.FECHA
    VENTABNEW.MSF.TextMatrix(1, 2) = obj.Despacho
    VENTABNEW.MSF.TextMatrix(1, 4) = obj.Aduana
    Set obj = Nothing
10  VENTABNEW.Show 1
    
    If VENTABNEW.APROBADO% > 0 Then
        'LEO LOS DATOS INGRESADOS
        FECHA = VENTABNEW.MSF.TextMatrix(1, 1)
        Despacho$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, 2))
        Aduana$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, 4))
        '
        'VALIDACIONES
        If IsDate(FECHA) <> True Then
           Call COMUNI("Fecha Incorrecta")
           GoTo 10
        End If
        If Despacho$ = "" Then
           Call COMUNI("Falta Ingresar Despacho")
           GoTo 10
        End If
        If Aduana$ = "" Then
           Call COMUNI("Falta Ingresar Aduana")
           GoTo 10
        End If
        'FIN VALIDACIONES
        '
        Dim NUBORE As Long ' lo declaro por que si no daba error de tipo en la llamada
        NUBORE = XVALO(Text5)
        'ASIGNO VALORES A LOS ATRIBUTOS DE LA CLASE
        Dim art_impo As New Art_Importados
    
        art_impo.Aduana = Aduana$
        art_impo.Despacho = Despacho$
        art_impo.FECHA = FECHA
        
        'FIN ASIGNACION
        
        'CARGO LA GRILLA Y GRABO
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

        For i& = 1 To GRID.Rows - 1
           LOTE$ = GRID.TextMatrix(i&, 1)
           GRID.TextMatrix(i&, 6) = FECHA
           GRID.TextMatrix(i&, 7) = Despacho$
           GRID.TextMatrix(i&, 8) = Aduana$
           Call art_impo.GrabarDatosAduanaDespFecha(LOTE$, NUBORE)
        Next i&
    Else 'QUE NO APROBO
        GoTo 99
    End If
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
    End If
99  Set VENTANA = Nothing
End Sub

Private Sub Form_Load()

    '
    Call CREAR_STORED_PROCEDURE
    If Label10(0).Caption = "B.Recepción:" Then
       Command5.Visible = True
    End If
    '
    Text2 = 0
    Text6 = -1
    Text7 = 4500 '4000
    Text8 = 4500 '1800
    PDES$ = Control$(IDENTER$, "PRIETIQ1")
    If PDES$ = "" Then PDES$ = Control$("*", "PRIETIQ1")

    
    Screen.MousePointer = 1
End Sub

Private Sub GRID_DblClick()
    Exit Sub
'    i& = GRID.MouseRow
'    j& = GRID.MouseCol
'    If i& = 0 Or i& > GRID.Rows Then Exit Sub
'    '
'    On Error Resume Next
'    GRID.Row = i&
'    GRID.COL = j&
'    On Error Resume Next
'
'    REG& = i&
'    If REG& < 1 Then
'        Exit Sub
'    End If
'
'    Dim MSF As MSFlexGrid
'    Set MSF = GRID
'    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
'
'    If TRIM$(Label10(0).Caption) <> "O.Fabricación" Then
'        '****RECEPCION****
'        Unload VENTABNEW
'        VENTANA.Inicializar = 0
'        CAMPOS$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A0;D.Trans/f8;Lote Proveedor/A20"
'        VENTANA.Campox = CAMPOS$
'        VENTANA.UtilizaFormula = 1
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.CampEditables = "4;8"
'        VENTANA.AgregaRegistro = 0
'        VENTANA.UtilizaFormula = 0
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.HabilitaInsertar = 0
'        VENTANA.NoMostrarBotonTilde = 0
'        VENTANA.TITULO = ""
'        VENTANA.Inicializar = 1
'        VENTABNEW.Height = 9500
'        VENTABNEW.MSF.Font = "Roboto Mono"
'        VENTABNEW.MSF.Font.Size = 8
'        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
'        '****************************************************************
'        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
'
'        VENTABNEW.Show 1
'        If VENTABNEW.APROBADO% > 0 Then
'           Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF, 1)
'    '       MSF.COL = 1
'    '       MSF.Sort = 3
'        End If
'        Set VENTANA = Nothing
'        Unload VENTABNEW
'        Label2 = CANTFILAS_SELECCIONADAS(GRID, "*")
'    Else
'        Unload VENTABNEW
'        VENTANA.Inicializar = 0
'        CAMPOS$ = "N°OF/A0;O.Fab/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A12;D.Trans/f8"
'        VENTANA.Campox = CAMPOS$
'        VENTANA.UtilizaFormula = 1
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.CampEditables = "5"
'        VENTANA.AgregaRegistro = 0
'        VENTANA.UtilizaFormula = 0
'        VENTANA.HabilitaBorrar = 1
'        VENTANA.HabilitaInsertar = 0
'        VENTANA.NoMostrarBotonTilde = 0
'        VENTANA.TITULO = ""
'        VENTANA.Inicializar = 1
'        VENTABNEW.Height = 9500
'        VENTABNEW.MSF.Font = "Roboto Mono"
'        VENTABNEW.MSF.Font.Size = 8
'        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
'        '****************************************************************
'        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
'
'        VENTABNEW.Show 1
'        If VENTABNEW.APROBADO% > 0 Then
'           Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF, 1)
'    '       MSF.COL = 1
'    '       MSF.Sort = 3
'        End If
'        Set VENTANA = Nothing
'        Unload VENTABNEW
'
'
'        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
'        '****************************************************************
'        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
'
'    End If
'    Label2 = CANTFILAS_SELECCIONADAS(GRID, "*")

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    i& = GRID.MouseRow
    j& = GRID.MouseCol

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
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If j& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
       Label2 = CANTFILAS_SELECCIONADAS(GRID, "*")
    End If
End Sub

Private Sub GRID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If GRID.Rows <= 1 Then Exit Sub
  With GRID
    ' verifica que el GRIDlexgrid contiene filas
    If .MouseRow > 0 And .MouseCol >= 0 Then
        ' le asigna el dato con TextMatrix
        .ToolTipText = VectorToolTipText(.MouseCol)
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub MC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MC.Visible = False
End Sub

Private Sub Label10_Change(Index As Integer)
   '
   
End Sub


Private Sub mnuasigDesvTransitorio_Click()

End Sub

Private Sub mnuConfigurar_Click()
   If DERACCE%("STPRECEP", "Configuración Procesos de Recepción") >= 2 Then
     OPRECEP07.Show 1
   End If
End Sub

Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub mnuMenu2_Click()
  MENU2% = 2
  Hide
End Sub
Private Sub mnuMenu3_Click()
   Screen.MousePointer = 11
   Dim MSFF As MSFlexGrid
   Set MSFF = Me.GRID
   If TRIM$(TXTBUSCAR.TEXT) <> "" Then Set MSFF = Me.MSF_2
   FIN& = MSFF.Rows - 1
   MSFF.Redraw = False
   For i& = 1 To MSFF.Rows - 1
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(MSFF, vbYellow, i&)
   Next i&
   MSFF.Redraw = True
   Screen.MousePointer = 1
End Sub


Private Sub MSF_2_Click()
    Static MODO  As Boolean
    i& = MSF_2.MouseRow
    j& = MSF_2.MouseCol

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
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If j& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.Row, 0, 0, 0)
    End If
     
End Sub


'Private Sub msf2_Click()
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



Private Sub MSF_2_DblClick()
  Call GRID_DblClick
End Sub



Private Sub Option1_Click()
   Call Text5_Change
   TXTBUSCAR = ""
End Sub

Private Sub Option2_Click()
   Call Text5_Change
   TXTBUSCAR = ""
End Sub

Private Sub Option5_Click()
    Call Text5_Change
    TXTBUSCAR = ""
End Sub

Sub Text5_Change()
    Dim MSF As MSFlexGrid
    Set MSF = GRID
'    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
    
    Dim obj As New RECORXET
    
    If Label10(0).Caption <> "O.Fabricación" Then
        If XVALO(Text5) < 1 Then
          Me.MSF_2.Rows = 1
          Me.GRID.Rows = 1
          Text1 = ""
          Text2 = ""
          Text3 = ""
          Text4 = ""
          Label2 = "": Label3 = "": Label6 = ""
          Exit Sub
        End If

        'RECEPCION
        CAMPOS$ = "Lote/A12;Código/A16;Descripción/A24;Cantidad/F9.1;CI/A0;Fecha/D8;Despacho/A32;Aduana/A32"
        VENTANA.Campox = CAMPOS$
        
        Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me)
        Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me)
        FRMZELECHO.Caption = ""
    
        Set RS1 = obj.RS_BoleReceYDespaimpo(XVALO(Text5))
        Call Carga_MSF_Recordset(RS1, CAMPOS$, MSF)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
       
    Else 'FABRICACION
    '****************************************************************
        If XVALO(Text5) < 1 Then Exit Sub
        CAMPOS$ = "N°OF/A0;O.Fab/A16;Código/A16;Descripción/A24;Cantidad/F10.1;CI/A0;Revisión/A12;D.Trans./F8"
        VENTANA.Campox = CAMPOS$
        
        Call CARGA_ENCABEZADO(Me.MSF_2, CAMPOS$, Me)
        Call CARGA_ENCABEZADO(Me.GRID, CAMPOS$, Me)
        FRMZELECHO.Caption = ""
    
        Set RS1 = obj.RS_DETAORFAB(XVALO(Text5))
        Call Carga_MSF_Recordset(RS1, CAMPOS$, MSF)
        Set obj = Nothing
        RS1.Close
        Set RS1 = Nothing
        Call APLICAR_REVISION_FORMATEADA(MSF, 7)
    End If
     
End Sub

'Sub APLICAR_REVISION_FORMATEADA(MSF As MSFlexGrid, COL%)
'    For i& = 1 To MSF.Rows - 1
'       VERSIOX$ = MSF.TextMatrix(i&, COL%)
'       MSF.TextMatrix(i&, COL%) = FormatoEYMRevision$(VERSIOX$)
'
'    Next i&
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
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


