VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PRINTETI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6960
   ClientLeft      =   2.45790e5
   ClientTop       =   600
   ClientWidth     =   10875
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Prueba Imprimir Etiquetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "PRINTETI.frx":0000
      TabIndex        =   37
      ToolTipText     =   "Imprime etiquetas Items Seleccionados"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9400
      Picture         =   "PRINTETI.frx":030A
      TabIndex        =   30
      ToolTipText     =   "Graba Impresora Seleccionada"
      Top             =   6480
      Width           =   1400
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6480
      Width           =   2865
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   10155
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   27
      Top             =   5160
      Width           =   630
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir Etiquetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   8760
      Picture         =   "PRINTETI.frx":0614
      TabIndex        =   26
      ToolTipText     =   "Imprime etiquetas Items Seleccionados"
      Top             =   5160
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   550
      Width           =   10695
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar Cambios en Grilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         Picture         =   "PRINTETI.frx":091E
         TabIndex        =   36
         ToolTipText     =   "Graba Datos de Grilla"
         Top             =   120
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         Caption         =   "Solo Trazables"
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
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   180
         Width           =   1980
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar Todos"
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
         Height          =   240
         Left            =   2760
         TabIndex        =   24
         Top             =   180
         Width           =   1500
      End
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
      Top             =   240
      Width           =   1215
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
      TabIndex        =   17
      Top             =   240
      Width           =   175
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
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   " "
      Top             =   240
      Width           =   3495
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
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   " "
      Top             =   240
      Width           =   1215
   End
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
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   " "
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   5160
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
         Picture         =   "PRINTETI.frx":0C28
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   160
         Width           =   705
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
      Left            =   1320
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PRINTETI.frx":0F32
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "PRINTETI.frx":1166
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   4035
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10700
      _ExtentX        =   18865
      _ExtentY        =   7117
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   4035
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   10700
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   2880
      TabIndex        =   34
      Top             =   5880
      Width           =   610
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Filas Seleccionadas:"
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
      Index           =   1
      Left            =   165
      TabIndex        =   33
      Top             =   5880
      Width           =   1785
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Selección de Impresora de Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      TabIndex        =   29
      Top             =   6435
      Width           =   2085
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social Proveedor"
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
      Index           =   3
      Left            =   4200
      TabIndex        =   22
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cta.Prov."
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
      Index           =   2
      Left            =   2985
      TabIndex        =   21
      Top             =   0
      Width           =   735
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
      TabIndex        =   20
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B.Recepción:"
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
      Index           =   0
      Left            =   135
      TabIndex        =   19
      Top             =   0
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   9480
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -120
      TabIndex        =   5
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
   Begin VB.Menu mnuTrans 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuasigDesvTransitorio 
         Caption         =   "Asignar Desvío Transitorio a Fab/Rec Activa"
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
   Begin VB.Menu mnuconfiguracion 
      Caption         =   "Configuración"
   End
End
Attribute VB_Name = "PRINTETI"
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




Sub Alta_Actu_Desvio_Transitorio(DESVTRAN, PARTEX$, CI1%, REFERENCIA$, Optional NUMESERIE$, Optional ID_TRAZAXSERIE)
   '
   If TRIM$(NUMESERIE$) <> "" Then
      'SI SE DIO DE ALTA A MODO GENERAL NO SE IDENTIFICA POR NUMERO DE SERIE
      CONDI$ = "DESV_TRANS = " & DESVTRAN & " AND (NUMEROSERIE = '' OR NUMEROSERIE IS NULL)"
      CRG& = XVALO(CantRegistros("DES_TRANS", CONDI$, "NOMESS"))
      If CRG& > 0 Then
        NUMESERIE$ = ""
      End If
   End If
   '
   If XVALO(DESVTRAN) < 1 Then
      DESVTRAN = PRONUM_DESVIO_TRANSITORIO
      Call CreaRegistro("DES_TRANS", "DESV_TRANS", XVALO(DESVTRAN), "NOMESS")
   End If
   '
   SQLX$ = "UPDATE DES_TRANS SET IDENLOTE = '" & PARTEX$ & "'"
   SQLX$ = SQLX$ + " , COD_INTE = " & CI1%
   SQLX$ = SQLX$ + " , COD_EXTE = '" & CODEXT$(CI1%) & "'"
   SQLX$ = SQLX$ + " , REF_DESVIO = '" & REFERENCIA$ & "'"
   SQLX$ = SQLX$ + " , NUMEROSERIE = '" & NUMESERIE$ & "'"
   SQLX$ = SQLX$ + " , IDSUBOR = ''"
   SQLX$ = SQLX$ + "  WHERE DESV_TRANS = " & DESVTRAN
   FLEXCONN.Execute SQLX$
   '
   If XVALO(ID_TRAZAXSERIE) > 0 Then
    SQLX$ = "UPDATE TRAZAXSERIE SET DESV_TRANS = '" & DESVTRAN & "'"
    SQLX$ = SQLX$ + "  WHERE TRAZAXSERIE_ID = " & ID_TRAZAXSERIE
    FLEXCONN.Execute SQLX$
   End If
End Sub

Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Sub CARGACOMBOPROPIEDADESQR(COMBO As comboBox)
   'LOS COMENTADOS DAN ERROR
       'COMBO.AddItem "All_1D"
       COMBO.AddItem "AZTEC"
       'COMBO.AddItem "CODABAR"
       COMBO.AddItem "CODE_128"
       COMBO.AddItem "CODE_39"
       'COMBO.AddItem "CODE_93"
       COMBO.AddItem "DATA_MATRIX"
       'COMBO.AddItem "EAN_13"
       'COMBO.AddItem "EAN_8"
       'COMBO.AddItem "ITF"
       'COMBO.AddItem "MAXICODE"
       'COMBO.AddItem "MSI"
       'COMBO.AddItem "PLESSEY"
       COMBO.AddItem "QR_CODE"
       'COMBO.AddItem "RSS_14"
       'COMBO.AddItem "UPC_A"
       'COMBO.AddItem "UPC_E"
       'COMBO.AddItem "UPC_EAN_EXTENSION"
        
       COMBO.TEXT = "QR_CODE"
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


Private Sub Command1_Click()
   Dim obj As New RECORXET
   If Label10(0).Caption <> "O.Fabricación" Then
        Campos$ = "N°.Recep./f9;N° Prov./F7;Proveedor/A24;Fecha/D8"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        FRMZELECHO.Caption = ""
        
        Set rs1 = obj.RS_TODARECE()
        Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        
        FRMZELECHO.Caption = "LISTA DE RECEPCIONES"
        FRMZELECHO.Width = 12000
        FRMZELECHO.Show 1
        If XVALO(RESP_ZELE_VECT(1)) > 0 Then
          Text5 = XVALO(RESP_ZELE_VECT(1))
          Text1 = TRIM$(RESP_ZELE_VECT(4))
          Text3 = TRIM$(RESP_ZELE_VECT(2))
          Text4 = TRIM$(RESP_ZELE_VECT(3))
        End If
        
   Else
        Campos$ = "N°.O.Fab./A12;N° OF/A0;CI/A0;Código/A16;Descripción/A32;Cantidad/F10.1;Fecha/D8"
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        FRMZELECHO.Caption = ""
        
        Set rs1 = obj.RS_TODAOF()
        Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        '
        FRMZELECHO.Caption = "LISTA DE O.FABRICACIÓN"
        FRMZELECHO.Width = 12000
        
        FRMZELECHO.Show 1
        If XVALO(RESP_ZELE_VECT(1)) > 0 Then
          Text5 = XVALO(RESP_ZELE_VECT(1))
        End If
        
        '
   End If


End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   '
10 ES_BR_O_OF% = 0 'RECEPCION
   If Label10(0).Caption = "O.Fabricación" Then
     ES_BR_O_OF% = 1 'O.FABRICACION
   End If
   

   PORBULTO% = 0
   
   If ES_BR_O_OF% <= 0 Then
      CANTICOLUMROLLO% = XVALO(Control("PRINTETI", "CANTCOLU")) 'RECEPCION
      If Control("PRINTETI", "PORBULTO") = "SI" Then PORBULTO% = 1
      FORIMPRE$ = Control("BOLRECEP", "ETIQRECE")
   Else
      CANTICOLUMROLLO% = XVALO(Control("ORDEFABR", "CANTCOLU")) 'OF
      If Control("ORDEFABR", "PORBULTO") = "SI" Then PORBULTO% = 1
      FORIMPRE$ = Control("ORDEFABR", "ETIQUEOF")
   End If
   '
   'SETEO GRILLA ACTIVA
   Dim MSF As msFlexGrid
   Set MSF = GRID
   If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2

   If InStr(EPAC$, "PERGAMINO") > 0 Then GoTo 50
   
   If TRIM$(FORIMPRE$) = "" Then
      Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
      GoTo 99
   End If
   
   If CANTICOLUMROLLO% < 1 Then CANTICOLUMROLLO% = 1
   If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
   '
   PUERIMPRE$ = Control("*", "PRIETIQ1")
   If Control(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = Control(IDENTER$, "PRIETIQ1")
   Call SETPRINPORT(PUERIMPRE$)
   
   
   If Label10(0).Caption = "O.Fabricación" And InStr(EPAC$, "CAMBACERES") > 0 Then
      NORFA$ = "OF-" & FORMATOCON0$(XVALO(Text5), 6) & "-1"
      For i& = 1 To MSF.Rows - 1
        MARCA$ = TRIM$(MSF.TextMatrix(i&, 0))
        If MARCA$ = "*" Then
          Image1.Picture = LoadPicture("")
          LOTE$ = MSF.TextMatrix(i&, 1)
          PESOREALX# = XVALO(MSF.TextMatrix(i&, 6))
          METROSREALX# = XVALO(MSF.TextMatrix(i&, 5))
          CANCOPIAS& = XVALO(MSF.TextMatrix(i&, 4))

         Call CIERRE_CMB.PRINTETIQORFA_CAMBACERES(NORFA$, LOTE$, PESOREALX#, METROSREALX#, CANCOPIAS&)
        End If
      Next i&
      GoTo 99
   End If
   R& = 0
   '
50 Image1.Picture = LoadPicture("")
   '
   CANTSELEC = Label2 'CANTI&DAD DE FILAS SELECCIONADAS
   For i& = 1 To MSF.Rows - 1
     MARCA$ = TRIM$(MSF.TextMatrix(i&, 0))
     If MARCA$ = "*" Then
        Image1.Picture = LoadPicture("")
        'VER ACA A MODO GENERAL EL TEMA DEL TEXTO A GENERAR QR
        If ES_BR_O_OF% <= 0 Then 'BR
            COD$ = MSF.TextMatrix(i&, 2)
            CIXX$ = SAFETEXT$(CODINT%(COD$))
            REVI$ = MSF.TextMatrix(i&, 6)
            LOTE$ = MSF.TextMatrix(i&, 1)
            CANTI& = XVALO(MSF.TextMatrix(i&, 4))
            If InStr(EPAC$, "PERGAMINO") > 0 Then
              Call PrintEtiqRecepcionPergamino(LOTE$) 'IMPRIME Y PASA AL SIGUIENTE EN EL CASO DE PERGAMINO
              GoTo 80
            End If

        Else
            COD$ = MSF.TextMatrix(i&, 3) 'OF
            CIXX$ = SAFETEXT$(CODINT%(COD$))
            REVI$ = MSF.TextMatrix(i&, 7)
            LOTE$ = MSF.TextMatrix(i&, 2) ' EN ESTE CASO ES LA OF
            CANTI& = XVALO(MSF.TextMatrix(i&, 5))
        End If
        '
        'SI ES CONFIGURACION POR BULTO CALCULA LA CANTIDAD DE BULTOS
        If PORBULTO% > 0 Then
            'POR BULTO
            CANX# = CANTI&
            CANTI& = CANBULTOS(CODINT%(COD$), CANX#, "COMPRA")
        End If
        If CANTI& < 1 Then CANTI& = 1
        '
        If CANTI& > CANTICOLUMROLLO% Then
               'ESTO ES PARA CALCULAR SEGUN LA CANTIDAD DE ETIQUETAS COMO REPARTE EN LAS FILAS
               'EJ. SI ES LA ETIQUETA TIENE CONFIGURADO 3 POR FILA Y EL TOTAL A IMPRIMIR SON 20 REALIZA 7 IMPRESIONES DE 3
               SALDO& = CANTI& Mod CANTICOLUMROLLO%: If SALDO& > 0 Then SALDO& = 1
               CANTI& = Int(CANTI& / CANTICOLUMROLLO%) + SALDO&
        Else
            CANTI& = CANBULTOS(CODINT%(COD$), CAN#, "COMPRA")
        End If
        
        R& = R& + 1
        Label2 = R&
        Label3 = COD$
'        If ES_BR_O_OF% <= 0 Then 'BR
'          LOTE$ = REPLACAR(LOTE$, "B", "P")
'        End If
        '
        REVI$ = FormatoEYMRevision$(REVI$)
        'TEXTOQR$ = COD$ & " Rev." & REVI$ & " " & LOTE$
        'el textoqr se arma con el codigo interno el n.revision y el lote sin los ceros centrales
        'ej:  727 13 OF-220-1     esto es para que tenga menos caractreres por problemas en la lectura
        TEXTOQR$ = CIXX$ & " " & REVI$ & " " & FORMATO_BR_OF_SIN_0$(LOTE$)
        If ES_BR_O_OF% <= 0 Then 'BR
           IMAG$ = "IMAGRECE_QR.JPG"
           IDENTIFICADOR$ = "QR_RECEPCION"
        Else                      'OF
           IMAG$ = "IMAGORFA_QR.JPG"
           IDENTIFICADOR$ = "QR_ORFA"
        End If
        '
        On Error Resume Next
        Kill "C:\FLEXOFT\" & IMAG$
        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
        QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, Combo1.TEXT, 0, 0))
        Call T_Espera(3)
        Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
        Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)

        User% = USNBR% Mod 100
        If User% < 1 Then User% = 1
        'FILTX$ = TRIM$(LOTE$) & ";" & TRIM$(COD$) & ";" & "Rev." & TRIM$(REVI$) & ";" & IDENTIFICADOR$ & ";" & TRIM$(Str$(User%))
        FILTX$ = TRIM$(COD$) & ";" & "Rev." & REVI$ & ";" & TRIM$(LOTE$) & ";" & IDENTIFICADOR$ & ";" & TRIM$(Str$(User%))
        Call STDFORM(FORIMPRE$, FILTX$, "PRINT", "", 0, CANTI&)
     End If
80 Next i&
   Call SETPRINPORT("RESTORE")
   '
99 Command2.Enabled = True

End Sub
Private Sub Command3_Click()
  Dim MSF As msFlexGrid
  Set MSF = GRID
  If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
  Call numerargrilla(MSF, 0)

End Sub



Private Sub Command4_Click()
    Call GRACONTROL(IDENTER$, "PRIETIQ1", Combo2.TEXT)
    Call GRACONTROL("*", "PRIETIQ1", Combo2.TEXT)


End Sub

Private Sub Command5_Click()
    Dim MSF As msFlexGrid
    Set MSF = GRID
    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    'VALIDO QUE AL MENOS HAYA UNA RECEPCION EN LA GRILLA
    On Error Resume Next
    BR$ = TRIM$(MSF.TextMatrix(1, 1))
    If BR$ = "" Then Exit Sub
    On Error GoTo 0
    
'   CAMPOS$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A0;D.TRANS./F8;Lote Proveedor"
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    
    For i& = 1 To MSF.Rows - 1
       BR$ = MSF.TextMatrix(i&, 1)
       LOTE$ = MSF.TextMatrix(i&, 8)
       If InStr(EPAC$, "ENVAPLAST") > 0 Then CANTIBULTOS& = XVALO(MSF.TextMatrix(i&, 9))
       SQLX$ = "UPDATE BOLRECEP SET LoteCol_Prov = '" & LOTE$ & "'"
       If InStr(EPAC$, "ENVAPLAST") > 0 Then SQLX$ = SQLX$ + ", CANTIDADBULTOS= " & CANTIBULTOS&
       SQLX$ = SQLX$ + " WHERE IdenLote = '" & BR$ & "'"
       FLEXCONN.Execute SQLX$
    Next i&
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Transacción Realizada Exitosamente")
    End If

End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   '
   FORIMPRE$ = Control("BOLRECEP", "ETIQRECE")
   '
   If TRIM$(FORIMPRE$) = "" Then
      Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
      GoTo 99
   End If
   '
   PUERIMPRE$ = Control("*", "PRIETIQ1")
   If Control(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = Control(IDENTER$, "PRIETIQ1")
   Call SETPRINPORT(PUERIMPRE$)
   '
   Image1.Picture = LoadPicture("")
   '
            COD$ = "12345"
            REVI$ = "0.25"
            LOTE$ = "BR-123-1"
            TEXTOQR$ = COD$ & " " & REVI$ & " " & LOTE$
            IMAG$ = "IMAGRECE_QR.JPG"
            IDENTIFICADOR$ = "QR_RECEPCION"
            '
            On Error Resume Next
            Kill "C:\FLEXOFT\" & IMAG$
             On Error GoTo 0
            'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
            QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, Combo1.TEXT, ANCHOX%, ALTOX%))
            Call T_Espera(3)
            Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
            Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)
    
            User% = USNBR% Mod 100
            If User% < 1 Then User% = 1
            'FILTX$ = TRIM$(LOTE$) & ";" & TRIM$(COD$) & ";" & "Rev." & TRIM$(REVI$) & ";" & IDENTIFICADOR$ & ";" & TRIM$(Str$(User%))
            FILTX$ = TRIM$(COD$) & ";" & "Rev." & REVI$ & ";" & TRIM$(LOTE$) & ";" & IDENTIFICADOR$ & ";" & TRIM$(Str$(User%))
            Call STDFORM(FORIMPRE$, FILTX$, "PRINT", "", 0, 3)
            Call SETPRINPORT("RESTORE")
   '
99          Command6.Enabled = True

End Sub

Private Sub Form_Load()

    '
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
    Combo2.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo2.TEXT = PDES$
      End If
    Next
    Call CARGACOMBOPROPIEDADESQR(Combo1)
    Call CREATABLA_ARCHIVOS_QR
    
    If XVALO(Control("BOLRECEP", "ETIQTRAZ")) > 0 Then
       Option1.Value = True 'SOLO TRAZABLES
    End If
    A = FormatoEYMRevision$("ABC 1 2 3")
    Screen.MousePointer = 1
End Sub

Sub PrintEtiqRecepcionPergamino(LoteRecepcion$)


    If LoteRecepcion$ = "" Then Exit Sub
    Static FORIPRE$
    Static USUARIO%
    Static EPAC$
    If Not PrintEtiqRecepcionPergaminoyaEjecutado Then
        FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
        If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then
            Call COMUNI("Falta Formulario de Impresión de Etiqueta")
            Exit Sub
        End If
        
        Call CREATABLA_ARCHIVOS_QR
        USUARIO% = USNBR% Mod 100
        EPAC$ = UCase$(TRIM$(EMPREAC$))
        PFAC$ = Control$(IDENTER$, "PORTSTIK")
        If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
        If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
        Call SETPRINPORT(PFAC$)
    End If
    ' OBTIENE DATOS DEL LOTE
    ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LoteCol_Prov"
    ConsultaSql$ = ConsultaSql$ + " from bolrecep with(nolock) where idenlote='" & LoteRecepcion$ & "'"
    Set Rst = READSET(ConsultaSql$)
    GoTo 70
        
    ' OBTIENE DATOS DE LA RECEPCION
65  ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LOTECOL_PROV from bolrecep with(nolock) where numebore='" & CStr(NroRecepcion) & "'"
    Set Rst = READSET(ConsultaSql$)
    
70  With Rst
        Do While Not .EOF
            CodInte% = XVALO(!COD_INTE)
            If InStr(1, EPAC$, "CAMBACERES") > 0 And TIMATE%(CodInte%) = 41 Then GoTo 80 'ACA VIENE AL FINAL de la rece.provisoria
            Codigo$ = CODEXT$(CodInte%)                                                  'si es bobina no imprime por que ya imprimio etiqueta
            Descripcion$ = DESCRIT0$(CodInte%)
            LOTE$ = SAFETEXT$(!idenlote)
            LoteProveedor$ = ""
            If InStr(1, EPAC$, "PERGAMINO") > 0 Then
              LoteProveedor$ = SAFETEXT$(!LOTECOL_PROV) 'ES EL DE CAMBACERES / SI ES DE OTRO LO VA A DEVOLVER VACIO
              If TRIM$(LoteProveedor$) = "" Then LoteProveedor$ = "--"
            End If
            
            Largo$ = "0": If XVALO(!M_Largo) > 0.0005 Then Largo$ = TRIM$(FORMATNUM$(!M_Largo, "F12.1"))
            Peso$ = "0": If XVALO(!Peso) > 0.005 Then Peso$ = TRIM$(FORMATNUM$(!Peso, "F12.1"))
            Proveedor$ = SAFETEXT$(!Raso_Prov)
            Remito$ = SAFETEXT$(!NumRem_Prov)
            Remito$ = TRIM$(Left$(Remito$, 18)) & " " & LoteProveedor$ ' SI LO TIENE LO CONCATENA SI NO VA CONCATENA VACIO
            cantidadNominal# = !CantNom
            FechaRecepcion% = CVINT(!FechRecep)
            cantidadBultos% = XVALO(!cantidadBultos)
               
            TEXTOQR$ = Codigo$ & "|" & LOTE$
            IMAG$ = "IMAGRECE_QR.JPG"
            IDENTIFICADOR$ = "QR_RECEPCION"
            ANCHOX% = 0
            ALTOX% = 0
            '
            On Error Resume Next
            Kill LUCOM$ & IMAG$
            On Error GoTo 0
            'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
            QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
            Call T_Espera(3)
            Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
            Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
    
            If cantidadBultos% > 0 Then
                For i = 1 To cantidadBultos%
                    numeroBulto$ = CStr(i) & " / " & CStr(cantidadBultos%)
                    FILTX$ = Codigo$ & ";" & Descripcion$ & ";" & LOTE$ & ";" & TRIM$(FORMATNUM$(cantidadNominal#, "F12.1")) & ";" & LoteProveedor$ & ";" & Proveedor$ & ";" & Remito$ & ";" & IDENTIFICADOR$
                    Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                Next i
            Else
                    FILTX$ = Codigo$ & ";" & Descripcion$ & ";" & LOTE$ & ";" & TRIM$(FORMATNUM$(cantidadNominal#, "F12.1")) & ";" & LoteProveedor$ & ";" & Proveedor$ & ";" & Remito$ & ";" & IDENTIFICADOR$
                    Call STDFORM(FORIPRE$, FILTX$)
            End If
             
80          .MoveNext
        Loop
    End With
        
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    If PFAC$ <> "" Then Call SETPRINPORT("RESTORE")
End Sub

Private Sub GRID_DblClick()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "CAMBACERES") > 0 Then Exit Sub
    
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.row = i&
    GRID.col = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    Dim MSF As msFlexGrid
    Set MSF = GRID
    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
    If TRIM$(Label10(0).Caption) <> "O.Fabricación" Then
        '****RECEPCION****
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        If InStr(EPAC$, "ENVAPLAST") > 0 Then
           Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A10;D.TRANS./F8;Lote Proveedor/A20;Cant.Bob./F9"
           IncluirCampoCantidadBultos% = 1
           VENTANA.CampEditables = "8;9"
         Else
          Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A10;D.TRANS./F8;Lote Proveedor/A20"
          VENTANA.CampEditables = "4;8"
        End If

        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 0
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = ""
        
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 9500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
        '****************************************************************
        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
        
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% > 0 Then
           Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF, 1)
    '       MSF.COL = 1
    '       MSF.Sort = 3
        End If
        Set VENTANA = Nothing
        Unload VENTABNEW
        Label2 = CANTFILAS_SELECCIONADAS(GRID, "*")
    Else
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        Campos$ = "N°OF/A0;O.Fab/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A12;D.Trans/f8"
        VENTANA.Campox = Campos$
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.CampEditables = "5"
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 0
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = ""
        VENTANA.Inicializar = 1
        VENTABNEW.Height = 9500
        VENTABNEW.MSF.Font = "Roboto Mono"
        VENTABNEW.MSF.Font.Size = 8
        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
        '****************************************************************
        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
        
        VENTABNEW.Show 1
        If VENTABNEW.APROBADO% > 0 Then
           Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF, 1)
    '       MSF.COL = 1
    '       MSF.Sort = 3
        End If
        Set VENTANA = Nothing
        Unload VENTABNEW
        
    
        Call CopiarMsfAOtroMsf(MSF, VENTABNEW.MSF, 1)
        '****************************************************************
        Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
        
    End If
    Label2 = CANTFILAS_SELECCIONADAS(GRID, "*")

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    i& = GRID.MouseRow
    J& = GRID.MouseCol

    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.col = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.col = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If J& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.row, 0, 0, 0)
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
    If TRIM$(Text5) = "" Then
       Call COMUNI("No se ha Seleccionado Ningún Parte")
       Exit Sub
    End If
    '
    If Label10(0).Caption = "O.Fabricación" Then
       Campos$ = "N°OF/A0;O.Fab/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A12;D.TRANS./F8"
       CI_ENFILA% = 6
       PR_ENFILA% = 2
       DV_ENFILA% = 8
    Else
       Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A0;D.TRANS./F8"
       CI_ENFILA% = 5
       PR_ENFILA% = 1
       DV_ENFILA% = 7
    End If
10  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Width = 10000
    
    Call CopiarMsfAOtroMsf(Me.GRID, FRMZELECHO.msf2, 1, GRID.Width)
    FRMZELECHO.Show 1
    If TRIM(RESP_ZELE_VECT(PR_ENFILA%)) = "" Then Exit Sub
    PARTEX$ = TRIM(RESP_ZELE_VECT(PR_ENFILA%))
    DESVTRAN = XVALO(RESP_ZELE_VECT(DV_ENFILA%))
    CI1% = XVALO(RESP_ZELE_VECT(CI_ENFILA%))
    
    OBSERVAX$ = ""
    If DESVTRAN > 0 Then
       OBSERVAX$ = VALOBADA("DES_TRANS", "REF_DESVIO", "DESV_TRANS = " & DESVTRAN, "NOMESS")
    End If
    
30  REFERENCIA$ = TRIM$(InputBox$("Observaciones", "Ingrese Observaciones", OBSERVAX$))
    If REFERENCIA$ = "" Then GoTo 10
    Call Alta_Actu_Desvio_Transitorio(DESVTRAN, PARTEX$, CI1%, REFERENCIA$)
    
    NPARTX = XVALO(Text5)
    Text5 = ""
    Text5 = NPARTX
    GoTo 10

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
   Dim MSFF As msFlexGrid
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
    J& = MSF_2.MouseCol

    If (MSF_2.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.col = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.col = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If J& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.row, 0, 0, 0)
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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    Dim MSF As msFlexGrid
    Set MSF = GRID
    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2

    OP% = 0 ' POR DEFECTO MUESTRA TODOS
    If Option1.Value = True Then OP% = 1 'SOLO TRAZABLES IMPRIMIBLES
    Dim obj As New RECORXET
    
    If Label10(0).Caption <> "O.Fabricación" Then
        If XVALO(Text5) < 1 Then
          PRINTETI.MSF_2.Rows = 1
          PRINTETI.GRID.Rows = 1
          Text1 = ""
          Text2 = ""
          Text3 = ""
          Text4 = ""
          Label2 = "": Label3 = "": Label6 = ""
          Exit Sub
        End If

        'RECEPCION
        IncluirCampoCantidadBultos% = 0
        If InStr(EPAC$, "ENVAPLAST") > 0 Then
           Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A10;D.TRANS./F8;Lote Proveedor/A20;Cant.Bob./F9"
           IncluirCampoCantidadBultos% = 1
         Else
          Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A10;D.TRANS./F8;Lote Proveedor/A20"
        End If
        
        VENTANA.Campox = Campos$
        
        Call CARGA_ENCABEZADO(PRINTETI.MSF_2, Campos$, PRINTETI)
        Call CARGA_ENCABEZADO(PRINTETI.GRID, Campos$, PRINTETI)
        FRMZELECHO.Caption = ""
    
        Set rs1 = obj.RS_DETARECE(XVALO(Text5), OP%, IncluirCampoCantidadBultos%) 'si es envaplast incluye el campo CantidadBultos
        Call Carga_MSF_Recordset(rs1, Campos$, MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
       
    Else 'FABRICACION
    '****************************************************************
        If XVALO(Text5) < 1 Then Exit Sub
        If InStr(EPAC$, "CAMBACERES") > 0 Then
          Campos$ = "LOTE/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;METROS/F10.1;KGS/F10.1;CI/A0"
        Else
          Campos$ = "N°OF/A0;O.Fab/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A12;D.TRANS./F8"
        End If
        VENTANA.Campox = Campos$
        
        Call CARGA_ENCABEZADO(PRINTETI.MSF_2, Campos$, PRINTETI)
        Call CARGA_ENCABEZADO(PRINTETI.GRID, Campos$, PRINTETI)
        FRMZELECHO.Caption = ""
    
    
        If InStr(EPAC$, "CAMBACERES") > 0 Then
           NORFA$ = "OF-" & FORMATOCON0$(XVALO(Text5), 6) & "-1"
           Set rs1 = obj.RS_DETAORFAB_CAMBACERES(NORFA$)
        Else
           Set rs1 = obj.RS_DETAORFAB(XVALO(Text5))
        End If
        
        Call Carga_MSF_Recordset(rs1, Campos$, MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
    End If
    '*****************************************************'
    
    
    '*** SE DEBERA CARGAR PARA EL PARAMETRO  VENTANA.Campox = LOS CAMPOS DEL MSF BASE  ****
    '****************************************************************
    '
    '*** SE DEBERA CARGAR PARA LOS PARAMETROS CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)   ****
    '*** REALIZAR LA LLAMADA
    'Call CopiarMsfAOtroMsf(GRILLABASE, VENTADISE.MSF)
    '****************************************************************
       
     
    '   Dim obj As New RECORXET
    '   Set RS1 = obj.RS_DETARECE()
    '   Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2)
    '   Set obj = Nothing
    '   RS1.Close
    '   Set RS1 = Nothing
    '
    '   FRMZELECHO.Caption = "LISTA DE RECEPCIONES"
    '   FRMZELECHO.Width = 12000
    '   FRMZELECHO.Show 1
    '   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
    '     CODIMAQSEL$ = TRIM$(RESP_ZELE_VECT(1))
    '   End If
    
    
    
    
    
    '*****************************************************'
    
    
    
    
    'VENTABNEW.Show 1
    'If VENTABNEW.APROBADO% > 0 Then
    '   Call CopiarMsfAOtroMsf(VENTADISE.MSF, GRILLABASE)
    'End If
    'Set VENTANA = Nothing
    'End Function

End Sub


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
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = GRID.TextMatrix(i&, h&)
            Next h&
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
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub












Private Sub TXTCAMPODESCRIPTIVO_Change()

End Sub

Private Sub TXTCANFILA_Change()

End Sub


Private Sub TXTSERIE1_Change()

End Sub


