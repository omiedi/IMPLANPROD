VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PENDIENTMDB 
   Caption         =   "Consultas Pendientes"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Restaurar-Color"
      Height          =   255
      Left            =   8760
      TabIndex        =   32
      Top             =   5360
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "Consulta:"
      Top             =   5385
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4020
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "Instrucciones:"
      Top             =   5385
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edición de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   105
      TabIndex        =   18
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Solicitante"
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   945
         Width           =   1215
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2535
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   3480
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Teléfono"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Consulta Nº"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Razon Social"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Cuenta"
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Modificables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6480
      TabIndex        =   8
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   0
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   22
         TabIndex        =   11
         ToolTipText     =   "Doble Click Para Limpiar "
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   22
         TabIndex        =   10
         ToolTipText     =   "Doble Click Para Limpiar"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   22
         TabIndex        =   9
         ToolTipText     =   "Doble Click Para Limpiar"
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   1
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Aprobó"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Verificado con"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Instalado con"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   0
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Width           =   10095
      Begin VB.CommandButton Command5 
         Caption         =   "Cierra Consulta Activa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   4020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5640
      Width           =   6200
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5640
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorBkg    =   16777215
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
   End
   Begin VB.Label TEPRUEBA 
      Caption         =   "Label3"
      Height          =   120
      Left            =   6510
      TabIndex        =   34
      Top             =   5460
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   0
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "PENDIENTMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  


Public Sub Aplicar_skin(ByVal FORMULARIO As Form)
    FORMULARIO.Skin1.LoadSkin App.Path & "\SKINS\Web-II.skn"
    FORMULARIO.Skin1.ApplySkin FORMULARIO.hwnd
End Sub


Private Sub Combo1_Change(Index As Integer)
  Text14(Index) = Combo1(Index).TEXT

End Sub

Private Sub Combo1_Click(Index As Integer)
  Text14(Index) = Combo1(Index).TEXT

End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   NUCOTE$ = TRIM$(Text7)
   If NUCOTE$ = "" Then GoTo 99
   
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   Call ACTUREGISTRO("COTELEF", "INSTALADO_CON", CONDI$, Text14(0).TEXT, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "VERIFICADO_CON", CONDI$, Text14(1).TEXT, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "APROBADO_POR", CONDI$, Text14(2).TEXT, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "DetaConsu", CONDI$, Text1.TEXT, RAFE&, "NOMESS")
   Call ACTUREGISTRO("COTELEF", "DiagInstru", CONDI$, Text2.TEXT, RAFE&, "NOMESS")
   Call CARGAGRILLA(XVALO(Label1), XVALO(Label2))
  
    

99 Command1.Enabled = True

End Sub

Private Sub Command3_Click()
   Call CARGAGRILLA(XVALO(Label1), XVALO(Label2))
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    NUCOTE$ = TRIM$(Text7)
    If NUCOTE$ = "" Then
       Call COMUNI("Número de Consulta No Válido")
       GoTo 99
    End If
    CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
    If CantRegistros&("COTELEF", CONDI$, "NOMESS") < 1 Then
       Call COMUNI("Número de Consulta No Válido")
       GoTo 99
    End If
    '
    If MsgBox("¿ Realmente Desea Cerrar La Consulta Activa ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmación de Coerre") = 6 Then

        FEXI% = HOY(HO$)
        FENUE$ = FETEXCOR1$(FEXI%)
        Call ACTUREGISTRO("COTELEF", "Fech_Cierre", CONDI$, FENUE$, RAFE&, "NOMESS")
        '
        Call COMUNI("Consulta Telefónica Cerrada")
    '
    End If
99  Command5.Enabled = True
End Sub

Private Sub Form_Load()
      Call Aplicar_skin(Me)
      
      Call DESHASFECHA(DES%, HAS%, PERIO$)
      If PERIO$ = "" Then Exit Sub
      Call CREACAMPO("ADMIFLEX", "COTELEF", "INSTALADO_CON", 10, 16)
      Call CREACAMPO("ADMIFLEX", "COTELEF", "VERIFICADO_CON", 10, 16)
      Call CREACAMPO("ADMIFLEX", "COTELEF", "APROBADO_POR", 10, 16)
      Call CREACAMPO("ADMIFLEX", "SEGUIMIENTO", "ANOTADOR", 12, 0)
      Call CREACAMPO("ADMIFLEX", "SEGUIMIENTO", "Nume_Cotel", 10, 16, 1)

      Screen.MousePointer = 11
      JJ& = 0
      '
      Label1 = TRIM$(Str(DES%))
      Label2 = TRIM$(Str(HAS%))
      Call CARGAGRILLA(DES%, HAS%)
      Screen.MousePointer = 1


End Sub
Private Sub Text14_DblClick(Index As Integer)
  Text14(Index).TEXT = ""

End Sub


Private Sub text7_Change()
   'MUESTRA DETALLE DE LA CONSULTA SEGUN FILA SELECCIONADA
   NUCOTE$ = TRIM$(Text7)
   
   CONDI$ = "Nume_Cotel = '" & NUCOTE$ & "'"
   Text1 = VALOBADA("COTELEF", "DetaConsu", CONDI$, "NOMESS")
   Text2 = VALOBADA("COTELEF", "DiagInstru", CONDI$, "NOMESS")
   
   

   Text4 = VALOBADA("COTELEF", "NROCLIE", CONDI$, "NOMESS")
   NC% = XVALO(Text4)
   Text6 = RASOCLI$(NC%)
   Text17 = TEL1CLI$(NC%)
   Text18 = VALOBADA("COTELEF", "Solicitante", CONDI$, "NOMESS")
   '
   Call CARGACOMBOUSER
   Text14(0).TEXT = VALOBADA("COTELEF", "INSTALADO_CON", CONDI$, "NOMESS")
   Text14(1).TEXT = VALOBADA("COTELEF", "VERIFICADO_CON", CONDI$, "NOMESS")
   Text14(2).TEXT = VALOBADA("COTELEF", "APROBADO_POR", CONDI$, "NOMESS")

End Sub

Private Sub MSF_Click()
   Text7 = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   '
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   
   I& = MSF.MouseRow
   J& = MSF.MouseCol
   '
   Call SELECCION_COLUMNA_0(MSF, I&, J&)
   '
End Sub
Private Sub MSF_DblClick()
   
   I& = MSF.MouseRow
   J& = MSF.MouseCol
   If I& = 0 Or I& > MSF.Rows Then Exit Sub
   '
   MSF.Row = I&
   MSF.COL = J&

   NUCONSU$ = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   SEGUIMIENTO.Text5 = NUCONSU$
   SEGUIMIENTO.Show 1
   
End Sub

Sub CARGAGRILLA(DES%, HAS%)
      CONDI$ = "  (Fech_Cierre <= '30/NOV/1999' OR Fech_Cierre IS NULL OR FECH_CIERRE = '')"
      If DES% > 0 And HAS% > 0 Then
         DESDE = FETEXCOR1$(DES%)
         HASTA = FETEXCOR1$(HAS%)
         CONDI$ = CONDI$ + " AND FechEmis >= '" & DESDE & "' "
         CONDI$ = CONDI$ & " AND FechEmis <= '" & HASTA & "' "
      End If

      CONDI$ = CONDI$ + " ORDER BY  FechEmis "
      MSF.Cols = 9
      NUCOLU% = 1: TEPRUEBA = "CT-0000000-OD": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 2: TEPRUEBA = "123456": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 3: TEPRUEBA = "ABCDEFGHIJKLLLMNÑOPQRSTUV": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 4: TEPRUEBA = "ABCDEFGHIJKLLLMNÑOPQRSTUVWXYZAAAAAAAA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 6: TEPRUEBA = "ABCDEFGH1JKLMNOPQR": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width ': GRID.ColAlignment(NUCOLU% - 1) = 1: GRID.TextMatrix(0, NUCOLU% - 1) = "Referencia"
      NUCOLU% = 7: TEPRUEBA = "AAAAAAAAAAAAAA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 8: TEPRUEBA = "AAAAAAAAAAAAAA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      NUCOLU% = 9: TEPRUEBA = "AAAAAAAAAAAAAA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width
      'APARENTEMENTE LO DE ARRIBA NO FUNCIONA SI SE PASA EL ENCABEZADO COMO ESTA ABAJO POR ESO AGRANDE LOS ESPACIOS
      Encabezado$ = "< C-T                 |>Cuenta   |<R.Social                                                   |<Referencia                                                           |<Insatalado Con             |<Verificado con              |<Aprobó           "
      PENDIENTES.MSF.MergeCells = flexMergeRestrictColumns
      PENDIENTES.MSF.MergeCol(3) = True

      Call CARGALISTA_LIST_GRID(PENDIENTES.MSF, "COTELEF", "Nume_Cotel/A12;NROCLIE/F5+DEUDOR12/A24;Referencia/A36;INSTALADO_CON/A14;VERIFICADO_CON/A14;APROBADO_POR/A14", CONDI$, "", Encabezado$, 1)
      For I& = 1 To MSF.Rows - 1
        INSTALADOX$ = TRIM$(MSF.TextMatrix(I&, 4))
        VERIFICADOX$ = TRIM$(MSF.TextMatrix(I&, 5))
        APROBADOX$ = TRIM$(MSF.TextMatrix(I&, 6))
        If APROBADOX$ <> "" Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HFFFF00, I&): GoTo 50
        If VERIFICADOX$ <> "" Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HC0E0FF, I&): GoTo 50
        If INSTALADOX$ <> "" Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HC0C0C0, I&): GoTo 50
50   Next I&

End Sub
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   I& = MSF.MouseRow
   J& = MSF.MouseCol
   If I& = 0 Or I& > MSF.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       MSF.Row = I&
       MSF.COL = J&
       MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

       NUCONSU$ = TRIM$(MSF.TextMatrix(MSF.Row, 0))
       SEGUIMIENTO.Text5 = NUCONSU$
       SEGUIMIENTO.Show 1

'       If FECH% < 1 Then Exit Sub
'       FECHAX = FETEXCOR1$(FECH%)
'       COMOVI$ = TRIM$(MSF.TextMatrix(I&, 1))
'       DOPRI$ = TRIM$(MSF.TextMatrix(I&, 2))
'       DOSEC$ = TRIM$(MSF.TextMatrix(I&, 3))
'       IDLOT$ = TRIM$(MSF.TextMatrix(I&, 4))
'       CONDI$ = "FECHMOV = '" & FECHAX & "' AND CODIMOVI = '" & COMOVI$ & "' AND DOPRILAR = '" & DOPRI$ & "' AND DOSECLAR = '" & DOSEC$ & "' And IDENLOTE = '" & IDLOT$ & "'"
'
'       FECHAYHORA = VALOBADA("MOVISTO", "FEREMOVI", CONDI$, "NOMESS")
'       MsgBox "Fecha y Hora Real del Movimiento  " & vbCrLf & vbCrLf & FECHAYHORA, , "FLEXOFT"
       MSF.CellBackColor = RGB(255, 255, 255)
   End If
End Sub

Sub CARGACOMBOUSER()

  NC% = XVALO(Text4)

  For I& = 0 To 2
    Combo1(I&).Clear
    IDLAN$ = HORARCLI$(NC%)
    RFF$ = RECFILT$("WEBUSER", 4, IDLAN$)
    For KKU& = 1 To Len(RFF$) Step 4
       REGI& = CVS(Mid$(RFF$, KKU&, 4))
       XX$ = REGLEIDO$("WEBUSER", REGI&)
       TTXX$ = AJUSTI$(Mid$(XX$, 3, 30), 64) ' + TRIM$(Mid$(XX$, 33, 12))
       Combo1(I&).AddItem TTXX$
    Next KKU&
    On Error Resume Next
    Combo1(I&).ListIndex = 0
    Combo1(I&).Refresh
    On Error GoTo 0
  Next I&

End Sub
Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command4_Click()
   SEGUIMIENTO.Text5 = Text7
   SEGUIMIENTO.Show 1
End Sub

Sub COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF As MSFlexGrid, COLOR As Variant, Optional REGISTRO)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      IROW% = XVALO(REGISTRO)
    Else
      IROW% = MSF.Row
    End If
    
    .Row = IROW%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    For ICOL% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .COL = ICOL%
          .CellBackColor = COLOR
    Next ICOL%
   End With
End Sub
Private Sub MSF_KeyPress(KeyAscii As Integer)
   Call Command4_Click
End Sub
Private Sub MSF_SelChange()
   Text7 = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   '
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   '
   I& = MSF.Row
   J& = MSF.COL
   '
   Call SELECCION_COLUMNA_0(MSF, I&, J&)
   
End Sub
Sub SELECCION_COLUMNA_0(MSF As MSFlexGrid, REGISTRO&, COLUMNA&)
   I& = REGISTRO&
   J& = COLUMNA&
   '
   'VALIDO
   If I& = 0 Or I& > MSF.Rows Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.COL = 0
   For H& = 0 To MSF.Rows - 1
     MSF.Row = H&
     MSF.CellBackColor = &H8000000F
   Next H&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.Row = I&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

End Sub

