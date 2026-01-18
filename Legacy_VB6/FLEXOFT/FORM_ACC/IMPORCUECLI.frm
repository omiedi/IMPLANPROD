VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPORCUECLI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "'NUMERO DE TRANSACCION"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Planilla Origen/Detalle/Recibos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "IMPORCUECLI.frx":0000
         Height          =   1995
         Left            =   105
         OleObjectBlob   =   "IMPORCUECLI.frx":0010
         TabIndex        =   23
         Top             =   210
         Width           =   11295
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   " IMPORTACION DEFINITIVA"
      Height          =   330
      Left            =   8190
      TabIndex        =   17
      Top             =   525
      Width           =   2640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PRE- IMPORTACION"
      Height          =   330
      Left            =   8190
      TabIndex        =   16
      Top             =   105
      Width           =   2640
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Planilla Origen/Detalle Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   11625
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   "F:\FLGUIDI\TRAIDOS\migracion nucleo\Clientes\REC.xls"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "REC$"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPORCUECLI.frx":09C2
         Height          =   1755
         Left            =   105
         OleObjectBlob   =   "IMPORCUECLI.frx":09D2
         TabIndex        =   8
         Top             =   210
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11025
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   1050
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   6
         Top             =   105
         Width           =   12000
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\FLGUIDI\TRAIDOS\migracion nucleo\Clientes\CASI.XLS"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   5565
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "CASI$"
      Top             =   840
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "IMPORCUECLI.frx":1384
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Planilla Origen/Encabezado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   11625
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   "F:\FLGUIDI\TRAIDOS\migracion nucleo\Clientes\CCOM.XLS"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Left            =   4305
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   0  'Table
         RecordSource    =   "CCOM$"
         Top             =   735
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "IMPORCUECLI.frx":14CE
         Height          =   1050
         Left            =   120
         OleObjectBlob   =   "IMPORCUECLI.frx":14DE
         TabIndex        =   15
         Top             =   210
         Width           =   11295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Width           =   7980
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   4815
         TabIndex        =   10
         Top             =   630
         Width           =   4875
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "CASI.xls"
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
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   945
         TabIndex        =   2
         Text            =   " "
         Top             =   945
         Width           =   4890
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F:\FLGUIDI\TRAIDOS\migracion nucleo\Clientes\"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         TabIndex        =   13
         Top             =   315
         Width           =   6840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   12
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   945
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   3
         Top             =   630
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      Caption         =   "DE"
      Height          =   225
      Left            =   9765
      TabIndex        =   21
      Top             =   945
      Width           =   330
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   225
      Left            =   9135
      TabIndex        =   20
      Top             =   945
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   225
      Left            =   10185
      TabIndex        =   19
      Top             =   945
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   225
      Left            =   8190
      TabIndex        =   18
      Top             =   945
      Width           =   645
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   105
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORCUECLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
End Sub

Private Sub Combo1_Click()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Recordset.MoveFirst
   Data1.Refresh
   '
End Sub

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'


Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.filename)
   Data1.DatabaseName = Label1
   Combo1.Clear
   Combo1.TEXT = ""
   On Error Resume Next
   Data1.Refresh
   '
   CATAB% = Data1.Database.TableDefs.Count
   For TTB% = 1 To CATAB%
     Combo1.AddItem Data1.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo1.TEXT = Data1.Database.TableDefs(TTB% - 1).Name
     End If
   Next TTB%
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Dim FECH As Date
   '
   Call ABRECABAN
   GoTo 12 ' PARA EVITAR BORRAR CAJABANC
   On Error GoTo 0
   SQLX$ = " SELECT ENCMOVCC.*, DETMOVCC.* FROM ENCMOVCC "
   SQLX$ = SQLX$ + " INNER JOIN DETMOVCC "
   SQLX$ = SQLX$ + " ON ENCMOVCC.N_Mov = DETMOVCC.N_Mov "
   SQLX$ = SQLX$ + " WHERE ENCMOVCC.N_Mov = DETMOVCC.N_Mov "
   SQLX$ = SQLX$ + " ORDER BY ENCMOVCC.N_Mov "
   Set CUECUICLITMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   With CUECUICLITMP
    On Error Resume Next
    .MoveLast
    XCONTOT& = .RecordCount
    Label6 = XCONTOT&
    Label6.Refresh
    .MoveFirst
    XCONT& = 0
    If Err Then
      Call MENSERR(24, "Imposible Completar la Importación Definitiva")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    Call ABRECABAN
    '/////////////***HASTA ACA OK***/////////////////////
    On Error Resume Next
      CajaBanc.MoveLast
    While CajaBanc.RecordCount > 0  'limpiar la tabla
      CajaBanc.MoveLast
      CajaBanc.Delete
    Wend
   '
    VEDEPOS$ = ""
    Do While .EOF = False
      '
      Select Case (!COMPRO)
        'AJUSTES DE SALDO
        Case "A-CTA"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "COMIN"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "DEART"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "GEREC"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "INICI"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "COMPE"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
        Case "V/PAG"
           TICOMX$ = "AS": OPMOVIX$ = "AJUP"
         'FACTURAS
        Case "FACNP"
           TICOMX$ = "FC": OPMOVIX$ = "FAC"
        Case "FACMO"
           TICOMX$ = "FC": OPMOVIX$ = "FAC"
        Case "FASNP"
           TICOMX$ = "FC": OPMOVIX$ = "FAC"
         'NOTAS DE DEBITO
        Case "NDBCO"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
        Case "NDCLS"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
        Case "NDCLI"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
        Case "NODEB"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
          'NOTAS DE CREDITO
        Case "NCREG"
           TICOMX$ = "NC": OPMOVIX$ = "CRE"
        Case "NCCLI"
           TICOMX$ = "NC": OPMOVIX$ = "CRE"
      End Select
      If TICOMX$ = "FC" Then
        TIPFAC$ = "A"
        If (!T_Comp) = "FVCA" Then TIPFAC$ = "A"
        If (!T_Comp) = "FVCB" Then TIPFAC$ = "B"
      Else
        TIPFAC$ = TICOMX$  'SI ES NC O ND  ES ND O NC
      End If
      
      '
      NFCZ$ = TRIM$(!N_Comp): While Len(NFCZ$) < 8: NFCZ$ = "0" + NFCZ$: Wend
      NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + TIPFAC$
      NFCY$ = TICOMX$ + "-" + "0001" + "-" + NFCZ$ + "-" + TIPFAC$
      '
      SIGX% = 1: If (!COMPRO) = "NCCLI" Then SIGX% = (-1)
      NC% = !N_Cli
      FFI% = CVINT(!FEC_MOV)
      '
      CoComCor$ = NFCX$
      CoComLar$ = NFCY$
      CodiEmp% = 0
      CodiSuc% = 0
      NumeClie% = NC%
      NumeProv% = 0
      FechEmi = CVDAT(FFI%)
      Referen$ = RASOCLI$(NC%)
      FeContab = CVDAT(FFI%)
      FeAcredi = CVDAT(FFI%)
      CuentaCo% = !N_CueImpu 'CUECOCLI%(NC%)
      Select Case !N_CueImpu
       Case 38
           TiMovi$ = "FVEN"
           OpMovi$ = TICOMX$
           VaMovi$ = TIPFAC$
         If TICOMX$ <> "NC" Then
           ImpoDeb = !IMPORMOVI * SIGX%
         Else
           impohabe = !IMPORMOVI * SIGX% * (-1)
         End If
         
        Case 163
           TiMovi$ = "IVA"
           OpMovi$ = "DeFi"
           Descrip$ = "IVA DEBITO FISCAL"
           VaMovi$ = "21%"
          If TICOMX$ <> "NC" Then
              ImpoDeb = !IMPORMOVI * SIGX%
          Else
              ImpoDeb = !IMPORMOVI * SIGX% * (-1)
          End If
        Case Else
           TiMovi$ = "VENTA"
           OpMovi$ = TICOMX$
           VaMovi$ = TIPFAC$
           Descrip$ = "Mercaderia Vendida"
          If TICOMX$ <> "NC" Then
              ImpoDeb = !IMPORMOVI * SIGX% * (-1)
          Else
              ImpoDeb = !IMPORMOVI * SIGX%
          End If
      End Select
      Descrip$ = TIPFAC$ + " " + NFCY$
      NumeTal = Val("0001")
      NumeCom = CDbl(!N_Comp)
      StatInit = 0
      
      Observa$ = RASOCLI$(NC%)
      Call GRACABAN
      '
      XCONT& = XCONT& + 1
      Label7 = XCONT&
      Label7.Refresh
      .MoveNext
    Loop
   End With
   '
12    TEPRUE = ""
   Data3.DatabaseName = Label9 & "REC.XLS"
   Data3.RecordSource = "REC$"
   Data3.Refresh
   On Error Resume Next
   URECORD& = Data3.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   '
   DBHANT = DBGrid3.Height
   DBGrid3.RowHeight = 50
   DBGrid3.Height = (URECORD& + 100) * DBGrid3.RowHeight * 1.1
   '
   URECORD& = Data3.Recordset.RecordCount
   Label6 = URECORD&
   Label6.Refresh
   Label5 = "DET-REC"
   '
   With Data3
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
       '
       Label7 = U&
       Label7.Refresh
       '
       NUMOVDET& = Val(.Recordset.Fields(1))  'NUME MOVI CTA. CTE.
       NOMTIPCOM$ = TRIM$(.Recordset.Fields(3)) 'NOMBRE COMPROBANTE
       NUMCOMP$ = Val(.Recordset.Fields(6)) 'NUMERO COMRPOBANTE
       FECMOV = CDate(TRIM$(.Recordset.Fields(7)))  'FECHA MOVI
       ImporMov# = .Recordset.Fields(16)  ' IMPORTE DEL MOVI
       If IsNull(.Recordset.Fields(17)) Then
          MONEDA% = 0
       Else
          MONEDA% = Val(.Recordset.Fields(17))  'MONEDA 0=PESOS,1=U$S
       End If
       If IsNull(.Recordset.Fields(18)) Then
          PARIMONE# = 1
       Else
          PARIMONE# = Val(.Recordset.Fields(18))  'PARIDAD DE LA MONEDA
       End If
       If IsNull(.Recordset.Fields(9)) Then   'NUMERO DE CLIENTE
          NUCLIE$ = "0"
       Else
          NUCLIE$ = TRIM$(.Recordset.Fields(9))
       End If
       NUCLI% = NUMECLI%(TRIM$(NUCLIE$))
       '/////*****HASTA ACA LLEGUE*****///////
       'FALTA ASIGNAR A LAS VARIABLES GLOBALES LOS DATOS _
       'NECESARIOS PARA EL SUB GRACABAN
'          NUMOVDET&
'          NOMTIPCOM$
'          NUMCOMP$
'          NUCLI%
'          FECMOV
'          ImporMov#
'          MONEDA%
'          PARIMONE#
        Call GRACABAN
       .Recordset.MoveNext
       U& = U& + 1
     Loop
   End With
   '
   Command3.Enabled = True
99 End Sub



Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Call IMPORCUECLI
   MsgBox "Fin de Pre-Importación"
99 Screen.MousePointer = 1
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub

Private Sub Form_Load()

   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   '
   Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call FINAL("")
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub Option1_Click(Index As Integer)
   Label1 = ""
   Combo1.Clear
   DBGrid1.ClearFields
   NARIM = TRIM$(Str$(Index))
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Sub IMPORCUECLI()
   '////****rutina de borrar TABLAS ENCACUECLI-DETACUECLI****/////
    Call ABRECABAN
    GoTo 12 'PARA SALTEAR LA LIMPIEZA DE DB'S Y PRE-IMPORTACION
    On Error Resume Next
      EncMovCc.MoveLast
    While EncMovCc.RecordCount > 0  'limpiar la base de datos ENCABEZADO
      EncMovCc.MoveLast
      EncMovCc.Delete
    Wend
      On Error Resume Next
      DetMovCc.MoveLast
    While DetMovCc.RecordCount > 0  'limpiar la base de datos DETALLES
      DetMovCc.MoveLast
      DetMovCc.Delete
    Wend
   TEPRUE = ""
   Data1.DatabaseName = Label9 + Label1
   Data1.RecordSource = "CASI$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   '
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   '
   Call ABRECABAN
   '
   URECORD& = Data1.Recordset.RecordCount
   Label6 = URECORD&
   Label6.Refresh
   Label5 = "DETA"
   '
   With Data1
   .Recordset.MoveFirst
   U& = 0
   Do While Not .Recordset.EOF
     '
     Label7 = U&
     Label7.Refresh
     '
     NUMOVDET& = Val(.Recordset.Fields(1))  'NUME MOVI CTA. CTE.
     
     NUMORITE& = Val(.Recordset.Fields(2))  'NUME ORDEN ITEM
     
     FECMOV = CDate(TRIM$(.Recordset.Fields(4)))  'FECHA MOVI
     
     If IsNull(.Recordset.Fields(5)) Then
        CUEIMPU% = CUEINT("")
     Else
        CUEIMPU% = CUEINT(TRIM$(.Recordset.Fields(5)))  ' CTA. IMPU CONTABLE
     End If
    
     ImporMov# = .Recordset.Fields(7)  ' IMPORTE DEL MOVI
     
     If IsNull(.Recordset.Fields(8)) Then
        MONEDA% = 0
     Else
        MONEDA% = Val(.Recordset.Fields(8))  'MONEDA 0=PESOS,1=U$S
     End If
     If IsNull(.Recordset.Fields(9)) Then
        PARIMONE# = 1
     Else
        PARIMONE# = Val(.Recordset.Fields(9))  'PARIDAD DE LA MONEDA
     End If
     
     
     '
     With DetMovCc
       .AddNew
       !N_Mov = NUMOVDET&
       !N_OrIt = NUMORITE&
       !FEC_MOV = FECMOV
       !N_CueImpu = CUEIMPU%
       !IMPORMOVI = ImporMov#
       !MoneMov = MONEDA%
       !ParMonMov = PARIMONE#
       .Update
     End With
     .Recordset.MoveNext
     U& = U& + 1
    Loop
   End With
   '
   TEPRUE = ""
   Data2.DatabaseName = Label9 & "CCOM.XLS"
   Data2.RecordSource = "CCOM$"
   Data2.Refresh
   On Error Resume Next
   URECORD& = Data2.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   '
   DBHANT = DBGrid2.Height
   DBGrid2.RowHeight = 50
   DBGrid2.Height = (URECORD& + 100) * DBGrid2.RowHeight * 1.1
   '
   Call ABRECABAN
   '
   URECORD& = Data2.Recordset.RecordCount
   Label6 = URECORD&
   Label6.Refresh
   Label5 = "ENCA"
   Label5.Refresh
   '
   U& = 0
   With Data2
     Do While Not .Recordset.EOF
     '
     Label7 = U&
     Label7.Refresh
     '
     NUMOVENC& = Val(.Recordset.Fields(1))  'NUME MOVI CTA. CTE.
     
     COMPROb$ = TRIM$(.Recordset.Fields(2))  'COMPROBANTE
     
     TIPCOMP$ = TRIM$(.Recordset.Fields(3)) 'TIPO COMPROBANTE
     
     NUCOMPRO& = Val(.Recordset.Fields(6))  'Nº COMPROBANTE
     
     CODCLI% = NUMECLI%(TRIM$(.Recordset.Fields(9)))  ' COD. ENTERO DEL CLIENTE
     '
     With EncMovCc
       .AddNew
       !N_Mov = NUMOVENC&
       !COMPRO = COMPROb$
       !T_Comp = TIPCOMP$
       !N_Comp = NUCOMPRO&
       !N_Cli = CODCLI%
       .Update
     End With
       U& = U& + 1
       .Recordset.MoveNext
     Loop
12  End With
    '
99 End Sub
Sub IMPORMOVSTK()
   '
   Screen.MousePointer = 11
   Dim FECHAMOVI As Date
   '
   Call ABRESTOCKS
   On Error Resume Next
      Movisto.MoveLast
    While Movisto.RecordCount > 0  'limpiar la base de datos MOVISTO
      Movisto.MoveLast
      Movisto.Delete
    Wend
   On Error GoTo 0
   '
   For AX% = 93 To 104
    Call ABREMOVSTK
    '
    On Error Resume Next
      DetaStock.MoveLast
    While DetaStock.RecordCount > 0  'limpiar la base de datos
      DetaStock.MoveLast
      DetaStock.Delete
    Wend
    '
    On Error Resume Next
     EncaStock.MoveLast
    While EncaStock.RecordCount > 0 'limpiar la base de datos
      EncaStock.MoveLast
      EncaStock.Delete
    Wend
    '
    On Error GoTo 0
    '
    Label5 = AX%
    Label5.Refresh
    '
    TEPRUE = ""
    '
    ANX$ = Right$(TRIM$(Str$(AX%)), 2)
    arc1$ = "MOV-" + ANX$
    arc2$ = "MOVO-" + ANX$
    '
    Select Case AX%
        Case 93
         noresou1$ = "'MOV-93'$"
         noresou2$ = "'MOVO-93'$"
        Case 94
         noresou1$ = "'MOV-93'$"
         noresou2$ = "'movo-94'$"
        Case 95
         noresou1$ = "'MOV 95'$"
         noresou2$ = "MOV95$"
        Case 96
         noresou1$ = "'MOV 95'$"
         noresou2$ = "MOVO96$"
        Case 97
         noresou1$ = "'MOV 97'$"
         noresou2$ = "MOV97$"
        Case 98
         noresou1$ = "'mov 98'$"
         noresou2$ = "'movo-98'$"
        Case 99
         noresou1$ = "'mov 99'$"
         noresou2$ = "'movo-99'$"
        Case 100
         noresou1$ = "'mov 2000'$"
         noresou2$ = "'movo-00'$"
        Case 101
         noresou1$ = "'mov 2001'$"
         noresou2$ = "'movo-01'$"
        Case 102
         noresou1$ = "'mov 2002'$"
         noresou2$ = "'movo-02'$"
        Case 103
         noresou1$ = "'mov 2003'$"
         noresou2$ = "'movo-03'$"
        Case 104
         noresou1$ = "'mov 2004'$"
         noresou2$ = "'movo-04'$"
    End Select
    '
    Data2.DatabaseName = Label9 + arc1$
    Data2.RecordSource = noresou1$
    Data2.Refresh
    On Error Resume Next
    URECORD& = Data2.Recordset.RecordCount
    If Err Then
      Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
      GoTo 99
    End If
   
    DBHANT = DBGrid2.Height
    DBGrid2.RowHeight = 10
    DBGrid2.Height = (URECORD& + 100) * DBGrid2.RowHeight * 1.1

    Call ABREMOVSTK
    On Error GoTo 0
    '
    URECORD& = Data2.Recordset.RecordCount
    Label6 = URECORD&
    Label6.Refresh
    '
    With Data2
      conti& = 1
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        CODCOM$ = ""
        If IsNull(.Recordset.Fields(1)) = False Then
          CODCOM$ = TRIM$(.Recordset.Fields(1))   'CODIGO COMPROBANT
        End If
        NUOC$ = ""
        If .Recordset.Fields(2) <> "" Then
         NUOC$ = Right$(.Recordset.Fields(2), 7)
         NUMCOM& = Val(NUOC$) 'NUMERO COMPROBANTE
        End If
        TIPOMOVI$ = ""
        If IsNull(.Recordset.Fields(4)) = False Then
          TIPOMOVI$ = TRIM$(.Recordset.Fields(4)) 'TIPO DE MOVIMIENTO
        End If
        CODART$ = ""
        If .Recordset.Fields(6) <> "" Then
         CODART$ = TRIM$(.Recordset.Fields(6)) 'CODIGO ART.
        End If
        CODDEPO$ = ""
        If IsNull(.Recordset.Fields(7)) = False Then
        CODDEPO$ = TRIM$(.Recordset.Fields(7)) 'CODIGO DEPOSITO
        End If
        DESCRIPC$ = ""
        If .Recordset.Fields(10) <> "" Then
         DESCRIPC$ = TRIM(.Recordset.Fields(10))  'DESCRIPCION
        End If
        FECHAMOVI = 1 / 1 / 80
        If IsNull(.Recordset.Fields(13)) = False Then
          FECHAMOVI = Format(TRIM$(.Recordset.Fields(13)), "DD/MM/YYYY") 'FECHA DE MOVIMIENTO
        End If
        NUMTRAN& = 0
        If IsNull(.Recordset.Fields(17)) = False Then
        NUMTRAN& = Val(.Recordset.Fields(17)) 'NUMERO DE TRANSACCION
        End If
'     Call TRACE(20, U&, URECORD&)
'     TTXX$ = Space$(254)
      '
        On Error GoTo 0
        With EncaStock
          .AddNew
          !C_Com = CODCOM$
          !N_Com = NUMCOM&
          !C_Tmv = TIPOMOVI$
          !FechaMov = FECHAMOVI
          !C_Ente = CODART$
          CODDEPO1$ = Right$(CODDEPO$, 2)
          !C_Dep = CODDEPO1$
          !N_Trn = NUMTRAN&
          !Descrip = DESCRIPC$
          .Update
        End With
        .Recordset.MoveNext
        conti& = conti& + 1
        Label7 = conti&
        Label7.Refresh
      Loop
    End With
    '
30  TEPRUE = ""
    Data1.DatabaseName = Label9 + arc2$
    Data1.RecordSource = noresou2$
    Data1.Refresh
    On Error Resume Next
    URECORD& = Data1.Recordset.RecordCount
    If Err Then
      Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
      GoTo 99
    End If
    '
    DBHANT = DBGrid1.Height
    DBGrid1.RowHeight = 10
    DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
    '
    Call ABREMOVSTK
    On Error GoTo 0
    '
    URECORD& = Data1.Recordset.RecordCount
    Label6 = URECORD&
    Label6.Refresh
    conti& = 1
    Data1.Recordset.MoveFirst
    '
    Do While Data1.Recordset.EOF = False
      CODCOM$ = ""
      If IsNull(Data1.Recordset.Fields(1)) = False Then
        CODCOM$ = TRIM$(Data1.Recordset.Fields(1))
      End If
      NUMCOM& = 0
      If Data1.Recordset.Fields(2) <> "" Then
        NUMCOM& = Val(Right$(Data1.Recordset.Fields(2), 7))
      End If
      TIPOMOVI$ = ""
      If IsNull(Data1.Recordset.Fields(4)) = False Then
        TIPOMOVI$ = TRIM$(Data1.Recordset.Fields(4))
      End If
      FECHAMOVI = 1 / 1 / 80
      If Data1.Recordset.Fields(5) <> "" Then
        FECHAMOVI = Format(TRIM$(Data1.Recordset.Fields(5)), "DD/MM/YYYY")
      End If
      SIGNO$ = ""
      If IsNull(Data1.Recordset.Fields(6)) = False Then
        SIGNO$ = TRIM$(Data1.Recordset.Fields(6))
      End If
      CODDEPO$ = ""
      If IsNull(Data1.Recordset.Fields(7)) = False Then
        CODDEPO$ = TRIM$(Data1.Recordset.Fields(7))
      End If
      CODART$ = ""
      If IsNull(Data1.Recordset.Fields(9)) = False Then
       CODART$ = TRIM$(Data1.Recordset.Fields(9))
      End If
      CANTID# = 0
      If IsNull(Data1.Recordset.Fields(10)) = False Then
       CANTID# = Val(Data1.Recordset.Fields(10))
      End If
      NUMTRANS& = 0
      NUCO$ = ""
      If IsNull(Data1.Recordset.Fields(14)) = False Then
        NUCO$ = Right$(Data1.Recordset.Fields(14), 7)
        If Mid$(NUCO$, 3, 2) = "E+" Then
        NUMTRANS& = 0
        Else
        NUMTRANS& = Val(NUCO$)
        End If
      End If
      '
      With DetaStock
        .AddNew
        !C_Com = CODCOM$
        !N_Com = NUMCOM&
        !C_Tmv = TIPOMOVI$
        !FECHA_mOV = FECHAMOVI
        !C_Sig = SIGNO$
        CODART1$ = Right$(CODART$, 16)
        !C_Art = CODART1$
        CODDEPO1$ = Left$(CODDEPO$, 2)
        !C_Dep = CODDEPO1$
        !Canti = CANTID#
        !N_Trn = NUMTRANS&
        .Update
      End With
      '
      U& = U& + 1            'PARA TRAZADOR
      '
      conti& = conti& + 1
      Label7 = conti&
      Label7.Refresh
    Data1.Recordset.MoveNext
    Loop
   '
   Call ABREMOVSTK
   On Error GoTo 0
   SQLX$ = " SELECT DETASTOCK.*, ENCASTOCK.* FROM DETASTOCK "
   SQLX$ = SQLX$ + " INNER JOIN ENCASTOCK "
   SQLX$ = SQLX$ + " ON DETASTOCK.N_Trn = ENCASTOCK.N_Trn"
   SQLX$ = SQLX$ + " ORDER BY Fecha_Mov "
   Set STOTEMP = STOCKORIGI.OpenRecordset(SQLX$, dbOpenSnapshot)
   With STOTEMP
    On Error Resume Next
    .MoveLast
    XCONTOT& = .RecordCount
    Label6 = XCONTOT&
    Label6.Refresh
    .MoveFirst
    XCONT& = 0
    If Err Then
      Call MENSERR(24, "Imposible Completar la Importación Definitiva")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
     '
    VEDEPOS$ = ""
    Label5 = Label5 & "GRAB "
    Label5.Refresh
    Do While .EOF = False
      '
      With Movisto
        .AddNew
          !FechMov = STOTEMP.Fields("FECHA_mOV")
          !Cod_Inte = CODINT%(STOTEMP.Fields("C_Art"))
          !Cod_Exte = STOTEMP.Fields("C_Art")
          !Descrip = DESCRIT0$(!Cod_Inte)
          !IdenLote = ""
          !CodiMovi = Left(STOTEMP.Fields("ENCASTOCK.C_Tmv"), 2)
          COMPCH$ = Left(STOTEMP.Fields("ENCASTOCK.C_Com"), 2)
          DOCU$ = COMPCH$ & "-" & CStr(STOTEMP.Fields("ENCASTOCK.N_Com"))
          !DoPriCor = DOCU$
          !DoPriLar = STOTEMP.Fields("ENCASTOCK.C_Com") & "-" & CStr(STOTEMP.Fields("ENCASTOCK.N_Com"))
          !DoSecCor = DOCU$
          !DoSecLar = STOTEMP.Fields("ENCASTOCK.C_Com") & "-" & CStr(STOTEMP.Fields("ENCASTOCK.N_Com"))
          !NuorItem = 0
          !ReferCor = STOTEMP.Fields("Descrip")
          !ValoUnit = 0
          !MoneVal = 0
            DEPX$ = "@@" + AJUSTI$(STOTEMP.Fields("ENCASTOCK.C_Dep"), 2)
            PPXX% = InStr(VEDEPOS$, DEPX$)
            If PPXX% < 1 Then
              PPXX% = 1 + Len(VEDEPOS$)
              VEDEPOS$ = VEDEPOS$ + DEPX$
            End If
          !DepoMovi = (PPXX% + 3) \ 4
          !Nume_Clie = 0
          !Nume_Prov = 0
          If STOTEMP.Fields("C_Sig") = "+" Then
            !CantIngre = STOTEMP.Fields("Canti")
          Else
            !CantSalid = STOTEMP.Fields("Canti")
          End If
          !CantSaldo = 0
          !FeReMovi = Now
          !NumUsuar = 0
        .Update
      End With
      '
      XCONT& = XCONT& + 1
      Label7 = XCONT&
      Label7.Refresh
      .MoveNext
    Loop
   End With
    '
25 Next AX%
   '
   EncaStock.Close
   DetaStock.Close
   '
20 Screen.MousePointer = 1
   '
99 End Sub

Sub IMPORMAS2()
   '
5  Label1 = Label9 + "MAESTRO PRODUCTOS\STO.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "STO$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   '
   Call ABREMASTER
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
10   CDXI% = CODINT%(CODEX$)
     If CDXI% = 0 Then
       Call MENSERR(24, "Código / Campo Clave '" + CODEX$ + "' Inexistente en Base de Datos")
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     DBGrid1.Col = 3
     STOMI = Val(DBGrid1.TEXT)
     '
     DBGrid1.Col = 4
     PUNPE = Val(DBGrid1.TEXT)
     If STOMI < 1 Then STOMI = PUNPE
     '
     DBGrid1.Col = 5
     STOMA = Val(DBGrid1.TEXT)
     If STOMA < STOMI Then STOMA = STOMI
     '
     LOREP = STOMA - STOMI
     If LOREP > 10000 Then LOREP = 10000
     If LOREP < STOMI Then LOREP = STOMI
     '
     With Master
       .Edit
       !STOMIN = STOMI
       !STOMAX = STOMA
       !lorepos = LOREP
       .Update
     End With
19 Next U&
   '
99 End Sub
'
Sub AGRECODBAJ(CODEX$)
    If TRIM$(CODEX$) <> "" Then
      With Master
       .AddNew
       !CODIGO = CODEX$
       !DESCRIPCION = "Codigo " + CODEX$
       !STATUS = (-1)
         NMS% = NUMAS% + 1
       !CODINT = NMS%
       .Update
     End With
   End If
End Sub

Sub IMPORSTRUC()
   '
   Call ABRESTRUC
   On Error Resume Next
   Formulas.MoveLast
   While Formulas.RecordCount > 0
     Formulas.MoveLast
     Formulas.Delete
   Wend
   '
6  Label1 = Label9 + "ESTRUCTURAS\EST.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 20
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "EST$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 20
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   Call ABREMASTER
   Call ABRESTRUC
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then GoTo 19
     '
10   If CODINT%(CODEX$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     DBGrid1.Col = 1
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Then GoTo 19
     If COELE$ = CODEX$ Then GoTo 19
     '
     If Len(COELE$) = 3 Then
       If Left$(COELE$, 2) = "MO" Then
         GoTo 19   ' MANO DE OBRA
       End If
     End If
     '
12   If CODINT%(COELE$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento Inexistente o No Valido"
       Call AGRECODBAJ(COELE$)
       GoTo 12
     End If
     '
     CICON% = CODINT%(CODEX$)
     CIELE% = CODINT%(COELE$)
     '
     DBGrid1.Col = 2
     USONE = CDbl(DBGrid1.TEXT)
     '
     DBGrid1.Col = 3
     UMEDEL$ = TRIM$(DBGrid1.TEXT)
     '
     If UMEDEL$ <> UNIMED$(CIELE%) Then
       With Master
         .Edit
         !UMEDIDA = UMEDEL$
         .Update
       End With
     End If
     '
     With Formulas
       .AddNew
       !CODICONJ = Abs(CICON%)
       !CODXCONJ = CODEX$
       !CODIELEM = Abs(CIELE%)
       !CODXELEM = COELE$
       !USOBRUTO = USONE
       !PLADISPO = 0
       .Update
     End With
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub IMPORRUTAS()
   '
   Dim UOP%(32767)
   '
   Call ABRESECOP
   On Error Resume Next
   Secoper.MoveLast
   While Secoper.RecordCount > 0
     Secoper.MoveLast
     Secoper.Delete
   Wend
   '
6  Label1 = Label9 + "ESTRUCTURAS\EST.XLS"
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 20
   End If
   '
   TEPRUE = ""
   Data1.DatabaseName = Label1
   Data1.RecordSource = "EST$"
   Data1.Refresh
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 20
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.RowHeight = 50
   DBGrid1.Height = (URECORD& + 100) * DBGrid1.RowHeight * 1.1
   Call ABREMASTER
   Call ABRESECOP
   '
   If DBGrid1.Columns.Count < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 1
     COELE$ = UCase$(TRIM$(DBGrid1.TEXT))
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Then GoTo 19
     If COELE$ = CODEX$ Then GoTo 19
     '
     If Len(COELE$) <> 3 Then GoTo 19
     '
     If UCase$(Left$(COELE$, 2)) <> "MO" Then
       GoTo 19   ' NO ES MANO DE OBRA
     End If
     '
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then GoTo 19
     '
10   If CODINT%(CODEX$) = 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       Call AGRECODBAJ(CODEX$)
       GoTo 10
     End If
     '
     CICON% = CODINT%(CODEX$)
     '
     DBGrid1.Col = 2
     USONE = CDbl(DBGrid1.TEXT)
     '
     DBGrid1.Col = 3
     UMEDEL$ = UCase$(TRIM$(DBGrid1.TEXT))
     If Left$(UMEDEL$, 3) <> "MIN" Then
        AAA = BBB
        MsgBox "Unidad de M.O. no Reconocida"
     End If
     '
     If USONE * LOTESTAN(CICON%) / 60 > 999 Then
       ' limita tamaño del lote a menos de 999 horas
       Master.Edit
       Master!lorepos = 1
       Master.Update
     End If
     '
     With Secoper
       .AddNew
       !CODICONJ = Abs(CICON%)
       !CODXCONJ = CODEX$
         UOP%(CICON%) = UOP%(CICON%) + 10
       !NUMEOPER = UOP%(CICON%)
         If COELE$ = "MOA" Then
             DESOP$ = "ARMADO"
           ElseIf COELE$ = "MOE" Then
             DESOP$ = "ESTAMPADO"
           ElseIf COELE$ = "MOP" Then
             DESOP$ = "PINTURA"
           Else
             AAA = BBB
         End If
       !DescOper = DESOP
       !codmaq0 = COELE$
       !STATOPER = 1
       !MinuStan = USONE
       !LOTESTAN = LOTESTAN(CICON%)
       If !LOTESTAN > 10000 Then
         !LOTESTAN = 10000
       End If
       !CantiOps = 1
       If USONE > 0.05 Then
         !Cadencia = 60 / USONE
       End If
       '
       ' PARA INICIALIZAR VALORES
       !HorsPrep = 0
       !HorsFijo = 0
       !HorsLimp = 0
       !HorsMHer = 0
       !PieCiclo = 0
       !PorDemor = 0
       !CoefSolp = 0
       !CategOps = 0
       .Update
     End With
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORUTS()
   '
   If DBGrid1.Columns.Count < 12 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     If CODEX$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = CODEX$
     '
     DBGrid1.Col = 1
     DEOPE$ = TRIM$(DBGrid1.TEXT)
     If DEOPE$ = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     NOPE% = NOPE% + 10
     VALO$ = FORMATNUM$(NOPE%, "I4")
     Call REPLA(TTXX$, VALO$, 17, 4)
     Call REPLA(TTXX$, DEOPE$, 23, 48)
     '
     DBGrid1.Col = 2                     ' MAQUINA
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 73, 6)
     '
     DBGrid1.Col = 3                     ' HORAS PREPARACION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 79, 10)
     '
     DBGrid1.Col = 4                     ' HORAS STANDARD FIJO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 89, 10)
     '
     DBGrid1.Col = 5                     ' MINUTOS STANDARD VARIABLE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.5")
     Call REPLA(TTXX$, VALO$, 99, 12)
     '
     DBGrid1.Col = 6                     ' LOTE STANDARD UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 1 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F10.1")
     Call REPLA(TTXX$, VALO$, 111, 10)
     '
     DBGrid1.Col = 7                     ' % DEM.DISTRIBUIDA
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     If VALOR < 0 Then VALOR = 0
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F8.1")
     Call REPLA(TTXX$, VALO$, 121, 8)
     '
     DBGrid1.Col = 8                     ' OPERARIOS
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0.01 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F8.2")
     Call REPLA(TTXX$, VALO$, 129, 8)
     '
     DBGrid1.Col = 9                     ' MAQUINA AUX 1
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 139, 6)
     '
     DBGrid1.Col = 10                     ' MAQUINA AUX 2
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 147, 6)
     '
     DBGrid1.Col = 11                     ' MAQUINA AUX 3
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 155, 6)
     '
     DBGrid1.Col = 12                     ' H.P.FABR.
     HOPROFA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, HOPROFA$, 163, 8)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
       TEPRUE = TRIM$(TTXX$) + "*"
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 9, 30)
     '
     DBGrid1.Col = 2                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 999999.99 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 41, 12)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(DBGrid1.TEXT)
     If VTX$ <> "" And Val(VTX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(VTX$)
     On Error GoTo 0
     If VALOR > 999.999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 53, 10)
     '
     DBGrid1.Col = 4
     CODEQ$ = TRIM$(DBGrid1.TEXT)
     If CODEQ$ <> "" Then
       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CODEQ$, 66, 6)
     '
     DBGrid1.Col = 5                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
     If CODEQ$ = "" Then VALOR = 0
     If VALOR < VAMIN Or VALOR > 999.9 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     Call REPLA(TTXX$, VALO$, 72, 12)
     '
     DBGrid1.Col = 6
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
       TEPRUE = TRIM$(TTXX$) + "*"
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   
20 End Sub

Sub PREIMPORCAT()
   '
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   If URECORD& > 255 Then
     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
     GoTo 20
   End If
   '
   If ULTREG&("CATOPER") > 0 Then
     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
       GoTo 20
     End If
   End If
   '
   JJ& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     DECAT$ = TRIM$(DBGrid1.TEXT)
     If DECAT$ <> "" Then
       JJ& = JJ& + 1
       COCAT$ = AJUSTD$(Str$(JJ&), 4)
       Call REPLA(TTXX$, COCAT$, 1, 4)
       Call REPLA(TTXX$, DECAT$, 7, 30)
       '
       DBGrid1.Col = 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       If VALOR < 0 Or VALOR > 999999.99 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
         List1.AddItem MERROX$
         GoTo 19
       End If
       VALO$ = FORMATNUM$(VALOR, "F12.4")
       Call REPLA(TTXX$, VALO$, 37, 12)
       '
       DBGrid1.Col = 2
       OBSER$ = TRIM$(DBGrid1.TEXT)
       Call REPLA(TTXX$, OBSER$, 51, 32)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
         TEPRUE = TRIM$(TTXX$) + "*"
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&
   
20 End Sub


Sub IMPORSTRU()
   '
   Dim PUNSTRU$(32767)
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   ATENTIA% = ATENTI%: ATENTI% = 12345 ' PA QUE NO REORGANICE IMPLOSI
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       CAIT% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = Space$(64)
         Call REPLA(TTYY$, Mid$(TTXX$, 19, 16), 1, 16)
         Call REPLA(TTYY$, Mid$(TTXX$, 69, 12), 49, 12)
         Call REPLA(TTYY$, Mid$(TTXX$, 81, 4), 61, 4)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("ESTRULAR")
   ATENTI% = ATENTIA%
   '
End Sub
'
Sub IMPORUTS()
   '
   Dim PUNSTRU$(32767)
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   REBLA$ = REGBLAN$("SECOPER")
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       RFF$ = RECFILT$("SECOPER", 1, MKI$(CDXI%))
       '
       For UKU& = 1 To Len(RFF$) Step 4
         RESE& = CVS(Mid$(RFF$, UKU&, 4))
         Call GRAREG("SECOPER", String$(128, 0), RESE&)
       Next UKU&
       '
       NOPEI% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = REBLA$
         Call REPLA(TTYY$, MKI$(CODINT%(Left$(TTXX$, 16))), 1, 2)
           NOPEI% = NOPEI% + 10
         Call REPLA(TTYY$, MKI$(NOPEI%), 3, 2)
         Call REPLA(TTYY$, Mid$(TTXX$, 23, 48), 5, 48)
         Call REPLA(TTYY$, Mid$(TTXX$, 73, 6), 53, 6)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 79, 10))), 59, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 89, 10))), 63, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 99, 12))), 67, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 111, 10))), 91, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 121, 8))), 75, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 129, 8))), 83, 4)
         Call REPLA(TTYY$, Mid$(TTXX$, 139, 6), 103, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 147, 6), 109, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 155, 6), 115, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 163), 121, 6)
         Call REGAPP("SECOPER", TTYY$)
       Next UKU&
     End If
   Next CDXI%
   '
   Call CIERRARCH("SECOPER")
   Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAQ()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADMAQ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COMAQ$ = TRIM$(Left$(TTXX$, 6))
     DEMAQ$ = TRIM$(Mid$(TTXX$, 9, 30))
     CUHOR = Val(Mid$(TTXX$, 41, 12))
     SECCI = Val(Mid$(TTXX$, 53, 10))
     MAQEQ$ = TRIM$(Mid$(TTXX$, 66, 6))
     COEFI% = Val(Mid$(TTXX$, 72, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 87, 30))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, COMAQ$, 1, 6)
     Call REPLA(REMAQ$, DEMAQ$, 7, 30)
     Call REPLA(REMAQ$, MKS$(SECCI), 37, 4)
     Call REPLA(REMAQ$, MAQEQ$, 43, 6)
     Call REPLA(REMAQ$, MKI$(COEFI), 49, 2)
     Call REPLA(REMAQ$, OBSER$, 51, 30)
     Call REPLA(REMAQ$, MKS$(CUHOR), 93, 4)
     Call FILTERSETRECORD("PADMAQ", 1, COMAQ$, REMAQ$) ' sobreescribe SI IGUAL CODIGO
     '
   Next U&
   Call CIERRARCH("PADMAQ")
   '
End Sub

Sub IMPORCAT()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("CATOPER")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     ICAT% = Val(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = Val(Mid$(TTXX$, 37, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 51, 32))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, Chr$(ICAT%), 1, 1)
     Call REPLA(REMAQ$, DECAT$, 2, 27)
     Call REPLA(REMAQ$, MKS$(CUHOR), 29, 4)
     Call REPLA(REMAQ$, OBSER$, 33, 32)
     JJ& = JJ& + 1
     Call GRAREG("CATOPER", REMAQ$, JJ&)
   Next U&
   '
   Call SETULTREG("CATOPER", JJ&)
   Call CIERRARCH("CATOPER")
   '
End Sub


