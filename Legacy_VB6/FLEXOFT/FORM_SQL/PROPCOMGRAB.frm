VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PROPCOMGRAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Grabación y/o Edición de Propuestas de Compra Guardadas."
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   17445
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   17445
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
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
      Left            =   14760
      TabIndex        =   5
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   7725
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   17415
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   16980
         TabIndex        =   20
         ToolTipText     =   "Limpiar Casillero de Ordenes de Compra"
         Top             =   1320
         Width           =   300
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
         Height          =   360
         Left            =   16680
         TabIndex        =   19
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox TXTNROPEDIDOS 
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
         Height          =   330
         Left            =   9000
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   17
         Top             =   800
         Width           =   8265
      End
      Begin VB.TextBox TXTOCOMPRAS 
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
         Height          =   330
         Left            =   9000
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   15
         Top             =   1300
         Width           =   7665
      End
      Begin VB.TextBox TxtReferencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9000
         MaxLength       =   64
         TabIndex        =   10
         Top             =   1800
         Width           =   8265
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   1755
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   3096
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   5415
         Left            =   0
         TabIndex        =   12
         Top             =   2280
         Width           =   17340
         _ExtentX        =   30586
         _ExtentY        =   9551
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de Pedidos:"
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
         Index           =   5
         Left            =   9000
         TabIndex        =   18
         Top             =   580
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N° de O.Compras:"
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
         Index           =   4
         Left            =   9000
         TabIndex        =   16
         Top             =   1125
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Propuesta"
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
         Index           =   3
         Left            =   15960
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta"
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
         Left            =   14520
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
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
         Left            =   9000
         TabIndex        =   11
         Top             =   1600
         Width           =   1215
      End
      Begin VB.Label LBLFECHA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   14520
         TabIndex        =   8
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label LBLNUMERO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   15960
         TabIndex        =   7
         Top             =   360
         Width           =   1275
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PROPCOMGRAB.frx":0000
      Top             =   0
   End
   Begin VB.PictureBox WebBrowser1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   16275
      TabIndex        =   6
      Top             =   0
      Width           =   16335
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9720
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label3 
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
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      Begin VB.Menu mnuAbrirPropCompra 
         Caption         =   "Seleccionar Compra Registradas"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBorrarPropuesta 
         Caption         =   "Borrar Propuestas de Compra"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Imprimir"
         Begin VB.Menu printHTML 
            Caption         =   "Formato HTML"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPrintSt 
            Caption         =   "Impresion Standard"
         End
      End
   End
End
Attribute VB_Name = "PROPCOMGRAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim LecturaDatosOriginales As Byte
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%
Public SinEfectoClickEnGrilla%
Public CamposMSF2$
'Public CamposMSF3$
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Command1_Click()
Call TXTOCOMPRAS_DblClick   '
 
End Sub


 
Private Sub Command2_Click()
TXTOCOMPRAS.TEXT = ""
End Sub

 Sub Command3_Click()
'PARA GRABAR DESDE EL OTRO FRM TRAER LOS DATOS AQUI VER SI FUNCIONA CON COPIA MSF
'SE DEBERIA PODER EDITAR DESDE AQUI.
'ANULAR DESDE AQUI, CON STATUS = -1
'LISTA DE SELECCION VIGENTES Y ANULADASSS
'CARGAR LA INFORMACION AL ELEGIR DE LA LISTA DE SELECCION.
'
     Command3.Enabled = False
'     AA = CantRegistros&("ENCA_PROPUCOMP", "", "NOMESS")
'     Call BORRAREGISTROS("ENCA_PROPUCOMP", "CANTIDAD >= 0", REGAF&, "NOMESS")
     NROPOP& = XVALO(LBLNUMERO)
     If NROPOP& < 1 Then
        CONDI$ = "NumPropCompra > 0"
        NROPOP& = 1 + XVALO(VALOBADA("ENCA_PROPUCOMP", "MAX(NumPropCompra)", CONDI$, "NOMESS"))
     Else
        FE% = CVINT(VALOBADA("ENCA_PROPUCOMP", "FeInsert", "NumPropCompra =" & NROPOP&, "NOMESS"))
     End If
     
     '
'     Call CREACAMPO("", "ENCA_PROPUCOMP", "Codint", 3, 0)
'     Call CREACAMPO("", "ENCA_PROPUCOMP", "NumPropCompra", 7, 0)
'     Call CREACAMPO("", "ENCA_PROPUCOMP", "FeInsert", 8, 0)
'     Call CREACAMPO("", "ENCA_PROPUCOMP", "FeUpdate", 8, 0)
'     Call CREACAMPO("", "ENCA_PROPUCOMP", "MARBORRA", 3, 0)
     On Error GoTo ERROR_GUARDAR
     FLEXCONN.BeginTrans
     
     SQLX$ = "UPDATE ENCA_PROPUCOMP SET MARBORRA = 1 WHERE NumPropCompra = " & NROPOP&
     FLEXCONN.Execute (SQLX$)
     'Call ACTUREGISTRO("ENCA_PROPUCOMP", "MARBORRA", "NumPropCompra = " & NROPOP&, 1, REGAF&, "NOMESS")
     
     SQLX$ = "SELECT * FROM ENCA_PROPUCOMP "
     SQLX$ = SQLX$ + " WHERE NumPropCompra = " & NROPOP&
     Set RS0 = New ADODB.Recordset
     RS0.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With RS0
        
        For REG& = 1 To MSF1.Rows - 1
           
           CODIX$ = TRIM$(MSF1.TextMatrix(REG&, 1))
           CI1% = CODINT%(CODIX$)
           Cantidad = XVALO(MSF1.TextMatrix(REG&, 3))
           If CI1% > 0 Then
            .AddNew
            !NumPropCompra = NROPOP&
            !CODINT = CI1%
            !Cantidad = Cantidad
            !FeUpdate = Date
            If FE% < 33 Then !FeInsert = Date
            !MARBORRA = 0
            !REFERENCIA = Left$(TxtReferencia.TEXT, 64)
            .Update
           End If
        Next REG&
    End With
    RS0.Close
    Set RS0 = Nothing
    
    CONDI$ = "NumPropCompra = " & NROPOP&
    CONDI$ = CONDI$ + " AND MARBORRA = 1"
    SQLX$ = "DELETE FROM ENCA_PROPUCOMP WHERE " & CONDI$
    FLEXCONN.Execute (SQLX$)
    'Call BORRAREGISTROS("ENCA_PROPUCOMP", CONDI$, REGAF&, "NOMESS")
     '
    SQLX$ = "UPDATE DET_PROPUCOMP SET MARBORRA = 1 WHERE NumPropCompra = " & NROPOP&
    FLEXCONN.Execute (SQLX$)
    ' Call ACTUREGISTRO("DET_PROPUCOMP", "MARBORRA", "NumPropCompra = " & NROPOP&, 1, REGAF&, "NOMESS")
     
     SQLX$ = "SELECT * FROM DET_PROPUCOMP "
     SQLX$ = SQLX$ + " WHERE NumPropCompra = " & NROPOP&
     Set rs = New ADODB.Recordset
     rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With rs
        For REG& = 1 To MSF3.Rows - 1
           CI1% = XVALO(MSF3.TextMatrix(REG&, 1))
           NECESIDAD = XVALO(MSF3.TextMatrix(REG&, 5))
           stockMinimo = XVALO(MSF3.TextMatrix(REG&, 6))
           loteMaximo = XVALO(MSF3.TextMatrix(REG&, 7))
           Indivisible = XVALO(MSF3.TextMatrix(REG&, 8))
           COMPRASUGERIDA = XVALO(MSF3.TextMatrix(REG&, 9))
           Stock = XVALO(MSF3.TextMatrix(REG&, 10))
           AUTORIZADA = XVALO(MSF3.TextMatrix(REG&, 11))
           NPROV& = XVALO(MSF3.TextMatrix(REG&, 12))
           oCompPendEntrega = XVALO(MSF3.TextMatrix(REG&, 14))
           If CI1% > 0 Then
            .AddNew
            !NumPropCompra = NROPOP&
            !CODINT = CI1%
            !NECESIDAD = NECESIDAD
            !STOMIN = stockMinimo
            !Stomax = loteMaximo
            !Indivisible = Indivisible
            !CompSugerida = COMPRASUGERIDA
            !Stock = Stock
            !CompAuto = AUTORIZADA
            !Num_Prov = NPROV&
            !Status = 0
            !MARBORRA = 0
            !OCPendEnt = oCompPendEntrega

            .Update
           End If
        Next REG&
     End With
     rs.Close
     Set rs = Nothing
     CONDI$ = "NumPropCompra = " & NROPOP&
     CONDI$ = CONDI$ + " AND MARBORRA = 1"
     
     SQLX$ = "DELETE FROM DET_PROPUCOMP WHERE " & CONDI$
     FLEXCONN.Execute (SQLX$)
     'Call BORRAREGISTROS("DET_PROPUCOMP", CONDI$, REGAF&, "NOMESS")
     '
     SQLX$ = "UPDATE PropCompPed SET MARBORRA = 1 WHERE NumPropCompra = " & NROPOP&
     FLEXCONN.Execute (SQLX$)
     SQLX$ = "INSERT INTO PropCompPed"
     SQLX$ = SQLX$ + " (NropedText,fechemis,MARBORRA,NumPropCompra)"
     SQLX$ = SQLX$ + " VALUES('" & Left$(Me.TXTNROPEDIDOS, 256) & "'"
     SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'"
     SQLX$ = SQLX$ + " ,0"
     SQLX$ = SQLX$ + " ," & NROPOP&
     SQLX$ = SQLX$ + " )"
     FLEXCONN.Execute (SQLX$)
     CONDI$ = "NumPropCompra = " & NROPOP&
     CONDI$ = CONDI$ + " AND MARBORRA = 1"
     SQLX$ = "DELETE FROM PropCompPed WHERE " & CONDI$
     FLEXCONN.Execute (SQLX$)
     '
     SQLX$ = "UPDATE PropCompComp SET MARBORRA = 1 WHERE NumPropCompra = " & NROPOP&
     FLEXCONN.Execute (SQLX$)
     
     SQLX$ = "INSERT INTO PropCompComp"
     SQLX$ = SQLX$ + " (Nro_OcomText,fechemis,MARBORRA,NumPropCompra)"
     SQLX$ = SQLX$ + " VALUES('" & Left$(Me.TXTOCOMPRAS, 256) & "'"
     SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'"
     SQLX$ = SQLX$ + " ,0"
     SQLX$ = SQLX$ + " ," & NROPOP&
     SQLX$ = SQLX$ + " )"
     FLEXCONN.Execute (SQLX$)
     CONDI$ = "NumPropCompra = " & NROPOP&
     CONDI$ = CONDI$ + " AND MARBORRA = 1"
     SQLX$ = "DELETE FROM PropCompComp WHERE " & CONDI$
     FLEXCONN.Execute (SQLX$)

     
     LBLNUMERO = NROPOP&
     
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
    End If

99   Command3.Enabled = True

End Sub

Private Sub Form_Load()
    Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
    Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
    Me.MSF1.Rows = 2
    Call CARGA_ENCABEZADO(Me.MSF3, CamposMSF3$, Me)
End Sub
Function CamposMSF3$()

   Campos$ = "CI/A0"    '  1
   Campos$ = Campos$ + ";Código/A15"                   '  2
   Campos$ = Campos$ + ";Componentes/A25"              '  3
   Campos$ = Campos$ + ";U.M./A2"                      '  4
   Campos$ = Campos$ + ";Necesidad/F11.4"              '  5
   Campos$ = Campos$ + ";St.Min./F8.0"              '  6
   Campos$ = Campos$ + ";LoteMax/F8.0"              '  7
   Campos$ = Campos$ + ";Indivisible/F9.2"            '  8
   Campos$ = Campos$ + ";Compra Sugerida/F10.2"        '  9
   Campos$ = Campos$ + ";Stock/F9.2"                  '  10
   Campos$ = Campos$ + ";Compra Autorizada/F9.2"      '  11
   Campos$ = Campos$ + ";Prov./F5"                  '  12
   Campos$ = Campos$ + ";R.Social/A15"              '  13
   Campos$ = Campos$ + ";Comp.Pend./F9.2"              '  14
   
   
   CamposMSF3$ = Campos$

End Function
   
   
  



Private Sub LBLNUMERO_Change()
   Call PRESENTA_INFO_SEGUN_PROPUESTA(XVALO(LBLNUMERO))
   'PRIMERO ORDENO POR LA COLUMNA DEL PROVEEDOR
   If MSF3.Rows < 2 Then MSF3.Rows = 2
   Call OrdenarGrid(MSF3, 11, 2)
   
   'AHORA AGREGO UNA FILA VACIA CADA VEZ QUE CAMBIA EL PROVEEDOR
   Call AgregarFilaVaciaCuandoCambiaElDato(MSF3, 11, 2)

End Sub
Sub PRESENTA_INFO_SEGUN_PROPUESTA(NROPOP&)
     If NROPOP& < 1 Then Exit Sub
     REG& = 0
     SQLX$ = "SELECT * FROM ENCA_PROPUCOMP WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NumPropCompra = " & NROPOP&
     SQLX$ = SQLX$ + " AND MARBORRA = 0"
     Set RS0 = READSET(SQLX$)
     With RS0
        Do While Not .EOF
         REG& = REG& + 1
         If REG& = 1 Then
            TxtReferencia.TEXT = SAFETEXT$(!REFERENCIA)
            LBLFECHA = FECHATEX$(CVINT(!FeInsert))
         End If
         MSF1.Rows = REG& + 1
         CI1% = XVALO(!CODINT)
         
         MSF1.TextMatrix(REG&, 1) = CODEXT$(CI1%)
         MSF1.TextMatrix(REG&, 2) = TRIM$(SAFETEXT$(DESCRIT0$(CI1%)))
         MSF1.TextMatrix(REG&, 3) = TRIM$(FORMATNUM$(XVALO(!Cantidad), "F10.1"))
         .MoveNext
        Loop
     End With
     RS0.Close
     Set RS0 = Nothing
     '
     REG& = 0
     SQLX$ = "SELECT * FROM DET_PROPUCOMP WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NumPropCompra = " & NROPOP&
     SQLX$ = SQLX$ + " AND MARBORRA = 0"
     Set rs = READSET(SQLX$)
     With rs
         Do While Not .EOF
           CI1% = XVALO(!CODINT)
           NECESIDAD = XVALO(!NECESIDAD)
           Indivisible = XVALO(!Indivisible)
           COMPRASUGERIDA = XVALO(!CompSugerida)
           Stock = XVALO(!Stock)
           AUTORIZADA = XVALO(!CompAuto)
           NPROV& = XVALO(!Num_Prov)
           oCompPendEntrega = XVALO(!OCPendEnt)
           REG& = REG& + 1
           MSF3.Rows = REG& + 1

           MSF3.TextMatrix(REG&, 1) = CI1%
           MSF3.TextMatrix(REG&, 2) = CODEXT$(CI1%)
           MSF3.TextMatrix(REG&, 3) = TRIM$(SAFETEXT$(DESCRIT0$(CI1%)))
           MSF3.TextMatrix(REG&, 4) = Unimed$(CI1%)
           MSF3.TextMatrix(REG&, 5) = FORMATNUM$(NECESIDAD, "F11.4")
           MSF3.TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(STOMIN(CI1%), "F9.0"))
           MSF3.TextMatrix(REG&, 7) = TRIM$(FORMATNUM$(LOREPOMA(CI1%), "F9.0"))
           MSF3.TextMatrix(REG&, 8) = FORMATNUM$(Indivisible, "F9.2")
           MSF3.TextMatrix(REG&, 9) = FORMATNUM$(COMPRASUGERIDA, "F9.2")
           MSF3.TextMatrix(REG&, 10) = FORMATNUM$(Stock, "F9.2")
           MSF3.TextMatrix(REG&, 11) = FORMATNUM$(AUTORIZADA, "F9.2")
           MSF3.TextMatrix(REG&, 12) = NPROV&
           MSF3.TextMatrix(REG&, 13) = rasocli$(-1 * NPROV&)
           MSF3.TextMatrix(REG&, 14) = FORMATNUM$(oCompPendEntrega, "F11.4")

           .MoveNext
        Loop
     End With
     rs.Close
     Set rs = Nothing
     '
     CONDI$ = "NumPropCompra = " & NROPOP&
     CONDI$ = CONDI$ + " AND MARBORRA = 0"
     Me.TXTNROPEDIDOS = VALOBADA("PropCompPed", "NropedText", CONDI$, "NOMESS")
    
     CONDI$ = "NumPropCompra = " & NROPOP&
     CONDI$ = CONDI$ + " AND MARBORRA = 0"
     Me.TXTOCOMPRAS = VALOBADA("PropCompComp", "Nro_OcomText", CONDI$, "NOMESS")
     
99   Command3.Enabled = True


End Sub
Private Sub mnuAbrirPropCompra_Click()

    LBLNUMERO.Caption = nRoPropuestaCompraSeleccionado&

End Sub
Function nRoPropuestaCompraSeleccionado&()
    Erase RESP_ZELE_VECT
    Campos$ = "N° Propuesta/F10;Referencia/A64;Fecha/D8"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "distinct NumPropCompra,REFERENCIA,FeUpdate"
    TABLA$ = "ENCA_PROPUCOMP" 'ES UNA VISTA
    CONDI$ = "NumPropCompra > 0 and Marborra = 0"
    CONDI$ = CONDI$ + " ORDER BY NumPropCompra DESC"
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
   
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then '
       nRoPropuestaCompraSeleccionado& = XVALO(RESP_ZELE_VECT(1))
    End If
    
End Function

Private Sub mnuBorrarPropuesta_Click()
    NroProp& = nRoPropuestaCompraSeleccionado&
    If NroProp& > 0 Then
      OPX% = COMALTER%("Atenciòn !!!\Eliminación de la Propuesta (Esta Transacción es Irreversible) \\Cancelar\Eliminar")
      If OPX% < 2 Then Exit Sub
      CONDI$ = "NumPropCompra = " & NroProp&
      Call BORRAREGISTROS("ENCA_PROPUCOMP", CONDI$, REGAF&, "NOMESS")
      Call BORRAREGISTROS("DET_PROPUCOMP", CONDI$, REGAF&, "NOMESS")
    End If
    
End Sub


Private Sub mnuPrintSt_Click()
    '
   FILTROX$ = ""
   For REG& = 1 To MSF1.Rows - 1
     If REG& > 10 Then Exit For 'SI SON MAS DE 10 NO ENTRAN EN LOS CASILLEROS
     Codigo$ = MSF1.TextMatrix(REG&, 1)
     DESC$ = MSF1.TextMatrix(REG&, 2)
     CANTI = XVALO(MSF1.TextMatrix(REG&, 3))
'     XXX$ = Space$(50)
     Call REPLA(XXX$, Codigo$, 1, 16)
     Call REPLA(XXX$, DESC$, 18, 32)
     Call REPLA(XXX$, TRIM$(FORMATNUM$(CANTI, "F8.1")), 52, 8)
     If TRIM$(FILTROX$) = "" Then
        FILTROX$ = FILTROX$ + XXX$
     Else
        FILTROX$ = FILTROX$ & ";" & XXX$
     End If

   Next REG&
   If REG& - 1 < 10 Then '10 SON EL TOTAL DE CASILLEROS DEL SQF
   For i& = REG& To 10
      FILTROX$ = FILTROX$ & ";" & "-"
   Next
   End If
   FECHA$ = TRIM$(Me.LBLFECHA)
   If FECHA$ = "" Then FECHA$ = FECHATEX$(HOY(HO$))
   NUMERO$ = Me.LBLNUMERO
   REFERENCIA$ = Me.TxtReferencia
   Call CARGA_TABLA_IMPRE(MSF3, CamposMSF3$, "PROPCMPR", "Propuesta de Compra", "TABLA_IMPRE", RET_NAMECONS$, 1)
'   Call FINAL("?" & RET_NAMECONS$)
   FILTX$ = FILTROX$ & ";" & TRIM$(FECHA$) & ";" & TRIM$(NUMERO$) & ";" & TRIM$(REFERENCIA$)
   Call STDFORM("PRP-JYM", FILTX$)

End Sub

Private Sub MSF3_DblClick()
    Unload VENTABNEW
   Campos$ = CamposMSF3$
'   Campos$ = "CI/A0"                       '  1
'   Campos$ = Campos$ + ";Código/A16"                   '  2
'   Campos$ = Campos$ + ";Componentes/A48"              '  3
'   Campos$ = Campos$ + ";U.M./A4"                      '  4
'   Campos$ = Campos$ + ";Necesidad/F19.2"              '  5
'   Campos$ = Campos$ + ";Indivisible/F9.2"            '  6
'   Campos$ = Campos$ + ";Compra Sugerida/F9.2"        '  7
'   Campos$ = Campos$ + ";Stock/F9.2"                  '  8
'   Campos$ = Campos$ + ";Compra Autorizada/F9.2"      '  9
'   Campos$ = Campos$ + ";Prov./F5"                  '  10
'   Campos$ = Campos$ + ";R.Social/A24"              '  11
   
   
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Agregar Stock"
    
    VENTANA.CampEditables = "7;8;9"
    VENTANA.Inicializar = 1
    '
    Call CopiarMsfAOtroMsf(Me.MSF3, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF3, 1)
       For i& = 1 To MSF3.Rows - 1
          CI1% = XVALO(MSF3.TextMatrix(i&, 1))
          necesidadReal = XVALO(MSF3.TextMatrix(i&, 10)) - XVALO(MSF3.TextMatrix(i&, 5))
          If XVALO(MSF3.TextMatrix(i&, 10)) > 0 Then
            If necesidadReal < 0 Then
               MSF3.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(REDON_INDIVIS(CI1%, Abs(necesidadReal)), "F12.2"))
            Else
               MSF3.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(0, "F12.2"))
            End If
          End If
       Next i&
       Else
         Set VENTANA = Nothing
         Unload VENTABNEW
         Exit Sub

    End If
    
    Set VENTANA = Nothing
    Unload VENTABNEW
    
   'PRIMERO ORDENO POR LA COLUMNA DEL PROVEEDOR
   Call OrdenarGrid(MSF3, 13, 2)
   
   'AHORA AGREGO UNA FILA VACIA CADA VEZ QUE CAMBIA EL PROVEEDOR
   Call AgregarFilaVaciaCuandoCambiaElDato(MSF3, 13, 2)


End Sub

Sub ARMAEXPLOAGRUPADA2(StringCodint$, CABA, OPCI%, Optional COMPLETO%, Optional AGRUPADOPORCODIGO%)
    '
    ' '17-03-22
    Call APERBASDAT("STRUC")
600 If OPCI% = 1 Or OPCI% = 3 Then
       '

       strsql$ = "SELECT CODIELEM,PLADISPO,SUM(usobruto)AS USOBRUTO,SUM(UNIS_FORM) AS UNIS_FORM  FROM Formulas WITH(NOLOCK) "
       strsql$ = strsql$ + " WHERE CODICONJ in (" & StringCodint$ & ")"
       strsql$ = strsql$ + " GROUP BY CODIELEM,PLADISPO "
       strsql$ = strsql$ + " ORDER BY CODIELEM ASC "

       Set FORMULASTEMP = READSET(strsql$)
615    CAIT% = 0
       On Error Resume Next
       FORMULASTEMP.MoveFirst
       If Err Then
          Exit Sub
       End If
       On Error GoTo 0
       '
       Do Until FORMULASTEMP.EOF = True
         CAIT% = CAIT% + 1
         CIJ%(CAIT%) = XVALO(FORMULASTEMP!codielem)
         USOI(CAIT%) = XVALO(FORMULASTEMP!usobruto)
         PDI%(CAIT%) = XVALO(FORMULASTEMP!PLADISPO)
         '
         USOUNI = 0: CAUNIDS = 0
         CAUNIDS = XVALO(FORMULASTEMP!UNIS_FORM)
         
         If OPCI% = 1 Then
           CODXX$ = CODEXT$(CIJ%(CAIT%))
           DEXX$ = DESCRIT$(CIJ%(CAIT%))
            'INCLUYE LEYENDA BAJA PARA LOS COMPONENTES QUE ESTAN INACTIVOS
            STAT% = XVALO(VALOBADA("MASTER", "STATUS", "CODINT=" & CIJ%(CAIT%)))
            If STAT% < 0 And COMPLETO% = 1 Then
                DEXX$ = "(Baja) " & DESCRIT$(CIJ%(CAIT%))
            End If
           UNXX$ = Unimed$(CIJ%(CAIT%))
           USON = CABA * USOI(CAIT%)
           If PORMILLAR% < 1 Then
             If HABIPORMI% > 0 Then
               If USON > 0 Then
                 If USON < 0.005 Then
                   PORMILLAR% = 1
                   GoTo 615
                 End If
               End If
             End If
           End If
           If PORMILLAR% = 1 Then
               USON = 1000 * USON
             Else
               ' CONVERSION DE FORMULA
               Call CONVERUNID(USON, UNXX$)
           End If
           '
           USOX$ = CSTRING$(MKS$(USON), 4, 12, 4, 2)
           If CAUNIDS = 0 Then CAUNIDS = USON
           PDX$ = CSTRING$(MKI$(PDI%(CAIT%)), 1, 4, 0, 2)
           XXX$ = Space$(128)
           Call REPLA(XXX$, CODXX$, 1, 16)
           Call REPLA(XXX$, DEXX$, 17, 28)
           Call REPLA(XXX$, UNXX$, 47, 2)
           Call REPLA(XXX$, USOX$, 49, 12)
           Call REPLA(XXX$, PDX$, 61, 4)
           Call REPLA(XXX$, FORMATNUM(CAUNIDS, "F9.4"), 89, 9)
           Call REPLA(XXX$, FORMATNUM(PESOMET, "F7.3"), 101, 7)
           Call REPLA(XXX$, FORMATNUM(LARCORTE, "F6.1"), 111, 6)
           Call REPLA(XXX$, FORMATNUM(USOUNI, "F8.4"), 119, 8)
           XXX$ = XXX$ + REFCOMP$(CAIT%)
           EXPLOLIN(CAIT%) = XXX$
         End If
         '
619      FORMULASTEMP.MoveNext
       Loop
       '
       FORMULASTEMP.Close
       Set FORMULASTEMP = Nothing
       End If
End Sub


Private Sub TXTMOVIL_Change()

End Sub

Sub GenerarHTMLDesdeFlexGrid(flexGrid As msFlexGrid, rutaArchivo As String, IMPRIMIR%, Optional Observacion$)
    Dim Fila As Integer
    Dim columna As Integer
    Dim contenidoHTML As String

    ' Encabezado del archivo HTML
    contenidoHTML = "<html><head><title>" & Observacion$ & "</title></head><body>"

    ' Contenido de la tabla
    contenidoHTML = contenidoHTML & "<table border='1'>"

    ' Generar encabezados de columna
    contenidoHTML = contenidoHTML & "<tr>"
    For columna = 0 To flexGrid.Cols - 1
        contenidoHTML = contenidoHTML & "<th>" & flexGrid.TextMatrix(0, columna) & "</th>"
    Next columna
    contenidoHTML = contenidoHTML & "</tr>"

    ' Generar filas con datos
    For Fila = 1 To flexGrid.Rows - 1
        contenidoHTML = contenidoHTML & "<tr>"
        For columna = 0 To flexGrid.Cols - 1
            If TRIM$(flexGrid.TextMatrix(Fila, 0)) = "" Then flexGrid.CellBackColor = vbYellow
            If TRIM$(flexGrid.TextMatrix(Fila, columna)) = "" Then flexGrid.TextMatrix(Fila, columna) = "."
            contenidoHTML = contenidoHTML & "<td>" & flexGrid.TextMatrix(Fila, columna) & "</td>"
            If TRIM$(flexGrid.TextMatrix(Fila, columna)) = "." Then flexGrid.TextMatrix(Fila, columna) = ""
        Next columna
        flexGrid.CellBackColor = vbWhite
        contenidoHTML = contenidoHTML & "</tr>"
    Next Fila

    ' Cierre del archivo HTML
    contenidoHTML = contenidoHTML & "</table></body></html>"

    ' Crear el archivo HTML y escribir el contenido
    Dim archivo As Object
    Set archivo = CreateObject("Scripting.FileSystemObject")
    Dim archivoSalida As Object
    Set archivoSalida = archivo.CreateTextFile(rutaArchivo, True)
    archivoSalida.Write contenidoHTML
    archivoSalida.Close
    If IMPRIMIR% > 0 Then
'      WebBrowser1.Navigate rutaArchivo
     Call T_Espera(1)
'     WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    End If
End Sub

Private Sub printHTML_Click()
Observa$ = InputBox$("Ingrese Observaciones, (64 caracteres)", "Observación", Now)
Observa$ = Left$(Observa$, 64)
Call GenerarHTMLDesdeFlexGrid(MSF3, "C:\FLEXOFT\Archivo.html", 1, Observa$)
   
End Sub


Private Sub Text1_Change()

End Sub


Private Sub Text1_DblClick()
    
End Sub


Sub TXTOCOMPRAS_DblClick()
    Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    OC$ = TRIM$(CStr(NPX&)) & ";"
    If InStr(TXTOCOMPRAS.TEXT, OC$) > 0 Then
       Call COMUNI("Orden de Compra yá Seleccionada")
       GoTo 99
    End If
    TXTOCOMPRAS.TEXT = TXTOCOMPRAS & OC$ 'ya incluye el separador ';'
    
    
99

End Sub


