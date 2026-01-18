VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PROPUCOM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   16545
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   16545
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   16500
      _ExtentX        =   29104
      _ExtentY        =   7699
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
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
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
      Left            =   10080
      MaxLength       =   64
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   360
      Width           =   6400
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8760
      TabIndex        =   20
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9735
      TabIndex        =   19
      ToolTipText     =   "Seleccion de Cliente"
      Top             =   0
      Width           =   250
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   0
      TabIndex        =   10
      Top             =   5160
      Width           =   16455
      Begin VB.CommandButton Command9 
         Caption         =   "Generar Solicitud de Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   15360
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   1965
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   3466
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Según Cantidad Calculada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   6720
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Considerando Lote Indivisible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   5160
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   9360
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Considerando stock y  Lote Indivisible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   1
         Left            =   7800
         TabIndex        =   13
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   11760
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Propuesta Editada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   10920
         TabIndex        =   11
         Top             =   240
         Width           =   2100
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   6720
      TabIndex        =   3
      Top             =   2640
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PROPUCOM.frx":0000
      Top             =   0
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
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
      Left            =   9720
      TabIndex        =   24
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Asignar a Cliente:"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   120
      Width           =   2220
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
      Left            =   5640
      TabIndex        =   21
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
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
      Left            =   1080
      TabIndex        =   9
      Top             =   45
      Width           =   1455
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
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label2 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Decripción:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   405
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   3960
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
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
      TabIndex        =   1
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
      Begin VB.Menu s1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Begin VB.Menu mnuMenu3 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOC 
      Caption         =   "Orden de Compra"
      Begin VB.Menu mnuctoc 
         Caption         =   "Control de Generación de O.Compra"
      End
      Begin VB.Menu munanusol 
         Caption         =   "Anular Solicitudes de Orden de Compra"
      End
      Begin VB.Menu mnuvidetasol 
         Caption         =   "Visualizar Detalle de Solicitud de Compra"
      End
   End
End
Attribute VB_Name = "PROPUCOM"
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Function SelecSolicCompra&(CI1%, TITULO$)
    '
    'ARMO LA PRESENTACION DE LA GRILLA EN EL FRMZELECHO
50  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Cli/F9;R.Social/A40;Cta.Prov/F9;R.Social/A40;N.Solic./F8.0;O.Compra/f9;Nropres/A0", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Cta.Cli/F9;R.Social/A40;Cta.Prov/F9;R.Social/A40;N.Solic./F8.0;O.Compra/f9;Nropres/A0", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1500
    FRMZELECHO.Caption = TITULO$
    '
    'PRESENTO TODAS LAS SOLICITUD DE Compras RELACIONADAS CON EL PRESUPUESTO
    i& = 0
    SQLX$ = "SELECT D.Nume_Cli,D.RASO_CLI AS DRASOCLI ,Num_Prov,PR.RASO_CLI AS PRRSOCLI,NumSolCompra,Nropres FROM PRESUCOMPRA P WITH(NOLOCK)"
    SQLX$ = SQLX$ + " INNER JOIN DEUDOR12 D ON D.NUME_CLI = P.NUME_CLI"
    SQLX$ = SQLX$ + " INNER JOIN PROVED12 PR ON PR.NUME_CLI = P.NUME_CLI"
    SQLX$ = SQLX$ + " WHERE CODICONJ = " & CI1%
    SQLX$ = SQLX$ + " group by D.Nume_Cli,D.RASO_CLI  ,Num_Prov,PR.RASO_CLI ,NumSolCompra,Nropres"
    SQLX$ = SQLX$ + " ORDER BY Num_Prov,NumSolCompra"
    
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        NP% = XVALO(!Num_Prov)
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!nume_cli)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!DRASOCLI)

        FRMZELECHO.MSF2.TextMatrix(i&, 3) = XVALO(!Num_Prov)
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!PRRSOCLI)
           NSOLIC& = XVALO(!NumSolCompra)
'           OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
'        FRMZELECHO.msf2.TextMatrix(i&, 3) = OCOMPRAX&
'        If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, i&)
        FRMZELECHO.MSF2.TextMatrix(i&, 5) = NSOLIC&
        OCOMPRAX& = OrdenCompraSegunSolicitud&(NSOLIC&)
        FRMZELECHO.MSF2.TextMatrix(i&, 6) = OCOMPRAX&
        If OCOMPRAX& > 0 Then Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF2, vbYellow, i&)
        FRMZELECHO.MSF2.TextMatrix(i&, 7) = XVALO(!Nropres)
        
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    'SI NO TIENE SOLICITUD AVISA Y VUELVE A LISTA DE PRESUPUESTOS
    If i& < 1 Then
       Call COMUNI("No se Detectaron Solicitud de Compras")
       Exit Function
    End If
    FRMZELECHO.Show 1
    '
    SelecSolicCompra& = XVALO(RESP_ZELE_VECT(3))
End Function

Private Sub Command1_Click()

End Sub

Private Sub Command19_Click()
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC = NCLIE
   Text1.TEXT = TRIM$(Str$(NC))

End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   '
   
   NC& = XVALO(Text1.TEXT)
   If NC& < 1 Then
     Call COMUNI("Falta Seleccionar Cliente")
     GoTo 99
   End If
   
'   If NP& < 1 Then
'     Call COMUNI("Falta Seleccionar Proveedor")
'     GoTo 99
'   End If
   '
   Codigo$ = Label8
   CICONJ% = CODINT%(Codigo$)
   If CICONJ% < 1 Then
     Call COMUNI("Falta Código de Conjunto")
     GoTo 99
   End If
   '
   If TRIM$(Text28.TEXT) = "" Then
     Call COMUNI("Falta Ingresar Referencia")
     GoTo 99
   End If
   Call GENERA_TABLA_COMPRAS_DE_PRESUPUESTO
   'NPRES& 'este es el identificador de toda la solicitud PUEDE ABARCAR VARIOS PROVEEDORES CON LOS ITEMS CORRESPONDIENTES PARA CADA UNO
   NPRES& = NUEVONUMPRESCOMPRA& 'este es el identificador de toda la solicitud
    Campos$ = "CI/A0"                       '  1
    Campos$ = Campos$ + ";Código/A16"                   '  2
    Campos$ = Campos$ + ";Descripción/A24"              '  3
    Campos$ = Campos$ + ";U -m/ A3"                    '  4
    Campos$ = Campos$ + ";Necesidad/F10.4"              '  5
    Campos$ = Campos$ + ";M/A4"                         '  6
    Campos$ = Campos$ + ";Pr.Unit/F10.4"                '  7
    Campos$ = Campos$ + ";Importe/F12.4"                '  8
    Campos$ = Campos$ + ";Cant-Indiv/F10.4"             '  9
    Campos$ = Campos$ + ";Importe/F12.4"                '  10
    Campos$ = Campos$ + ";Stock/F9.4"                   '  11
    Campos$ = Campos$ + ";Cant-Stock//F10.4"            '  12
    Campos$ = Campos$ + ";Importe/F12.4"                '  13
    Campos$ = Campos$ + ";Cta/f5"                       '  14
    Campos$ = Campos$ + ";R.Social/A16"                 '  15
    Campos$ = Campos$ + ";Cant.Compra/F10.4"                 '  16
    Campos$ = Campos$ + ";Imp.Compra/F12.4"                 '  17
    

    '
'    OC_GENERADA% = VALIDAR_OCOMP_GENERADA%(NPRES&)
'    If OC_GENERADA% > 0 Then
'       GoTo 99
'    End If
    '
    NORD% = 0: NPROV_ANT& = 0: VUELTA% = 0
    SQLX$ = "SELECT * FROM PRESUCOMPRA "
    SQLX$ = SQLX$ + " WHERE NroPres = " & NPRES&
    Set RS2 = New ADODB.Recordset
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS2
        For i& = 1 To MSF1.Rows - 1
             CANTI = XVALO(MSF1.TextMatrix(i&, 16))

             NPROV& = XVALO(MSF1.TextMatrix(i&, 14))
             If NPROV& > 0 And CANTI > 0 Then
                If NPROV& <> NPROV_ANT& Then
                   NORD% = 0
                   NUMSOLCOMP& = ProxNumSolCompra&
                End If
                CI1% = XVALO(MSF1.TextMatrix(i&, 1))
                COSTO1# = XVALO(MSF1.TextMatrix(i&, 7))
                
                NORD% = NORD% + 1
                .AddNew
                !Nropres = NPRES&
                !nume_cli = NC&
                !Cantidad = CANTI
                  FEC1 = CVDAT(HOY(HO$))
                !FECHEMISI = FEC1
                !Num_Prov = NPROV&
                !CODINT = CI1%
                !MARBORRA = 0
                !Status = 0
                !ORDEN = NORD%
                !COSTO = COSTO1#
                !NumSolCompra = NUMSOLCOMP&
                !mone_cos = 1
                !REFERENCIA = Left$(Text28, 64)
                !CODICONJ = CICONJ%
                !CODXCONJ = Label8
                !DescripCodConj = Label2
                .Update

                VUELTA% = 1
             End If
             '
          NPROV_ANT& = NPROV&
       Next i&
   End With
   'CARGAR LA GRILLA MOSTRANDO LOS ITEMS CARGADO
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40", FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Cta.Prov/F9;R.Social/A40", FRMZELECHO, , , ANTOT)
    If ANTOT < 5500 Then ANTOT = 5500
    FRMZELECHO.Width = ANTOT + 1100
    FRMZELECHO.Caption = "Solicitud de Compra Generada n°: " & NPRES&
    '
    i& = 0
    FRMZELECHO.MSF2.Rows = 1
    SQLX$ = "SELECT DISTINCT(NroPres),Num_Prov FROM PRESUCOMPRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPres = " & NPRES&
    SQLX$ = SQLX$ + " ORDER BY Num_Prov"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        NP& = XVALO(!Num_Prov)
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = NP&
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = rasocli$(-1 * NP&)
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    FRMZELECHO.Show 1
    
    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
    '
    If NPROV& > 0 Then
       Call COMUNI("Para Generar Ordenes de Compra Deberá Realizarlas desde el Menú: [Orden de Compra] > [Control de Generación de Orden de Compra]")
    End If
    '


99 Command9.Enabled = True


End Sub


Private Sub Form_Activate()
MENU1% = 0
MENU2% = 0
MENU3% = 0
SeleccionaFilaDobleClick% = 0
Call refresca_botones(Me)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     Unload Me
   End If
End Sub

'Private Sub Form_Resize()
'   On Error Resume Next
'   Me.MSF_2.Width = Me.Width - 300
'   Me.GRID.Width = Me.Width - 300
'   Me.MSF_2.Height = Me.Height - 1600
'   Me.GRID.Height = Me.Height - 1600
'   Frame1.Top = Me.Height - 1300
'   Frame2.Top = Me.Height - 1300
'   On Error GoTo 0
'End Sub

Private Sub cmdcancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()

'INSTALACION
'REPORTE PRESCOM2
End Sub
Private Sub Form_Resize()
   On Error Resume Next
   Me.MSF1.Width = Me.Width - 300
   Me.MSF2.Width = Me.Width - 1500
   Me.MSF1.Height = Me.Height - (Me.Frame1.Height + 1000)
   Me.Frame1.Top = Me.Height - (Me.Frame1.Height + 300)
   Frame1.Width = Me.Width - 300
   Command9.Left = Frame1.Width - (Command9.Width + 200)
   
   On Error GoTo 0
End Sub

Sub Label8_Change()
    CI1% = CODINT%(Label8)
    If CI1% < 1 Then Exit Sub
    Label2 = DESCRIT0$(CI1%)
    
      '*******************************************
    PROPUCOM.mnuMenu1.Caption = "Imprimir"
    PROPUCOM.mnuMenu1.Visible = True
    PROPUCOM.MSF1.ZOrder 0
    
    PROPUCOM.SinEfectoClickEnGrilla% = 1
    Call CargaPropuestaCompras(CI1%)

    


End Sub

Sub ARMAEXPLOAGRUPADA(CIBAS%, CABA, OPCI%, Optional COMPLETO%, Optional AGRUPADOPORCODIGO%)
    '
    ' '17-03-22
    Call APERBASDAT("STRUC")
600 If OPCI% = 1 Or OPCI% = 3 Then
       '

       strsql$ = "SELECT CODIELEM,PLADISPO,SUM(usobruto)AS USOBRUTO,SUM(UNIS_FORM) AS UNIS_FORM  FROM Formulas WITH(NOLOCK) "
       strsql$ = strsql$ + " WHERE CODICONJ = " & CIBAS%
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
         CIJ%(CAIT%) = XVALO(FORMULASTEMP!CODIELEM)
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

Sub CargaPropuestaCompras(CI1%)
   '
   If CI1% < 1 Then Exit Sub
   '
   CANTIACALCULAR = XVALO(PROPUCOM.Label5)
 Campos$ = "CI/A0"                       '  1
 Campos$ = Campos$ + ";Código/A16"                   '  2
 Campos$ = Campos$ + ";Descripción/A24"              '  3
 Campos$ = Campos$ + ";U -m/ A3"                    '  4
 Campos$ = Campos$ + ";Necesidad/F10.4"              '  5
 Campos$ = Campos$ + ";M/A4"                         '  6
 Campos$ = Campos$ + ";Pr.Unit/F10.4"                '  7
 Campos$ = Campos$ + ";Importe/F12.4"                '  8
 Campos$ = Campos$ + ";Cant-Indiv/F10.4"             '  9
 Campos$ = Campos$ + ";Importe/F12.4"                '  10
 Campos$ = Campos$ + ";Stock/F9.4"                   '  11
 Campos$ = Campos$ + ";Cant-Stock//F10.4"            '  12
 Campos$ = Campos$ + ";Importe/F12.4"                '  13
 Campos$ = Campos$ + ";Cta/f5"                       '  14
 Campos$ = Campos$ + ";R.Social/A16"                 '  15
 Campos$ = Campos$ + ";Cant.Compra/F10.4"                 '  16
 Campos$ = Campos$ + ";Imp.Compra/F12.4"                 '  17
 CODD$ = CODEXT$(CI1%)
   Screen.MousePointer = 11
   PORMILLAR% = 0
   
   CAIT% = 0
 
   If CI1% > 0 Then
      Call ARMAEXPLOAGRUPADA(CI1%, 1, 1, 0)
   End If

   'Call LEEEXPLO(CODD$, 1)
   '
   ACUMULAIMPORTESEGUNCANTIDADINGRESADA = 0
   ACUMULAIMPORTESEGUNCANTIDALOTEINDIVISIBLE# = 0
   ACUMULAImporteConsiderandoStockYLoteIndivis# = 0
'   List1.Clear
'   Call SETULTREG("!BOLREDET", 0)
   For i% = 1 To CAIT%
     ZZ1$ = EXPLOLIN$(i%)
     CICOMPO% = CODINT%(TRIM$(Left$(ZZ1$, 15)))
     Codigo$ = TRIM$(Left$(ZZ1$, 15))
     DESCRIX$ = TRIM$(Mid$(ZZ1$, 17, 28))
     UNMED$ = Unimed$(CICOMPO%)
     CANTIUNI# = XVALO(Mid$(ZZ1$, 49, 12))
     
     CANTICALCULADA# = CANTIUNI# * CANTIACALCULAR
     MONECO% = MONECOSTO%(CICOMPO%)
     SIMBOLOMONEDA$ = SIMBOLPES$(MONECO%)
     COSTOUNITARIO = FORMATNUM$(COSUNI(CICOMPO%), "F9.4")
     IMPORTESEGUNCANTIDADINGRESADA = COSPESOS(CICOMPO%) * CANTICALCULADA#
     CANTISEGUNLOTEINDIVISIBLE# = REDON_INDIVIS(CICOMPO%, CANTICALCULADA#) 'COMPRA SIN CONSIDERAR STOCK
     XX = COSPESOS(CICOMPO%)
     IMPORTESEGUNCANTIDALOTEINDIVISIBLE# = COSPESOS(CICOMPO%) * CANTISEGUNLOTEINDIVISIBLE#
     STOCKDISPONIBLE = STOCKDIS(CICOMPO%)
     DifEntreStockYNecesidad# = 0
     If STOCKDISPONIBLE > 0 Then
        DifEntreStockYNecesidad# = CANTICALCULADA# - STOCKDISPONIBLE
        If DifEntreStockYNecesidad# < 0 Then DifEntreStockYNecesidad# = 0
        If ConsideraComprasPendientes% > 0 Then ' sin activar aun
           COMPRAS_PENDIENTE = STOCOPEN(CICOMPO%)
           DifEntreStockYNecesidad# = DifEntreStockYNecesidad# + COMPRAS_PENDIENTE
        End If
        
     Else
        DifEntreStockYNecesidad# = CANTICALCULADA#
        STOCKDISPONIBLE = 0
     End If
     
     DeberiaComprarConsiderandoElStock = 0: ImporteConsiderandoStockYLoteIndivis# = 0
'     If DifEntreStockYNecesidad# > 0 Then
        DeberiaComprarConsiderandoElStock = REDON_INDIVIS(CICOMPO%, DifEntreStockYNecesidad#)
        ImporteConsiderandoStockYLoteIndivis# = COSPESOS(CICOMPO%) * DeberiaComprarConsiderandoElStock
'     End If
     CuentaProv = PROVHABI(CICOMPO%)
     razonSocial$ = rasocli$((-1) * CuentaProv)
     MSF1.Rows = i% + 1
     MSF1.TextMatrix(i%, 1) = CICOMPO%
     MSF1.TextMatrix(i%, 2) = Codigo$ 'CODIGO
     MSF1.TextMatrix(i%, 3) = DESCRIX$
     MSF1.TextMatrix(i%, 4) = UNMED$
     MSF1.TextMatrix(i%, 5) = FORMATNUM$(CANTICALCULADA#, "F12.4") 'Cantidad Unitaria * Cantidad ingresada por usuario
     MSF1.TextMatrix(i%, 6) = SIMBOLOMONEDA$
     If SIMBOLOMONEDA$ = "$" Then
         MSF1.TextMatrix(i%, 7) = FORMATNUM$(COSPESOS(CICOMPO%), "F12.4")   'Costo en pesos x Unidad ESTA ERA LA UNICA OPCION
     Else  '17-03-22  SE AGREGA ESTE ELSE PARA QUE LOS IMPORTES QUE ESTAN EN DOLARES NO CALCULE EL COSTO DE NACIONALIZACION COMO HACE COSPESOS
         MSF1.TextMatrix(i%, 7) = FORMATNUM$(COSUNI(CICOMPO%) * XVALO(ANACOSI09.Label11.Caption), "F12.4")  'EL COSTO SIN % NACIONALIZACION
     End If
     
     MSF1.TextMatrix(i%, 8) = FORMATNUM$(IMPORTESEGUNCANTIDADINGRESADA, "F12.4") 'Costo en pesos x Unidad * (Cantidad Calculada (uso unit * Cant Ingresada))
     ACUMULAIMPORTESEGUNCANTIDADINGRESADA = ACUMULAIMPORTESEGUNCANTIDADINGRESADA + IMPORTESEGUNCANTIDADINGRESADA
     MSF1.TextMatrix(i%, 9) = FORMATNUM$(CANTISEGUNLOTEINDIVISIBLE#, "F12,4")    'Multiplo Superior Inmediato de Cantidad Calculada)
     MSF1.TextMatrix(i%, 10) = FORMATNUM$(IMPORTESEGUNCANTIDALOTEINDIVISIBLE#, "F12,4") 'Costo unitario $ * Multiplo Superior Inmediato de Cantidad Calculada
     ACUMULAIMPORTESEGUNCANTIDALOTEINDIVISIBLE# = ACUMULAIMPORTESEGUNCANTIDALOTEINDIVISIBLE# + IMPORTESEGUNCANTIDALOTEINDIVISIBLE#
     MSF1.TextMatrix(i%, 11) = FORMATNUM$(STOCKDISPONIBLE, "F12,4")   'Stock Disponible (Stock - Reservado)
     MSF1.TextMatrix(i%, 12) = FORMATNUM$(DeberiaComprarConsiderandoElStock, "F12,4") 'Propuesta de Compra Considerando el Stock disponible y Cantidad Indivisible
     MSF1.TextMatrix(i%, 13) = FORMATNUM$(ImporteConsiderandoStockYLoteIndivis#, "F12,4") 'importe de 'Propuesta de Compra Considerando el Stock disponible y Cantidad Indivisible
     ACUMULAImporteConsiderandoStockYLoteIndivis# = ACUMULAImporteConsiderandoStockYLoteIndivis# + ImporteConsiderandoStockYLoteIndivis#
     MSF1.TextMatrix(i%, 14) = CuentaProv
     MSF1.TextMatrix(i%, 15) = razonSocial$
     MSF1.TextMatrix(i%, 16) = FORMATNUM$(CANTISEGUNLOTEINDIVISIBLE#, "F12,4")
     MSF1.TextMatrix(i%, 17) = FORMATNUM$(IMPORTESEGUNCANTIDALOTEINDIVISIBLE#, "F12,4")

     
   Next i%
   '
   '   'ORDENO POR LA FILA DE PROVEEDOR
    MSF1.COL = 14 'COLUMNA PROVEEDOR
    MSF1.Sort = 3

   Label7(0) = FORMATNUM$(ACUMULAIMPORTESEGUNCANTIDADINGRESADA, "F12.2")
   Label7(1) = FORMATNUM$(ACUMULAIMPORTESEGUNCANTIDALOTEINDIVISIBLE#, "F12.2")
   Label7(2) = FORMATNUM$(ACUMULAImporteConsiderandoStockYLoteIndivis#, "F12.2")
   Screen.MousePointer = 1
   Call ActualizarSubTotalProveedor
   PROPUCOM.Show 1
   '


End Sub
Sub ActualizarSubTotalProveedor()
' Campos$ = "CI/A0"                                   '  1
' Campos$ = Campos$ + ";Código/A16"                   '  2
' Campos$ = Campos$ + ";Descripción/A24"              '  3
' Campos$ = Campos$ + ";U-M/A3"                       '  4
' Campos$ = Campos$ + ";Necesidad/F10.4"              '  5
' Campos$ = Campos$ + ";M/A4"                         '  6
' Campos$ = Campos$ + ";Pr.Unit/F10.4"                '  7
' Campos$ = Campos$ + ";Importe/F12.4"                '  8
' Campos$ = Campos$ + ";Cant-Indiv/F10.4"             '  9
' Campos$ = Campos$ + ";Importe/F12.4"                '  10
' Campos$ = Campos$ + ";Stock/F9.4"                   '  11
' Campos$ = Campos$ + ";Cant-Stock/F10.4"             '  12
' Campos$ = Campos$ + ";Importe/F12.4"                '  13
' Campos$ = Campos$ + ";Cta/f5"                       '  14
' Campos$ = Campos$ + ";R.Social/A16"                 '  15
' Campos$ = Campos$ + ";Cant.Compra/F10.4"            '  16
' Campos$ = Campos$ + ";Imp.Compra/F12.4"             '  17
 
 Campos$ = "Cta/f5"                                   '  1
 Campos$ = Campos$ + ";R.Social/A16"                  '  2
 Campos$ = Campos$ + ";Imp.Seg.Necesidad/F24.4"       '  3
 Campos$ = Campos$ + ";Imp.Seg.ant-Indiv./F24.4"      '  4
 Campos$ = Campos$ + ";Imp.Cons.Stock/F24.4"          '  5
 Campos$ = Campos$ + ";Imp.Seg.Comprak/F24.4"          ' 6
 Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
 'ORDENO POR LA FILA DE PROVEEDOR
  MSF1.COL = 14 'COLUMNA PROVEEDOR
  MSF1.Sort = 3
  For i& = 1 To MSF1.Rows - 1
     NP& = XVALO(MSF1.TextMatrix(i&, 14))
     If NP& <> NP_ANT& Then
        'GENERO LOS SUBTOTALES SEGLUN DIFERETES (NECESIDAD, STOCK,....)
        SUBTOT1 = SUMA_CANT_SEGUNDATO(MSF1, SAFETEXT$(NP&), 14, 8)
        SUBTOT2 = SUMA_CANT_SEGUNDATO(MSF1, SAFETEXT$(NP&), 14, 10)
        SUBTOT3 = SUMA_CANT_SEGUNDATO(MSF1, SAFETEXT$(NP&), 14, 13)
        SUBTOT4 = SUMA_CANT_SEGUNDATO(MSF1, SAFETEXT$(NP&), 14, 17)
        MSF2.Rows = MSF2.Rows + 1
        MSF2.TextMatrix(MSF2.Rows - 1, 1) = NP&
        MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF1.TextMatrix(i&, 15)
        MSF2.TextMatrix(MSF2.Rows - 1, 3) = FORMATNUM$(XVALO(SUBTOT1), "F12.2")
        MSF2.TextMatrix(MSF2.Rows - 1, 4) = FORMATNUM$(XVALO(SUBTOT2), "F12.2")
        MSF2.TextMatrix(MSF2.Rows - 1, 5) = FORMATNUM$(XVALO(SUBTOT3), "F12.2")
        MSF2.TextMatrix(MSF2.Rows - 1, 6) = FORMATNUM$(XVALO(SUBTOT4), "F12.2")
     End If
     NP_ANT& = NP&
  Next i&
  MSF2.Rows = MSF2.Rows + 1
  MSF2.TextMatrix(MSF2.Rows - 1, 3) = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF2, 3)), "F12.2")
  MSF2.TextMatrix(MSF2.Rows - 1, 4) = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF2, 4)), "F12.2")
  MSF2.TextMatrix(MSF2.Rows - 1, 5) = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF2, 5)), "F12.2")
  MSF2.TextMatrix(MSF2.Rows - 1, 6) = FORMATNUM$(XVALO(SUMA_COLUMNA_SELECCION(MSF2, 6)), "F12.2")
  Call COLOREAR_GRILLA_SOLAFILA(MSF2, 4353311, MSF2.Rows - 1)
  Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF2, 16251129, MSF2.Rows - 1)
End Sub

Private Sub mnuctoc_Click()
10  COD$ = ListaSeleccionArticulos("", True)
    CI1% = CODINT%(COD$)
    '
    If COD$ = "" Then Exit Sub
    
    NPROV& = SelecSolicCompra&(CI1%, "Selección de Solicitudes Para Generar Orden de Compra")
    
    If NPROV& < 1 Then GoTo 10
    OCOMPRA& = XVALO(RESP_ZELE_VECT(6))
    NumePresu& = XVALO(RESP_ZELE_VECT(7))
    '
    'SI TIENE ORDEN DE COMPRA AVISA Y VUELVE A LA LISTA DE SELECCION
    If OCOMPRA& > 0 Then
       Call COMUNI("Solicitud Seleccionada Generó La Orden de Compra: " & OCOMPRA& & "\Modifique Orden de Compra o Anular Para Generar una Nueva.")
       Exit Sub
    End If
    '
    If NPROV& > 0 Then
       APLI$ = "CPRE" & NPROV& & "_" & NumePresu&
       Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra x Presupuesto")
    Else
       GoTo 10
    End If
    '
End Sub

Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub MSF1_DblClick()
 Campos$ = "CI/A0"                       '  1
 Campos$ = Campos$ + ";Código/A16"                   '  2
 Campos$ = Campos$ + ";Descripción/A24"              '  3
 Campos$ = Campos$ + ";U-M/A3"                    '  4
 Campos$ = Campos$ + ";Necesidad/F10.4"              '  5
 Campos$ = Campos$ + ";M/A4"                         '  6
 Campos$ = Campos$ + ";Pr.Unit/F10.4"                '  7
 Campos$ = Campos$ + ";Importe/F12.4"                '  8
 Campos$ = Campos$ + ";Cant-Indiv/F10.4"             '  9
 Campos$ = Campos$ + ";Importe/F12.4"                '  10
 Campos$ = Campos$ + ";Stock/F9.4"                   '  11
 Campos$ = Campos$ + ";Cant-Stock/F10.4"            '  12
 Campos$ = Campos$ + ";Importe/F12.4"                '  13
 Campos$ = Campos$ + ";Cta/F5"                       '  14
 Campos$ = Campos$ + ";R.Social/A16"                 '  15
 Campos$ = Campos$ + ";Cant.Compra/F10.4"                 '  16
 Campos$ = Campos$ + ";Imp.Compra/F12.4"                 '  17
  ACUMULAImporteEditado# = 0

    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    'VENTANA.formula(15) = "=EXCEL;(L1/K1)*(L1/K1)*M1"
    'VENTANA.formula(14) = "=EXCEL;(12/11)*(12/11)*13"
    VENTANA.Formula(17) = "7;*;16"
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Propuesta de Compra"
    
    VENTANA.ColEcoCampo(14) = "nume_cli;Raso_Cli;PROVED12"
    VENTANA.CampEditables = "14;16"
    VENTANA.VentVectorToolTipText(5) = "Cantidad Unitaria * Cantidad Ingresada por el Usuario"
    VENTANA.VentVectorToolTipText(7) = "Precio unitario en pesos"
    VENTANA.VentVectorToolTipText(8) = "Precio unitario en pesos * (Cantidad Calculada (Uso unit * Cant Ingresada))"
    VENTANA.VentVectorToolTipText(9) = "Multiplo Superior Inmediato de Lote Indivisible de Cantidad Calculada"
    VENTANA.VentVectorToolTipText(10) = "Precio unitario en pesos * Multiplo Superior Inmediato de Cantidad Calculada"

    VENTANA.VentVectorToolTipText(11) = "Stock Disponible (Stock - Reservado)"
    VENTANA.VentVectorToolTipText(12) = "Propuesta de Compra Considerando el Stock disponible y Cantidad Indivisible"
    VENTANA.VentVectorToolTipText(13) = "importe de 'Propuesta de Compra Considerando el Stock disponible y Cantidad Indivisible"
    VENTANA.TOTALES(8) = "S"
    VENTANA.TOTALES(10) = "S"
    VENTANA.TOTALES(13) = "S"
    VENTANA.TOTALES(17) = "S"

    VENTANA.Inicializar = 1
    '
    Call CopiarMsfAOtroMsf(Me.MSF1, VENTABNEW.MSF, 1)
    
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    VENTABNEW.Width = Me.Width
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
       For i& = 1 To MSF1.Rows - 1
            ACUMULAImporteEditado# = ACUMULAImporteEditado# + XVALO(MSF1.TextMatrix(i&, 17))
       Next i&

    End If
    Label7(3) = FORMATNUM$(ACUMULAImporteEditado#, "F12.2")
    Unload VENTABNEW
     
    'ORDENO POR LA FILA DE PROVEEDOR
    MSF1.COL = 14 'COLUMNA PROVEEDOR
    MSF1.Sort = 3
    Call ActualizarSubTotalProveedor
 
End Sub



Private Sub MSF2_DblClick()
    
    i& = MSF2.MouseRow
    j& = MSF2.MouseCol
    If i& = 0 Or i& > MSF2.Rows Then Exit Sub
    '


    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    
 
    NP& = XVALO(MSF2.TextMatrix(REG&, 1))
    If NP& < 1 Then
       Call COMUNI("Falta Número de Proveedor en el Registro Seleccionado")
       Exit Sub
    End If
    
    OPX% = COMALTER%("Confirma Cargar Datos para Generar Orden de Compra al Proveedor: " & rasocli$((-1) * NP&) & "\\Cancelar\Cargar Orden de Compra")
    If OPX% < 2 Then Exit Sub
    
    NPRES& = NUEVONUMPRESCOMPRA& 'este es el identificador de toda la solicitud

    NORD% = 0: NPROV_ANT& = 0: VUELTA% = 0
    SQLX$ = "SELECT * FROM PRESUCOMPRA "
    SQLX$ = SQLX$ + " WHERE NroPres = " & NPRES&
    Set RS2 = New ADODB.Recordset
    RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS2
        For i& = 1 To MSF1.Rows - 1
             CANTI = XVALO(MSF1.TextMatrix(i&, 16))

             NPROV& = XVALO(MSF1.TextMatrix(i&, 14))
             If NPROV& = NP& And CANTI > 0 Then
                If VUELTA% < 1 Then
                   NUMSOLCOMP& = ProxNumSolCompra&
                   VUELTA% = 1
                End If
                CI1% = XVALO(MSF1.TextMatrix(i&, 1))
                COSTO1# = XVALO(MSF1.TextMatrix(i&, 7))
                NORD% = NORD% + 1
                .AddNew
                !Nropres = NPRES&
                !nume_cli = 0
                !Cantidad = CANTI
                  FEC1 = CVDAT(HOY(HO$))
                !FECHEMISI = FEC1
                !Num_Prov = NP&
                !CODINT = CI1%
                !MARBORRA = 0
                !Status = 0
                !ORDEN = NORD%
                !COSTO = COSTO1#
                !NumSolCompra = NUMSOLCOMP&
                !mone_cos = 1
                !REFERENCIA = Left$(Text28, 64)
                !CODICONJ = CICONJ%
                !CODXCONJ = Label8
                !DescripCodConj = Label2
                .Update
             End If
             '
       Next i&
   End With
   
   APLI$ = "CPRE" & NP& & "_" & NPRES&
   Call CONECRUN("ORCOMP07/" + APLI$, "Generacion de Ordenes de Compra x Presupuesto")

End Sub


Private Sub munanusol_Click()


10  COD$ = ListaSeleccionArticulos("", True)
    CI1% = CODINT%(COD$)
    '
    If COD$ = "" Then Exit Sub
    
50  NPROV& = SelecSolicCompra&(CI1%, "Anulación De Solicitud de Compras")
    
    If NPROV& < 1 Then GoTo 10
    OCOMPRA& = XVALO(RESP_ZELE_VECT(6))
    NSOLIC& = XVALO(RESP_ZELE_VECT(5))
    
'    NPROV& = XVALO(FRMZELECHO.RESP_ZELE$)
'    OCOMPRA& = XVALO(RESP_ZELE_VECT(3))
'    NSOLIC& = XVALO(RESP_ZELE_VECT(4))
    '
    CONDI$ = " NUME_OCOM = " & OCOMPRA& & " AND Stat_Ocom < 9"
    OCOMPRAOCOMENCA& = XVALO(VALOBADA("OCOMENCA", "NUME_OCOM", CONDI$, "NOMESS"))
    'SI TIENE ORDEN DE COMPRA AVISA Y VUELVE A LA LISTA DE SELECCION
    If OCOMPRAOCOMENCA& > 0 Then
       Call COMUNI("Solicitud Seleccionada Generó La Orden de Compra: " & OCOMPRA& & "\Modifique Orden de Compra o Anular Para Generar una Nueva.")
       GoTo 50
    End If
    '
    If NSOLIC& > 0 Then
        OPX% = COMALTER%("Confirma Realmente Anular Solicitud de Compra N°: " & NSOLIC& & "\\ No,Cancelar\Si,Anular")
        If OPX% = 2 Then
           CONDI$ = "NumSolCompra = " & NSOLIC&
           Call BORRAREGISTROS("PRESUCOMPRA", CONDI$, REGAF&, "NOMESS")
           Call COMUNI("Anulación Procesada")
           GoTo 50
        Else
           GoTo 50
        End If
    End If
    '
End Sub

Private Sub mnuvidetasol_Click()
10  COD$ = ListaSeleccionArticulos("", True)
    CI1% = CODINT%(COD$)
    '
    If COD$ = "" Then Exit Sub
    
    NPROV& = SelecSolicCompra&(CI1%, "Seleccione Para Visualizar Detalle")
    
    If NPROV& < 1 Then GoTo 10
    OCOMPRA& = XVALO(RESP_ZELE_VECT(6))
    NSOLIC& = XVALO(RESP_ZELE_VECT(5))
    '
    If NSOLIC& > 0 Then
       Call FINAL("? PRESCOM2 (" & TRIM$(SAFETEXT$(NSOLIC&)) & ")")
    End If
    'Call FINAL("? MOVCONSU (" & NORFA$ & ")")

End Sub


Private Sub Text1_Change()
 Label11 = rasocli$(XVALO(Text1))
End Sub


