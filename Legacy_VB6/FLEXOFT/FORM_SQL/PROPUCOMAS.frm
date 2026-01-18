VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form PROPUCOMAS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBCA9D&
   BorderStyle     =   1  'Fixed Single
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
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
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
      Left            =   11040
      TabIndex        =   13
      Top             =   7750
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>>>Compra Sugerida>>>>"
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
      Left            =   13800
      TabIndex        =   10
      Top             =   7750
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   7725
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   17415
      Begin VB.CommandButton Command43 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2760
         Picture         =   "PROPUCOMAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Actualizar Indice de Selección"
         Top             =   120
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11280
         TabIndex        =   18
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1920
         Width           =   255
      End
      Begin VB.ListBox LIST1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   15
         ToolTipText     =   "Tilde los Pedidos a Considerar"
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar Grillas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   7
         Top             =   120
         Width           =   4455
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   17205
         _ExtentX        =   30348
         _ExtentY        =   9128
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
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   1845
         Left            =   11520
         TabIndex        =   9
         Top             =   360
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   3254
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
      Begin MSFlexGridLib.MSFlexGrid MSF4 
         Height          =   1845
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   7920
         _ExtentX        =   13970
         _ExtentY        =   3254
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
      Begin VB.CommandButton Command1 
         Caption         =   "Calcular Materiales a Comprar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13920
         TabIndex        =   8
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Pedidos Seleccionados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   22
         Top             =   120
         Width           =   3420
      End
      Begin VB.Label lblCanped 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   2160
         TabIndex        =   21
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedidos Seleccionados:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione Pedidos"
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
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle Acumulado de Pedidos Selecccionados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   11520
         TabIndex        =   11
         Top             =   120
         Width           =   4860
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "PROPUCOMAS.frx":014A
      Top             =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   7470
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   17340
      _ExtentX        =   30586
      _ExtentY        =   13176
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   16335
      ExtentX         =   28813
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
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
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIrPantGrab 
         Caption         =   "Ir a Pantalla de Grabación"
      End
      Begin VB.Menu mnuComprasPenProv 
         Caption         =   "Compras Pendirentes x Proveedor"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Imprimir"
         Visible         =   0   'False
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
Attribute VB_Name = "PROPUCOMAS"
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
Public CamposMSF2$


Sub CargarLabel11_ConCodigosInternos()
       StringCodint$ = ""
       For i& = 1 To MSF1.Rows - 1
          If TRIM$(MSF1.TextMatrix(i&, 1)) <> "" Then 'campo de la descripcion del codigo
            CODIX$ = TRIM$(MSF1.TextMatrix(i&, 1))
            CI1% = CODINT%(CODIX$)
            If i& = 1 Then
               StringCodint$ = StringCodint$ + TRIM$(Str$(CI1%))
            Else
               StringCodint$ = StringCodint$ + "," + TRIM$(Str$(CI1%))
            End If
          End If
       Next i&
       Label11 = StringCodint$

End Sub

Sub CrearTablaPropCompPed()
       'tabla PropCompPed
       Call CREACAMPO("", "PropCompPed", "MARBORRA", 3, 0, 1)
       If ValidaExisteCampo%("PropCompPed_ID", "PropCompPed") < 1 Then
          SQLX$ = "ALTER TABLE PropCompPed ADD PropCompPed_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
        End If

       Call CREACAMPO("", "PropCompPed", "NumPropCompra", 7, 0)
       Call CREACAMPO("", "PropCompPed", "NropedText", 10, 256)
       Call CREACAMPO("", "PropCompPed", "Fechemis", 8, 0)
       
       'tabla PropCompComp
       Call CREACAMPO("", "PropCompComp", "MARBORRA", 3, 0, 1)
       If ValidaExisteCampo%("PropCompComp_ID", "PropCompComp") < 1 Then
          SQLX$ = "ALTER TABLE PropCompComp ADD PropCompComp_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
        End If

       Call CREACAMPO("", "PropCompComp", "NumPropCompra", 7, 0)
       Call CREACAMPO("", "PropCompComp", "Nro_OcomText", 10, 256)
       Call CREACAMPO("", "PropCompComp", "Fechemis", 8, 0)


End Sub

Private Sub cmdGrabar_Click()
    PROPCOMGRAB.TXTNROPEDIDOS = ""
    PROPCOMGRAB.TXTOCOMPRAS = ""
    If MSF1.Rows < 2 Then MSF1.Rows = 2
    Call CopiarMsfAOtroMsf(Me.MSF1, PROPCOMGRAB.MSF1, 1)
    If MSF3.Rows < 2 Then MSF3.Rows = 2
    Call CopiarMsfAOtroMsf(Me.MSF3, PROPCOMGRAB.MSF3, 1)
    Observa$ = InputBox$("Ingrese Referencia, (Total 64 caracteres)", "Referencia", Now)
    Observa$ = Left$(Observa$, 64)
    'cargo los nros de pedidos SELECCIONADOS
    PEDIDO$ = "": PEDIDO_ACUM$ = ""
    For Fila = 1 To MSF4.Rows - 1
        PEDIDO$ = TRIM$(MSF4.TextMatrix(Fila, 1)) & ";"
        If TRIM$(PEDIDO_ANT$) <> PEDIDO$ Then
            PEDIDO_ACUM$ = PEDIDO_ACUM$ & PEDIDO$   ' Columna de PEDIDO
            PEDIDO_ANT$ = PEDIDO$
        End If
    Next Fila
    PROPCOMGRAB.TXTNROPEDIDOS = PEDIDO_ACUM$
    
    PROPCOMGRAB.TxtReferencia = Observa$
    PROPCOMGRAB.Command3_Click
    PROPCOMGRAB.MSF1 = 1
    PROPCOMGRAB.MSF3 = 1
    PROPCOMGRAB.LBLFECHA = ""
    PROPCOMGRAB.LBLNUMERO = ""
    
    '
    
    PROPCOMGRAB.MSF1 = 1
    PROPCOMGRAB.MSF3 = 1
    PROPCOMGRAB.LBLFECHA = ""
    PROPCOMGRAB.LBLNUMERO = ""

    Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
    Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
    Me.MSF1.Rows = 2
    Call CARGA_ENCABEZADO(Me.MSF2, CamposMSF2$, Me)
    MSF2.Rows = 2

    Call CARGA_ENCABEZADO(Me.MSF3, CamposMSF3$, Me)
    MSF3.Rows = 2
    Call CARGA_ENCABEZADO(Me.MSF4, CamposMSF4$, Me)
    MSF4.Rows = 2
    LIST1.Clear
    Call CARGALISTA(LIST1, "PEDENCA P + PEDDETA PD/P.NROPED = PD.NROPED", "P.NROPED/F8;P.RASOCLIE/A24", "PD.STATUS < 8 AND PD.CANPED > PD.CANTENT GROUP BY  PD.NROPED , P.NROPED,P.RASOCLIE ORDER BY PD.NROPED DESC", "")

    
    

End Sub

Private Sub Command1_Click()
   '
   Command1.Enabled = False
   If MSF1.Rows <= 1 Then GoTo 99
   '
   
   
   '
   Screen.MousePointer = 11
   PORMILLAR% = 0
   
   CAIT% = 0
   StringCodint$ = TRIM$(Label11)
   If StringCodint$ = "" Then GoTo 99
   Call TEXTOLOTES$(Label11, ",")
   CantiCodigo% = UBound(CADENATEX$)
   
   Call ARMAEXPLOAGRUPADA2(StringCodint$, 1, 1, 0)

   Campos$ = "CI/A0"                       '  1
   Campos$ = Campos$ + ";Código/A16"                   '  2
   Campos$ = Campos$ + ";Componentes/A24"              '  3
   Campos$ = Campos$ + ";U.M./A4"                      '  4
   For i% = 1 To CantiCodigo% 'cargo una columna para cada codigo
      CodigoOriginal$ = CODEXT$(XVALO(CADENATEX$(i%)))
      CodigoReemplazaBarraXGuion$ = REPLACAR$(CodigoOriginal$, "/", "-")
      Campos$ = Campos$ + ";" & CodigoReemplazaBarraXGuion$ & "/F12.4"
   Next i%
   Campos$ = Campos$ + ";Uso Unit./F12.4"
   Campos$ = Campos$ + ";Total/F12.2"
   CamposMSF2$ = Campos$
   Call CARGA_ENCABEZADO(Me.MSF2, CamposMSF2$, Me)

   'cargo la cantidad en la primera fila lluego del encabezado
    MSF2.Rows = 2
    
    MSF2.TextMatrix(1, 2) = "CANTIDAD SOLICITADA"
    Call COLOREAR_GRILLA_SOLAFILA(MSF2, &HFBCA9D, 1, 0, 0, 0)
   For i% = 1 To MSF1.Rows - 1
       MSF2.TextMatrix(1, i% + 4) = TRIM$(FORMATNUM$(XVALO(MSF1.TextMatrix(i%, 3)), "F12.2")) 'MAS 4 POR QUE SON LOS 4 QUE SE ANTEPONEN DE MANERA FIJA
       MSF2.TextMatrix(0, i% + 4) = MSF1.TextMatrix(i%, 1) 'LEYENDO LA GRILLA SUPERIOR VUELVO A PONER EL CODIGO ORIGINAL
   Next i%
   '

   'CARGA LOS COMPONENTES AGRUPADOS DE TODOS LOS ITEMS ELEGIDOS

   For i% = 1 To CAIT%
     ZZ1$ = EXPLOLIN$(i%)
     CICOMPO% = CODINT%(TRIM$(Left$(ZZ1$, 15)))
     Codigo$ = TRIM$(Left$(ZZ1$, 15))
     DESCRIX$ = TRIM$(Mid$(ZZ1$, 17, 28))
     UNMED$ = Unimed$(CICOMPO%)
     CANTIUNI# = XVALO(Mid$(ZZ1$, 49, 12))
     MSF2.Rows = i% + 2
     MSF2.TextMatrix(i% + 1, 1) = CICOMPO%
     MSF2.TextMatrix(i% + 1, 2) = Codigo$ 'CODIGO
     MSF2.TextMatrix(i% + 1, 3) = DESCRIT0$(CICOMPO%)
     MSF2.TextMatrix(i% + 1, 4) = UNMED$
     MSF2.TextMatrix(i% + 1, 5 + CantiCodigo%) = ""
'     MSF2.TextMatrix(i% + 1, 5 + CantiCodigo%) = TRIM$(FORMATNUM$(CANTIUNI#, "F12.4")) 'Cantidad Unitaria * Cantidad ingresada por usuario
   Next i%
   '
   total1 = 0
   'RECORRO LA GRILLA CON UNA FUNCION QUE ME DEVUELVA SI EXISTE EL COMPONENTE SEGUN EL CODICONJ Y APLICA EN LA COLUMNA DEBAJO DE CADA CODIGO
    For REG& = 0 To MSF2.Rows - 1
     For COL1% = 1 To CantiCodigo%
      CIELEM1% = XVALO(MSF2.TextMatrix(REG&, 1))
      If CIELEM1% > 0 Then
        CodigoConj$ = MSF2.TextMatrix(0, COL1% + 4) 'ACA SIEMPRE LEE EL CODIGO DE LA FILA 0 (4 SON LAS PRIMERAS COLUMNAS)
        Cantidad = XVALO(MSF2.TextMatrix(1, COL1% + 4)) 'ACA SIEMPRE LEE la cantidad DE LA FILA 1
        'USOX = XVALO(MSF2.TextMatrix(REG&, 5 + CantiCodigo%))
        If Es_componente%(CODINT%(CodigoConj$), CIELEM1%) Then
            USOX = Uso_de_componente#(CODINT%(CodigoConj$), CIELEM1%)
            MSF2.TextMatrix(REG&, COL1% + 4) = TRIM$(FORMATNUM$(Cantidad * USOX, "F12.4"))
            total1 = total1 + XVALO(MSF2.TextMatrix(REG&, COL1% + 4))
        End If
      End If
     Next COL1%
     If CIELEM1% > 0 Then
        MSF2.TextMatrix(REG&, 6 + CantiCodigo%) = TRIM$(FORMATNUM$(total1, "F12.4")) 'IMPRIME EL TOTAL DE ESA FILA
     End If
     total1 = 0
     If REG& = MSF2.Rows - 1 Then Exit For
   Next REG&
   
   'ACA DEBERIA CALCULAR EL USO POR LA CANTIDAD E IMPRIMIRLA EN LA GRILLA

'   For REG& = 2 To MSF2.Rows - 1
'      total1 = XVALO(SUMA_HORIZONTAL(MSF2, REG&, 5, 5 + CantiCodigo% - 1))
'      total1 = TRIM$(FORMATNUM$(total1, "F12.2"))
'      MSF2.TextMatrix(REG&, 6 + CantiCodigo%) = total1
'   Next REG&
   '
   '
   '**** GRILLA OCULTA *****
   'cargo la grilla no visible de la compra sugerida y el lote indivisilbe
   '

   Call CARGA_ENCABEZADO(Me.MSF3, CamposMSF3$, Me)
   '
   For REG& = 2 To MSF2.Rows - 1
     CI1% = XVALO(MSF2.TextMatrix(REG&, 1))
     If CI1% > 0 Then
       MSF3.Rows = REG& + 1
       MSF3.TextMatrix(REG&, 1) = MSF2.TextMatrix(REG&, 1)
       MSF3.TextMatrix(REG&, 2) = MSF2.TextMatrix(REG&, 2)
       MSF3.TextMatrix(REG&, 3) = MSF2.TextMatrix(REG&, 3)
       MSF3.TextMatrix(REG&, 4) = MSF2.TextMatrix(REG&, 4)
       MSF3.TextMatrix(REG&, 5) = MSF2.TextMatrix(REG&, CantiCodigo% + 6)
       MSF3.TextMatrix(REG&, 6) = TRIM$(FORMATNUM$(STOMIN(CI1%), "F9.0"))
       MSF3.TextMatrix(REG&, 7) = TRIM$(FORMATNUM$(LOREPOMA(CI1%), "F9.0"))
       MSF3.TextMatrix(REG&, 8) = TRIM$(FORMATNUM$(LINDIVIS(CI1%), "F12.2"))
       
       MSF3.TextMatrix(REG&, 9) = TRIM$(FORMATNUM$(REDON_INDIVIS(CI1%, XVALO(MSF3.TextMatrix(REG&, 5))), "F12.2"))
       MSF3.TextMatrix(REG&, 10) = ""
       MSF3.TextMatrix(REG&, 11) = ""
       CuentaProv = PROVHABI(CI1%)
       razonSocial$ = rasocli$((-1) * CuentaProv)
       MSF3.TextMatrix(REG&, 12) = TRIM$(Str$(CuentaProv))
       MSF3.TextMatrix(REG&, 13) = razonSocial$
       MSF3.TextMatrix(REG&, 14) = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F9.2")        ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
       

     End If
   Next REG&
   '
   'PRIMERO ORDENO POR LA COLUMNA DEL PROVEEDOR
   Call OrdenarGrid(MSF3, 12, 2)
   
   'AHORA AGREGO UNA FILA VACIA CADA VEZ QUE CAMBIA EL PROVEEDOR
   Call AgregarFilaVaciaCuandoCambiaElDato(MSF3, 12, 2)

99 Screen.MousePointer = 1
   Command1.Enabled = True

End Sub
Function CamposMSF3$()
   Campos$ = "CI/A0"    '  1
   Campos$ = Campos$ + ";Código/A15"                   '  2
   Campos$ = Campos$ + ";Componentes/25"              '  3
   Campos$ = Campos$ + ";U.M./A2"                      '  4
   Campos$ = Campos$ + ";Necesidad/F9.4"              '  5
   Campos$ = Campos$ + ";St.Min./F9.0"              '  6
   Campos$ = Campos$ + ";LoteMax/F9.0"              '  7
   Campos$ = Campos$ + ";Indivisible/F9.2"            '  8
   Campos$ = Campos$ + ";Compra Sugerida/F9.2"        '  9
   Campos$ = Campos$ + ";Stock/F9.2"                  '  10
   Campos$ = Campos$ + ";Compra Autorizada/F9.2"      '  11
   Campos$ = Campos$ + ";Prov./F5"                  '  12
   Campos$ = Campos$ + ";R.Social/A15"              '  13
   Campos$ = Campos$ + ";Comp.Pend./F9.2"              '  14
   
   
   CamposMSF3$ = Campos$
End Function

Private Sub Command2_Click()
        Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
        Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
        Me.MSF1.Rows = 2
        Call CARGA_ENCABEZADO(Me.MSF3, CamposMSF3$, Me)
        MSF3.Rows = 2
        Call CARGA_ENCABEZADO(Me.MSF2, CamposMSF2$, Me)
        MSF2.Rows = 2

   Me.Label11 = ""
End Sub

Private Sub Command3_Click()
   Dim estadoVisible As Boolean
   estadoVisible = Frame1.Visible
'    ' Llamar a la función Conmutar para cambiar el valor entre true o false segun posea al momento
   Frame1.Visible = Conmutar(estadoVisible)
   If Frame1.Visible = False Then
      Command3.Caption = "<<<< Volver <<<<<<"
      cmdGrabar.Visible = True
   Else
      Command3.Caption = ">>>> Compra Sugerida >>>>"
      cmdGrabar.Visible = False
   End If
End Sub

Private Sub Command4_Click()
    MSF4.Rows = 1
    For U& = 0 To LIST1.ListCount - 1
     If LIST1.Selected(U&) = True Then
         PEDIDO = XVALO(Left(LIST1.List(U&), 8))
         RASO$ = TRIM$(Right(LIST1.List(U&), 24))
         SQLX$ = "SELECT CODIINT,CODIEXT,SUM(CANPED)AS CANPED "
         SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) WHERE NROPED = " & PEDIDO
         SQLX$ = SQLX$ + " GROUP BY CODIINT,CODIEXT"
         SQLX$ = SQLX$ + " ORDER BY CODIINT"
         Set RS = READSET(SQLX$)
         With RS
            Do While Not .EOF
               MSF4.Rows = MSF4.Rows + 1
               REG& = MSF4.Rows - 1
              CI1% = XVALO(!CodiInt)
              Codigo$ = SAFETEXT$(!CODIEXT)
              CANTI = XVALO(!CanPed)
              If CI1% > 0 Then
                MSF4.TextMatrix(REG&, 1) = PEDIDO
                MSF4.TextMatrix(REG&, 2) = RASO$
                MSF4.TextMatrix(REG&, 3) = Codigo$
                MSF4.TextMatrix(REG&, 4) = DESCRIT0$(CI1%)
                MSF4.TextMatrix(REG&, 5) = CANTI
                MSF4.TextMatrix(REG&, 6) = CI1%
              End If
              .MoveNext
                
            Loop
         End With
         RS.Close
         Set RS = Nothing
     End If
   Next U&
   MSF4.COL = 1
   MSF4.Sort = 3 'ordena por nro de pedido

End Sub

Function CamposMSF4$()
     CamposMSF4$ = "Pedido/F3;R.Social/A14;Código/A12;Descripción/A18;Cantidad/f4.0;CI/A0"
     
End Function

Private Sub Command43_Click()
   Command43.Enabled = False
   OPPX% = COMALTER%("Si Refresca la Lista de Pedidos, Inicie Nuevamente el proceso de Selección:\\Cancelar\Refrescar")
   If OPPX% < 1 Or OPPX% > 2 Then Exit Sub
   If OPPX% = 2 Then

     Call CARGALISTA(LIST1, "PEDENCA P + PEDDETA PD/P.NROPED = PD.NROPED", "P.NROPED/F8;P.RASOCLIE/A24", "PD.STATUS < 8 AND PD.CANPED > PD.CANTENT GROUP BY  PD.NROPED , P.NROPED,P.RASOCLIE ORDER BY PD.NROPED DESC", "")
   End If
   
99 Command43.Enabled = True
   
End Sub

Private Sub Command5_Click()
    MSF1.Rows = 1
    Dim dictCodigos As Object
    Set dictCodigos = CreateObject("Scripting.Dictionary")
    Dim Fila As Integer
    For Fila = 1 To MSF4.Rows - 1
        Dim Codigo As String
        Dim Cantidad As Double
        
        Codigo = MSF4.TextMatrix(Fila, 3) ' Columna de códigos
        Cantidad = XVALO(MSF4.TextMatrix(Fila, 5)) ' Columna de cantidad
        
        If dictCodigos.Exists(Codigo) Then
            dictCodigos(Codigo) = dictCodigos(Codigo) + Cantidad
        Else
            dictCodigos.Add Codigo, Cantidad
        End If
    Next Fila
    
    MSF1.Rows = dictCodigos.Count + 1
    
    
    Dim filaDestino As Integer
    filaDestino = 1
            Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"

    Dim codigoExistente As Variant
    For Each codigoExistente In dictCodigos.Keys
        MSF1.TextMatrix(filaDestino, 1) = codigoExistente
         COD$ = codigoExistente
         CI1% = CODINT%(COD$)
        MSF1.TextMatrix(filaDestino, 2) = DESCRIT0$(CI1%)
        MSF1.TextMatrix(filaDestino, 3) = dictCodigos(codigoExistente)
        MSF1.TextMatrix(filaDestino, 4) = CI1%
        filaDestino = filaDestino + 1
    Next codigoExistente
    '
    Call CargarLabel11_ConCodigosInternos
    Set dictCodigos = Nothing
    If MSF1.Rows < 2 Then MSF1.Rows = 2
End Sub

Private Sub Form_Load()
        'Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
        Campos$ = "Código/A12;Descripción/A18;Cantidad/f4.0;CI/A0"
        Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
        Me.MSF1.Rows = 2
        Call GENERA_TABLA_PROPUCOMP
        Call CARGA_ENCABEZADO(Me.MSF3, CamposMSF3$, Me)
        MSF3.Rows = 2
        Call CARGA_ENCABEZADO(Me.MSF2, CamposMSF2$, Me)
        MSF2.Rows = 2
        
        Call CARGA_ENCABEZADO(Me.MSF4, CamposMSF4$, Me)
        MSF4.Rows = 2
        
        Call CARGALISTA(LIST1, "PEDENCA P + PEDDETA PD/P.NROPED = PD.NROPED", "P.NROPED/F8;P.RASOCLIE/A24", "PD.STATUS < 8 AND PD.CANPED > PD.CANTENT GROUP BY  PD.NROPED , P.NROPED,P.RASOCLIE ORDER BY PD.NROPED DESC", "")
        Call CrearTablaPropCompPed

    
End Sub
 Function FormatearNumero(ByVal NUMERO As Double, CANTDECIMALES) As String
    If CANTDECIMALES < 1 Then CANTDECIMALES = 2
    
    If Int(NUMERO) = NUMERO Then
        FnumInt$ = String$(CANTDECIMALES, "0")
        FormatearNumero = Format(NUMERO, "0." & FnumInt$)
    Else
        fnumDecimal$ = String$(CANTDECIMALES, "0")
        fnumDecimalNumeral$ = String$(CANTDECIMALES, "#")
        FormatearNumero = Format(NUMERO, "0." & fnumDecimal$ & fnumDecimalNumeral$)
    End If
End Function

Private Sub Label11_Change()
   MSF2.Clear
End Sub

Private Sub Label18_DblClick()
    Unload VENTABNEW

    Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Productos Terminados Solicitados"
    
    VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
    VENTANA.CampEditables = "1;3"
    VENTANA.Inicializar = 1
    '
    Call CopiarMsfAOtroMsf(Me.MSF1, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       StringCodint$ = ""
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
       For i& = 1 To VENTABNEW.MSF.Rows - 1
          If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)) <> "" Then 'campo de la descripcion del codigo
            CODIX$ = TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1))
            CI1% = CODINT%(CODIX$)
            If i& = 1 Then
               StringCodint$ = StringCodint$ + TRIM$(Str$(CI1%))
            Else
               StringCodint$ = StringCodint$ + "," + TRIM$(Str$(CI1%))
            End If
          End If
       Next i&
       Label11 = StringCodint$
    End If
    
    Set VENTANA = Nothing
    Unload VENTABNEW

End Sub

Private Sub List1_Click()
    CONTA% = 0
    lblCanped = "0"
    For U& = 1 To LIST1.ListCount
     If LIST1.Selected(U& - 1) = True Then
        CONTA% = CONTA% + 1
     End If
    Next U&
    lblCanped = TRIM$(CStr(CONTA%))
End Sub

Private Sub mnuComprasPenProv_Click()
   EBOL$ = OBJPLA$("SELPROVE", "")
   NPRO% = CVI(EBOL$)

   Screen.MousePointer = 11
   
    
    Campos$ = "Numero/f10"               '1
    Campos$ = Campos$ + ";F.Emis/D8"     '2
    Campos$ = Campos$ + ";Cta.Prov/F7"   '3
    Campos$ = Campos$ + ";Proveedor/A24" '4
    Campos$ = Campos$ + ";Codigo/A16"    '5
    Campos$ = Campos$ + ";Descripcion/A36" '6
    Campos$ = Campos$ + ";Cantidad/F10.1"  '7
    Campos$ = Campos$ + ";UMC/A5"         '8
    Campos$ = Campos$ + ";Mon/A5"         '9
    Campos$ = Campos$ + ";Prec.Compra/F12.4"  '10
    Campos$ = Campos$ + ";CantEntregada/F10.1" '11
    Campos$ = Campos$ + ";Pendiente/F10.1"     '12
    Campos$ = Campos$ + ";Imp-Pend./F12.4"     '13
    Campos$ = Campos$ + ";Acumulado/F14.4"     '14
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)

   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where Stat_Ocom < 3 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   If NPRO% > 0 Then SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO%
   SQLX$ = SQLX$ + " ORDER BY Npro_Ocom, Nume_Ocom, Cod_Exte, FEntre_Sol "
   
   Set OCOMTEMP = READSET(SQLX$)

   With OCOMTEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No se han encontrado ordenes de compras pendientes.")
        On Error GoTo 0
        GoTo 90
     End If
     On Error GoTo 0
     
     JJ& = 0: FIN& = OCOMTEMP.RecordCount
     Acumulado = 0

     Do While (.EOF = False)
         Call TRACE(24, JJ&, FIN&)
         
         MONEMI% = XVALO(!Mone_Emis)
         If MONEMI% <= 1 Then MONEX$ = "$"
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
         If CAPEN < 0.05 Then CAPEN = 0
         '
         JJ& = JJ& + 1
         GRILLAGRAL.GRID.Rows = JJ& + 1
         GRILLAGRAL.GRID.TextMatrix(JJ&, 1) = XVALO(!nume_ocom)
         GRILLAGRAL.GRID.TextMatrix(JJ&, 2) = FECHATEX$(CVINT(!Femi_Ocom))
         GRILLAGRAL.GRID.TextMatrix(JJ&, 3) = XVALO(!Npro_Ocom)
         GRILLAGRAL.GRID.TextMatrix(JJ&, 4) = rasocli$(-1 * XVALO(!Npro_Ocom))
         GRILLAGRAL.GRID.TextMatrix(JJ&, 5) = SAFETEXT$(!Cod_Exte)
         GRILLAGRAL.GRID.TextMatrix(JJ&, 6) = SAFETEXT$(!DESCRITEM)
         GRILLAGRAL.GRID.TextMatrix(JJ&, 7) = FORMATNUM$(XVALO(!Cant_Ocom), "F10.1")
         GRILLAGRAL.GRID.TextMatrix(JJ&, 8) = SAFETEXT(!Uni_Com)
         GRILLAGRAL.GRID.TextMatrix(JJ&, 9) = MONEX$
         GRILLAGRAL.GRID.TextMatrix(JJ&, 10) = FORMATNUM$(XVALO(!prun_neto), "F12.2")
         GRILLAGRAL.GRID.TextMatrix(JJ&, 11) = FORMATNUM$(XVALO(!Cant_Rec), "F12.2")
         GRILLAGRAL.GRID.TextMatrix(JJ&, 12) = FORMATNUM$(CAPEN, "F10.1")
         GRILLAGRAL.GRID.TextMatrix(JJ&, 13) = FORMATNUM$(XVALO(!prun_neto) * CAPEN, "F14.2")
           Acumulado = Acumulado + XVALO(GRILLAGRAL.GRID.TextMatrix(JJ&, 13))
         GRILLAGRAL.GRID.TextMatrix(JJ&, 14) = FORMATNUM$(Acumulado, "F14.2")
         .MoveNext

    Loop
   End With
   '
    GRILLAGRAL.Caption = "Ordenes de Compra Pendientes - Proveedor: " & rasocli$(-1 * NPRO%)
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Visible = False
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.NoOrdenarColumnas% = 1
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    Call numerargrilla(GRILLAGRAL.GRID, 0)
    Screen.MousePointer = 1
    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "CoPenPro", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If


90 Screen.MousePointer = 1

End Sub


Private Sub mnuIrPantGrab_Click()
'    If MSF1.Rows < 2 Then MSF1.Rows = 2
'    Call CopiarMsfAOtroMsf(Me.MSF1, PROPCOMGRAB.MSF1, 1)
'    If MSF3.Rows < 2 Then MSF3.Rows = 2
'    Call CopiarMsfAOtroMsf(Me.MSF3, PROPCOMGRAB.MSF3, 1)
    Campos$ = "Código/A18;Descripción/A32;Cantidad/f12.0;CI/A0"
    Call CARGA_ENCABEZADO(PROPCOMGRAB.MSF1, Campos$, Me)
    PROPCOMGRAB.MSF1.Rows = 2
    Call CARGA_ENCABEZADO(PROPCOMGRAB.MSF3, CamposMSF3$, Me)
    PROPCOMGRAB.MSF3.Rows = 2

    PROPCOMGRAB.LBLNUMERO = ""
    PROPCOMGRAB.LBLFECHA = ""
    PROPCOMGRAB.Show 1
End Sub

Private Sub mnuPrintSt_Click()
   Dim MSF As msFlexGrid
   Set MSF = Me.MSF2
   Campos$ = CamposMSF2$
   If Frame1.Visible = False Then Set MSF = Me.MSF3: Campos$ = CamposMSF3$
   '
   FILTROX$ = ""
   For REG& = 1 To MSF1.Rows - 1
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
   If REG& - 1 < 10 Then '16 SON EL TOTAL DE CASILLEROS DEL SQF
   For i& = REG& To 10
      FILTROX$ = FILTROX$ & ";" & "-"
   Next
   End If
   
   Call CARGA_TABLA_IMPRE(MSF, Campos$, "PROPCMPR", "Propuesta de Compra", "TABLA_IMPRE", RET_NAMECONS$, 1)
'   Call FINAL("?" & RET_NAMECONS$)

   FILTX$ = FILTROX$
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
'   Campos$ = Campos$ + ";Comp.Pend./F9.2"              '  12

   
    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Agregar Stock"
    
    VENTANA.CampEditables = "9;10;11"
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
Private Sub MSF1_DblClick()
       Call Label18_DblClick
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
      WebBrowser1.Navigate rutaArchivo
     Call T_Espera(1)
     WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    End If
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub MSFlexGrid1_DblClick()

End Sub


Private Sub MSF4_DblClick()
    i& = MSF4.MouseRow
    J& = MSF4.MouseCol
    If i& = 0 Or i& > MSF4.Rows Then Exit Sub
    '


    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    Call cmdEliminaItem(MSF4, REG&)
 
End Sub


Private Sub printHTML_Click()
Observa$ = InputBox$("Ingrese Observaciones, (64 caracteres)", "Observación", Now)
Observa$ = Left$(Observa$, 64)
Call GenerarHTMLDesdeFlexGrid(MSF3, "C:\FLEXOFT\Archivo.html", 1, Observa$)
   
End Sub


