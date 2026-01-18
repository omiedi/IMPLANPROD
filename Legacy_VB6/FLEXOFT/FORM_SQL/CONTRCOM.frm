VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CONTRCOM 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Trazabilidad"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   17070
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   620
      Left            =   16320
      Picture         =   "CONTRCOM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
      Top             =   240
      Width           =   650
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VISTA DETALLE"
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
      Left            =   11760
      TabIndex        =   9
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   9000
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACUMULAR X CÓDIGO"
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
      Left            =   13920
      TabIndex        =   8
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   9000
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Seleccionar Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   10440
      TabIndex        =   5
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   300
      Width           =   1170
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16170
      Begin VB.CommandButton Command8 
         Caption         =   "Depósitos Posibles con Saldo de Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   13320
         TabIndex        =   34
         ToolTipText     =   "Muestra los depósitos > 30 que tienen sus O.F. cerradas pero con Saldo de materiales"
         Top             =   1800
         Width           =   2250
      End
      Begin VB.CommandButton Command7 
         Caption         =   "prueba Grabar N.Serie"
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
         Left            =   13320
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.TextBox TXTNIVELES 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   10800
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text4 
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
         Height          =   300
         Left            =   7695
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   " "
         Top             =   1560
         Width           =   540
      End
      Begin VB.CommandButton BOTDEPO 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8220
         Picture         =   "CONTRCOM.frx":014A
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTDEPOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   3375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Grabar N.Serie"
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
         Left            =   3480
         TabIndex        =   24
         Top             =   1560
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   22
         ToolTipText     =   "Doble Clic Para ingresar N.Serie"
         Top             =   1560
         Width           =   1890
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   1890
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3250
         TabIndex        =   13
         ToolTipText     =   "Ordenes de Fabricación Abiertas"
         Top             =   480
         Width           =   175
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1380
         LinkTimeout     =   0
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox TXTPENDIENTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3450
         TabIndex        =   10
         ToolTipText     =   "Ordenes de Fabricación (Todas)"
         Top             =   480
         Width           =   175
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   11640
         TabIndex        =   6
         Top             =   160
         Width           =   4095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Niveles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9960
         TabIndex        =   32
         Top             =   1605
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito  a Transferir:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   6480
         TabIndex        =   29
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7695
         TabIndex        =   28
         Top             =   1905
         Width           =   4965
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N°.Serie: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3630
         TabIndex        =   21
         Top             =   480
         Width           =   4770
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   20
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   19
         Top             =   915
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1380
         TabIndex        =   18
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3630
         TabIndex        =   17
         Top             =   840
         Width           =   4770
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.O.F..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   315
         TabIndex        =   16
         Top             =   1200
         Width           =   1005
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   " Pendiente O.F.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5550
         TabIndex        =   15
         Top             =   1200
         Width           =   1605
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CONTRCOM.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONTRCOM.frx":0688
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6315
      Left            =   120
      TabIndex        =   3
      Top             =   2610
      Width           =   16185
      _ExtentX        =   28549
      _ExtentY        =   11139
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
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   4515
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   7964
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421504
      BackColorSel    =   4210752
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
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "CONTRCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CamposOriginal$

Sub CargarDatosControSerie(NORFA$, NUSERIE$)
      MSF.Rows = 1
      CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND (CANTINGRE <= 0 AND CANTSALID <= 0)  " ' LOS APLICA POR VALOBADA CON UN SUM
      SQLX$ = "SELECT * FROM CONTROSERIE WITH(NOLOCK) WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY ORDEN "
      Set RS = READSET(SQLX$)
      With RS
      Do While Not .EOF
         REG& = REG& + 1
         CI1% = XVALO(!cod_inte)
         MSF.Rows = REG& + 1
         MSF.TextMatrix(REG&, 1) = CI1%
         espaciosEnBlanco% = XVALO(!EspIdentado)
         SP$ = Space$(espaciosEnBlanco%)
         MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Cod_Exte)   'mantiene el identado
         MSF.TextMatrix(REG&, 3) = SP$ & SAFETEXT$(!Descrip)
         MSF.TextMatrix(REG&, 4) = Unimed$(CI1%)
         MSF.TextMatrix(REG&, 5) = XVALO(!usobruto)
         ID = XVALO(!CONTROSERIE_ID)
         CONDI2$ = "IdCodAsignado = " & ID 'TRAIGO MEDIANTE SUM LAS CANTIDADES INGRESADAS QUE MACHEAN CON LA FILA ORIGINAL
         CantidadAsignada = XVALO(VALOBADA("CONTROSERIE", "Sum(CANTINGRE - CANTSALID)", CONDI2$, "NOMESS"))
         MSF.TextMatrix(REG&, 6) = CantidadAsignada
         MSF.TextMatrix(REG&, 9) = ID
        .MoveNext
      Loop
      End With
      RS.Close: Set RS = Nothing
End Sub

Private Sub BOTSEL_Click(Index As Integer)

End Sub

Sub CrearDepositoSerie(NUSERIE$, NORFA$, OK%)
'CRITERIO DE GENERACION DE N. SERIE
'DESDE AQUI SIEMPRE GENERA DEPOSITOS >= 30 Y MENORES A 58
'SI ESTAN LIBRES GENERA EN ORDEN CORRELATIVO, EN EL CASO QUE LA CORRELATIVIDAD HAGA UN SALTO EJ DEL 35 SALTE AL 39
'SIGUE ENUMERANDO A PARTIR DEL 39 ES DECIR EL SIGUIENTE SERA EL 40
'CUANDO LLEGUE A 58 VA A BUSCAR UN HUECO ENTRE LOS EXISTENTES, SIGUIENDO EL EJEMPLO CREARIA EL 36 LUEGO 37 Y 38
'CUANDO NO ENCUENTRE HUECOS ENTRE EL 30 Y EL 58, TOMA EL DE LA DEPOSITO DE LAS OF QUE ESTEN CERRADA
'EN CASO QUE NO HAYA OF CERRADAS NI HUECOS PRESENTA MENSAJE PARA QUE SE CIERREN O.F.
    OK% = 0

    NUSERIE$ = TRIM$(NUSERIE$)
    NORFA$ = TRIM$(NORFA$)
    
    Dim PosicionRegistroSegunDeposito&(64)
    Erase PosicionRegistroSegunDeposito&
    
    Call TRALOCAL("TADEPOSI", "")
    For U& = 1 To ULTREG&("TADEPOSI")
      XX$ = REGLEIDO$("!TADEPOSI", U&)
      If ECOARCH$("PROVED1", Mid$(XX$, 61, 2)) = "" Then
         Call REPLA(XX$, MKI$(0), 61, 2)
         Call GRAREG("!TADEPOSI", XX$, U&)
      End If
    Next U&
    Call CIERRARCH("!TADEPOSI")
'    'si hay mas de 60 registros de deposito no hay mas capacidad
'    If u& > 60 Then
'       Call COMUNI("Sin Capacidad Para Crear Más Registros")
'       GoTo 99
'    End If

'_______________________________________________________
    valorMaximo% = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        XX$ = REGLEIDO$("!TADEPOSI", U&)
        NseriedLeido$ = Mid$(XX$, 2, 8)
        If TRIM$(NseriedLeido$) = TRIM$(NUSERIE$) Then
           Call COMUNI("El N.Serie ya se Ha Utilizado")
           OK% = -1
           GoTo 99
        End If
        OfLeida$ = Mid$(XX$, 10, 12)
        If TRIM$(OfLeida$) = TRIM$(NORFA$) Then
           Call COMUNI("La Orden de Fabricación ya Posee N.Serie")
           OK% = -1
           GoTo 99
        End If
        '
        
        'cargo un vector con la posicion del registro segun el numero de deposito
        NroDepositoLeido% = Asc(Left$(XX$, 1))
        PosicionRegistroSegunDeposito&(NroDepositoLeido%) = U&
        
        'me quedo con el nroDeposito maximo
        If NroDepositoLeido% > valorMaximo% Then
          valorMaximo% = NroDepositoLeido%
        End If
        '
        'me quedo con el primer deposerie que encuentre con la o.f cerrada, por si lo tengo que usar para reemplazar
        If NroDepositoLeido% >= 30 Then
          XZ = AgregarValorACadena(NroDepositoLeido%, numerosDepoTodosMayoresA30$) 'guardo todos los depo > 30 por si busco huecos
             ST% = STATORFA%(OfLeida$)
             If ST% = 3 Then
               PrimerOFabricacionCerrada$ = OfLeida$
               NSerieAReemplazar$ = NseriedLeido$
               NDepoAReemplazable% = NroDepositoLeido%
               'genero una cadena con los Nros de  deposerie reemplazables, por si luego tengo que usar el primer hueco
               XR = AgregarValorACadena(NDepoAReemplazable%, numerosDepoReemplazable$)
             End If
        End If
      
    Next U&
    
    'AGREGADO 30_10_2025
    'HAGO UNA LLAMADA PARA VALIDAR QUE LOS numerosDepoReemplazable$ NO TENGAN SALDOS DE STOCK
    'ESTO ME RETORNARIA UNA CADENA CON SOLO LOS QUE NO TIENE STOCK
    'SI TRAE DEPOSITOS DEBE SEGUIR NORMALMENTE
    'VALIDAR QUE ESE N.SERIE NO EXISTA COMO DEPOSITO

    DepoReemplazableSinSaldoStock$ = DepositosPosibles(numerosDepoReemplazable$)
    If TRIM$(DepoReemplazableSinSaldoStock$) <> "" Then
       numerosDepoReemplazable$ = DepoReemplazableSinSaldoStock$
    '   MsgBox numerosDepoReemplazable$
    Else
         'SI NO TRAE DEPOSITOS PRESENTA UN MENSAJE INDICANDO QUE LOS DEPOSITOS POSIBLES DE USAR POSEEN SALDOS "Ver Información/Cancelar"
          OPX% = COMALTER%("Depósitos Posibles Poseen Saldos de Stock\\Cancelar\Ver Depósitos Posibles Con Saldo")
          If OPX% <= 1 Then
            OK% = -1
            GoTo 99
          Else
            Call Command8_Click 'MUESTRA LA GRILLA CON LOS DEPOSITOS POSIBLES QUE POSEEN AUN SALDOS DE STOCK
            OK% = -1
            GoTo 99
          End If
    End If
    'FIN VALIDAR QUE ESE N.SERIE NO EXISTA COMO DEPOSITO
'_______________________________________________________

    J& = U& 'verificar si siempre al final con el   Next u& viene con uno mas.
    
    
    POSICIONJ& = 0
    Call CIERRARCH("!TADEPOSI")
    CODEPO% = 0
    If valorMaximo% < 58 Then
       CODEPO% = valorMaximo% + 1
    Else
       POSICIONJ& = 0
       CODEPO% = EncontrarPrimerSaltoNoCorrelativo%(numerosDepoTodosMayoresA30$)
       If CODEPO% < 1 Then
          CODEPO% = EncontrarValorMinimo(numerosDepoReemplazable$) 'de los nros de depositos reemplazables busca el valor  menor
          If CODEPO% > 0 Then
             POSICIONJ& = PosicionRegistroSegunDeposito&(CODEPO%)
          ElseIf CODEPO% < 1 Then
            Call COMUNI("No se Detectaron Depósitos Disponibles para Crear\Deberá Cerrar Alguna Orden de Fabricación\Que haya generado Depósito")
            OK% = -1
            GoTo 99
          End If
       End If
    End If
'    CODEPO% = 41
    If valorMaximo% < 30 Then CODEPO% = 30 'si el deposito a aplicar es menor a 30 inicializa con 30 (deberia ser solo la primera vez)
    
    JJ& = J& 'RETENGO EL VALOR DE J& PARA LOS SETULTREG
    If POSICIONJ& > 0 Then J& = POSICIONJ& 'si POSICIONJ& > 0 x q leyo la posicion para usar un deposito reemplazable
    
    X$ = REGBLAN$("TADEPOSI")
    Call REPLA(X$, Chr$(CODEPO%), 1, 1)
    Call REPLA(X$, NUSERIE$, 2, 8)
    Call REPLA(X$, NORFA$, 10, 12)
    
    Y$ = REGBLAN$("SUCURSAL")
    Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
    Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
    Call GRAREG("SUCURSAL", Y$, J&)
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Left$(X$, 21), 1, 21)
    Call GRAREG("ECODEP", Y$, J&)
    Call GRAREG("TADEPOSI", X$, J&)

    Call SETULTREG("SUCURSAL", JJ&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", JJ&)
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    JJ& = JJ& + 1
    Call GRAREG("ECODEP", Y$, JJ&)
    Call SETULTREG("ECODEP", JJ&)
    Call CIERRARCH("ECODEP")
99  Call CIERRARCH("!TADEPOSI")
    Call CIERRARCH("TADEPOSI")

    Call FRESHECHO("TADEPOSI")
    Call T_Espera(1 / 4)
    Text4 = CODEPO% 'ASIGNO EL VALOR DEL NRO DE DEPOSITO *** OJO DESPUES LO LEO CUANDO GRABO EN SQL
'    J& = 0
''______________________________________________________
'    'VALIDA SI LA OF ESTA CERRADA EN ESE CASO TOMA ESE N.SERIE
'    'STATORFA%
'    For u& = 1 To ULTREG&("!TADEPOSI")
'        X$ = REGLEIDO$("!TADEPOSI", u&)
'        CODEPO% = Asc(Left$(X$, 1))
'        If CODEPO% > 0 Then
'           If CODEPO% < 99 Then
'             If J& < 62 Then
'               J& = J& + 1
'               Y$ = REGBLAN$("SUCURSAL")
'               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
'               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
'               Call GRAREG("SUCURSAL", Y$, J&)
'               Y$ = REGBLAN$("ECODEP")
'               Call REPLA(Y$, Left$(X$, 21), 1, 21)
'               Call GRAREG("ECODEP", Y$, J&)
'               Call GRAREG("TADEPOSI", X$, J&)
'             End If
'           End If
'        End If
'    Next u&
'
'    J& = J& + 1
'    CODEPO% = 11
'    X$ = NUSERIE$
'    Call REPLA(Y$, NUSERIE$, 1, 63)
'    Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
'
'    Call REPLA(X$, Chr$(CODEPO%), 1, 1)
'    Call REPLA(X$, NUSERIE$, 2, 21)
'
'
'    Y$ = REGBLAN$("SUCURSAL")
'    Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
'    Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
'    Call GRAREG("SUCURSAL", Y$, J&)
'    Y$ = REGBLAN$("ECODEP")
'    Call REPLA(Y$, Left$(X$, 21), 1, 21)
'    Call GRAREG("ECODEP", Y$, J&)
'    Call GRAREG("TADEPOSI", X$, J&)
'
'
'
'    Call SETULTREG("SUCURSAL", J&)
'    Call CIERRARCH("SUCURSAL")
'    Call SETULTREG("TADEPOSI", J&)
'
'    '
'    Y$ = REGBLAN$("ECODEP")
'    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
'    J& = J& + 1
'    Call GRAREG("ECODEP", Y$, J&)
'    Call SETULTREG("ECODEP", J&)
'    Call CIERRARCH("ECODEP")
'99  Call CIERRARCH("!TADEPOSI")
'    Call CIERRARCH("TADEPOSI")
    End Sub


Function EncontrarValorMinimo(Cadena As String) As Integer
    Dim numeros() As String
    Dim i As Integer
    Dim valorMinimo As Integer
    If TRIM$(Cadena) = "" Then EncontrarValorMinimo = 0: Exit Function
    ' Dividir la cadena en un array de números
    numeros = Split(Cadena, ",")
    
    ' Inicializar el valor mínimo con el primer número de la cadena
    valorMinimo = CInt(numeros(0))
    
    ' Encontrar el valor mínimo en la cadena
    For i = 1 To UBound(numeros)
        If CInt(numeros(i)) < valorMinimo Then
            valorMinimo = CInt(numeros(i))
        End If
    Next i
    
    ' Devolver el valor mínimo encontrado
    EncontrarValorMinimo = valorMinimo
End Function

Function AgregarValorACadena(valor As Integer, ByRef Cadena As String)
    If Cadena = "" Then
        Cadena = CStr(valor)  ' Agregar el primer valor a la cadena
    Else
        Cadena = Cadena & "," & CStr(valor)  ' Agregar un valor adicional separado por  coma
    End If
End Function

 
Function EncontrarPrimerSaltoNoCorrelativo(Cadena As String) As Integer
    Dim numeros() As String
    Dim i As Integer
    Dim VALORES() As Integer
    Dim valorMinimo As Integer
    Dim valorMaximo As Integer
    
    ' Dividir la cadena en un array de números
    numeros = Split(Cadena, ",")
    
    ' Convertir los números a enteros y guardar en un array
    ReDim VALORES(UBound(numeros))
    For i = 0 To UBound(numeros)
        VALORES(i) = CInt(numeros(i))
    Next i
    
    ' Encontrar el valor mínimo y máximo en la cadena
    valorMinimo = Minimo(VALORES)
    valorMaximo = Maximo(VALORES)
    
    ' Buscar el primer valor faltante dentro del rango
    For i = valorMinimo To valorMaximo
        If Not ContieneValor(VALORES, i) Then
            EncontrarPrimerSaltoNoCorrelativo = i
            Exit Function
        End If
    Next i
    
    ' Si no falta ningún valor dentro del rango, devolver 0
    EncontrarPrimerSaltoNoCorrelativo = 0
End Function

Function Minimo(VALORES() As Integer) As Integer
    Dim min As Integer
    min = VALORES(0)
    For i = 1 To UBound(VALORES)
        If VALORES(i) < min Then
            min = VALORES(i)
        End If
    Next i
    Minimo = min
End Function

Function Maximo(VALORES() As Integer) As Integer
    Dim max As Integer
    max = VALORES(0)
    For i = 1 To UBound(VALORES)
        If VALORES(i) > max Then
            max = VALORES(i)
        End If
    Next i
    Maximo = max
End Function

Function ContieneValor(VALORES() As Integer, valor As Integer) As Boolean
    For i = 0 To UBound(VALORES)
        If VALORES(i) = valor Then
            ContieneValor = True
            Exit Function
        End If
    Next i
    ContieneValor = False
End Function



Sub GeneraIngresoDeMaterialAControserie(CI1%, NORFA$, NUSERIE$, IdCodAsignado, DEPOINGRE%, CANTI)
        CANTIINGRESADA = 0: CANTISALIDA = 0
        If CANTI > 0 Then CANTIINGRESADA = CANTI
        If CANTI < 0 Then CANTISALIDA = Abs(CANTI)
        Codigo$ = CODEXT$(CI1%)
        DESXRIP$ = DESCRIT0$(CI1%)
        SQLX$ = "INSERT INTO CONTROSERIE "
        SQLX$ = SQLX$ + " (COD_INTE,COD_EXTE,CODIEMPR,IDSUBOR,NUMEROSERIE,IdCodAsignado,DepoMovi,DESCRIP,STATUS,FEREMOVI"
        SQLX$ = SQLX$ + " ,CANTINGRE"
        SQLX$ = SQLX$ + " ,CANTSALID)"
        SQLX$ = SQLX$ + " VALUES(" & CI1%
        SQLX$ = SQLX$ + " ,'" & Codigo$ & "'"
        SQLX$ = SQLX$ + " ," & CodiEmp%
        SQLX$ = SQLX$ + " ,'" & NORFA$ & "'"
        SQLX$ = SQLX$ + " ,'" & NUSERIE$ & "'"
        SQLX$ = SQLX$ + " ," & IdCodAsignado
        SQLX$ = SQLX$ + " ," & DEPOINGRE%
        SQLX$ = SQLX$ + " ,'" & Left$(DESXRIP$, 64) & "'"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,'" & Now & "'"
        
        SQLX$ = SQLX$ + " ," & CANTIINGRESADA
        SQLX$ = SQLX$ + " ," & CANTISALIDA
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
End Sub

Private Sub BOTDEPO_Click()
      Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
          Text4 = TRIM$(VALACT1$("TADEPOSI"))
        End If
End Sub


Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   SELDEPOS.Show 1
   STRINGDEPOS$ = ""
   List1.Clear
   
   For U& = 1 To SELDEPOS.LISNUM.ListCount
       If SELDEPOS.LISNUM.Selected(U& - 1) = True Then
        TXTX$ = SELDEPOS.LISNUM.List(U& - 1)
        List1.AddItem TXTX$
        NDEP$ = TRIM$(Str$(Mid$(TXTX$, 1, 6)))
        STRINGDEPOS$ = STRINGDEPOS$ + NDEP$ + ","
      End If
     Next U&
     If TRIM$(STRINGDEPOS$) <> "" Then
        Me.TXTDEPOS = Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1) 'para quitar la coma
     Else
        TXTDEPOS = ""
     End If

   Command15.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Text3.Enabled = False
   Command15.Enabled = False
   List1.Enabled = False
   MSF2.Rows = 1
   Dim VectCodigoAcumulaCantSimulada()
   CantInicial% = 0
   ReDim Preserve VectCodigoAcumulaCantSimulada(CantInicial%)
   '
   'ACA QUEDARIAN CARGADOS SOLO UNA VEZ CADA CODIGO Y APARTE ACUMULO EN UN VECTOR LA CANTIDAD ACUMULADA DE LA COL 6
   For i& = 1 To MSF.Rows - 1 'RECORRO
    CI1% = XVALO(MSF.TextMatrix(i&, 1))
        If CI1% >= UBound(VectCodigoAcumulaCantSimulada) Then
           ReDim Preserve VectCodigoAcumulaCantSimulada(CI1%)
        End If
       If VectCodigoAcumulaCantSimulada(CI1%) > 0 Then 'ES QUE A ESTE CODIGO  YA LO LEYO
          'ACUMULA LA CANTIDAD SIMULADA
          VectCodigoAcumulaCantSimulada(CI1%) = VectCodigoAcumulaCantSimulada(CI1%) + XVALO(MSF.TextMatrix(i&, 6))
       Else
          REG& = REG& + 1
          MSF2.Rows = REG&
          VectCodigoAcumulaCantSimulada(CI1%) = XVALO(MSF.TextMatrix(i&, 6))
          Call AGREGAR_FILA_OTRO_MSF(MSF, MSF2, REG&)
       End If
       
   
   Next i&
   '
   'ACA APLICO LA CANTIDAD ACUMULADA
   For i& = 1 To MSF2.Rows - 1 'RECORRO
       CI1% = XVALO(MSF2.TextMatrix(i&, 1))
      MSF2.TextMatrix(i&, 6) = FORMATNUM$(VectCodigoAcumulaCantSimulada(CI1%), "F12.3") 'LE APLICA EN LA CANTIDAD SIMULADA EL ACUMULADO QUE CARGO EN EL FOR ANTERIOR
   Next i&
   MSF.Visible = False
   MSF.ZOrder 1
   MSF2.Visible = True
   MSF2.ZOrder 0
   
99 Command3.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Command2.Enabled = True
   Text3.Enabled = True
   Command15.Enabled = True
   List1.Enabled = True
   MSF.Visible = True: MSF.ZOrder 0
   TXTDEPOS.Enabled = True
End Sub

Private Sub Command5_Click()
    Label4 = OrdenFabricacionAbiertas$

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Screen.MousePointer = 11
    NORFA$ = TRIM$(Label4)
    If NORFA$ = "" Then
       Call COMUNI("Falta Seleccionar Orden de Fabricación")
       GoTo 99
    End If
    '
    NUSERIE$ = TRIM$(Text1)
    If NUSERIE$ = "" Then
       Call COMUNI("Falta Seleccionar N° de Serie")
       GoTo 99
    End If
    '
    CI1% = CODINT%(Text10)
    If CI1% < 1 Then
       Call COMUNI("Falta Código")
       GoTo 99
    End If
    
    Call CrearDepositoSerie(NUSERIE$, NORFA$, OK%)
    If OK% < 0 Then GoTo 99
    Call FRESHECHO("TADEPOSI")
    REG& = 0
    
    'EL CRITERIO ES QUE SE GRABE SOLO UNA VEZ CUANDO SE LE ASIGNA EL N.SERIE, LUEGO SE AGREGARAN REGISTROS, DE INGRESO O SALIDA.
    'SI YA EXISTE CON EL N.SERIE TRAE TODA LA INFO DE LA TABLA CONTROSERIE
    '- PARA QUE TRAIGA LOS REGISTROS ORIGINALES Y NO LOS QUE TIENEN INGRESOS
     CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "'"  ' LOS APLICA POR VALOBADA CON UN SUM
     If CantRegistros("CONTROSERIE", CONDI$, "NOMESS") > 0 Then GoTo 99
     
     Dim RS As ADODB.Recordset
     Set RS = New ADODB.Recordset
     SQLX$ = "SELECT * FROM CONTROSERIE WHERE " & CONDI$
     RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     '
     With RS
       For REG& = 1 To MSF.Rows - 1
        .AddNew
        !CODIEMPR = CodiEmp%
        !ORDEN = REG&
        !cod_inte = XVALO(MSF.TextMatrix(REG&, 1))
        espaciosEnBlanco = XVALO(Len(MSF.TextMatrix(REG&, 3))) - XVALO(Len(LTrim(MSF.TextMatrix(REG&, 3)))): If espaciosEnBlanco < 0 Then espaciosEnBlanco = 0
        !Cod_Exte = TRIM$(MSF.TextMatrix(REG&, 2))
        !Descrip = Left$(TRIM$(MSF.TextMatrix(REG&, 3)), 64)
        !usobruto = XVALO(MSF.TextMatrix(REG&, 5))
        !NUMEROSERIE = NUSERIE$
        !idsubor = NORFA$
        !CANTINGRE = 0
        !CANTSALID = 0
        !idenlote = ""
        !MARBORRA = 0
        !FeReMovi = Now
        !Status = 0
        !DEPOMOVI = 0
        !EspIdentado = espaciosEnBlanco
        !DEPOMOVI = XVALO(Text4)
        
        '!IDSUBOR2 = ""
        .Update
       Next REG&
    End With
    RS.Close: Set RS = Nothing
    
    'CREAR DEPOSITO.
    'RECORRER LOS DELPOSITOS MAYORES A 30 VERFICAR QUE SEGUN OF ESTE CERRADA, Y QUE PARA ESA OF NO TENGA MATERIAL.
    'DONDE ENCUENTRE UNO LIBRE VER SI ESTA CREADO Y CAMBIARLE LA DESCX
    Call COMUNI("Grabación Realizada")
    Command6.Visible = False

    Screen.MousePointer = 1
   
99 Command6.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command7_Click()
Command7.Enabled = False
    NORFA$ = TRIM$(InputBox$("Ingrese N° ORDEN FABRICACION", "Número de O.F.", NORFA$))

    NUSERIE$ = TRIM$(InputBox$("Ingrese N° Serie", "Número de Serie", NUSERIE$))

    Call CrearDepositoSeriePrueba(NUSERIE$, NORFA$, OK%)
    
    If OK% < 0 Then GoTo 99
99  Command7.Enabled = True
End Sub


Sub CrearDepositoSeriePrueba(NUSERIE$, NORFA$, OK%)
'CRITERIO DE GENERACION DE N. SERIE
'DESDE AQUI SIEMPRE GENERA DEPOSITOS >= 30 Y MENORES A 58
'SI ESTAN LIBRES GENERA EN ORDEN CORRELATIVO, EN EL CASO QUE LA CORRELATIVIDAD HAGA UN SALTO EJ DEL 35 SALTE AL 39
'SIGUE ENUMERANDO A PARTIR DEL 39 ES DECIR EL SIGUIENTE SERA EL 40
'CUANDO LLEGUE A 58 VA A BUSCAR UN HUECO ENTRE LOS EXISTENTES, SIGUIENDO EL EJEMPLO CREARIA EL 36 LUEGO 37 Y 38
'CUANDO NO ENCUENTRE HUECOS ENTRE EL 30 Y EL 58, TOMA EL DE LA DEPOSITO DE LAS OF QUE ESTEN CERRADA
'EN CASO QUE NO HAYA OF CERRADAS NI HUECOS PRESENTA MENSAJE PARA QUE SE CIERREN O.F.
    OK% = 0

    NUSERIE$ = TRIM$(NUSERIE$)
    NORFA$ = TRIM$(NORFA$)
    
    Dim PosicionRegistroSegunDeposito&(64)
    Erase PosicionRegistroSegunDeposito&
    
    Call TRALOCAL("TADEPOSI", "")
    For U& = 1 To ULTREG&("TADEPOSI")
      XX$ = REGLEIDO$("!TADEPOSI", U&)
      If ECOARCH$("PROVED1", Mid$(XX$, 61, 2)) = "" Then
         Call REPLA(XX$, MKI$(0), 61, 2)
         Call GRAREG("!TADEPOSI", XX$, U&)
      End If
    Next U&
    Call CIERRARCH("!TADEPOSI")
'    'si hay mas de 60 registros de deposito no hay mas capacidad
'    If u& > 60 Then
'       Call COMUNI("Sin Capacidad Para Crear Más Registros")
'       GoTo 99
'    End If

'_______________________________________________________
'VALIDAR QUE ESE N.SERIE NO EXISTA COMO DEPOSITO
    valorMaximo% = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        XX$ = REGLEIDO$("!TADEPOSI", U&)
        NseriedLeido$ = Mid$(XX$, 2, 8)
        If TRIM$(NseriedLeido$) = TRIM$(NUSERIE$) Then
           Call COMUNI("El N.Serie ya se Ha Utilizado")
           OK% = -1
           GoTo 99
        End If
        OfLeida$ = Mid$(XX$, 10, 12)
        If TRIM$(OfLeida$) = TRIM$(NORFA$) Then
           Call COMUNI("La Orden de Fabricación ya Posee N.Serie")
           OK% = -1
           GoTo 99
        End If
        '
        'cargo un vector con la posicion del registro segun el numero de deposito
        NroDepositoLeido% = Asc(Left$(XX$, 1))
        PosicionRegistroSegunDeposito&(NroDepositoLeido%) = U&
        
        'me quedo con el nroDeposito maximo
        If NroDepositoLeido% > valorMaximo% Then
          valorMaximo% = NroDepositoLeido%
        End If
        '
        'me quedo con el primer deposerie que encuentre con la o.f cerrada, por si lo tengo que usar para reemplazar
        If NroDepositoLeido% >= 30 Then
          XZ = AgregarValorACadena(NroDepositoLeido%, numerosDepoTodosMayoresA30$) 'guardo todos los depo > 30 por si busco huecos
             ST% = STATORFA%(OfLeida$)
             If ST% = 3 Then
               PrimerOFabricacionCerrada$ = OfLeida$
               NSerieAReemplazar$ = NseriedLeido$
               NDepoAReemplazable% = NroDepositoLeido%
               'genero una cadena con los Nros de  deposerie reemplazables, por si luego tengo que usar el primer hueco
               XR = AgregarValorACadena(NDepoAReemplazable%, numerosDepoReemplazable$)
             End If
        End If
      
    Next U&
MsgBox "XR = AgregarValorACadena " & XR
MsgBox "numerosDepoTodosMayoresA30$ " & numerosDepoTodosMayoresA30$
MsgBox "numerosDepoReemplazable$ " & numerosDepoReemplazable$

    'HAGO UNA LLAMADA PARA VALIDAR QUE LOS numerosDepoReemplazable$ NO TENGAN SALDOS DE STOCK
    'ESTO ME RETORNARIA UNA CADENA CON SOLO LOS QUE NO TIENE STOCK
    'SI TRAE DEPOSITOS DEBE SEGUIR NORMALMENTE
    DepoReemplazableSinSaldoStock$ = DepositosPosibles(numerosDepoReemplazable$)
    If TRIM$(DepoReemplazableSinSaldoStock$) <> "" Then
       numerosDepoReemplazable$ = DepoReemplazableSinSaldoStock$
    '   MsgBox numerosDepoReemplazable$
    Else
         'SI NO TRAE DEPOSITOS PRESENTA UN MENSAJE INDICANDO QUE LOS DEPOSITOS POSIBLES DE USAR POSEEN SALDOS "Ver Información/Cancelar"
          OPX% = COMALTER%("Depósitos Posibles Poseen Saldos de Stock\\Cancelar\Ver Depósitos Posibles Con Saldo")
          If OPX% <= 1 Then Exit Sub
          Call DepositosConSaldoArcon(numerosDepoReemplazable$)
          'GRILLA GRAL DE LOS DEPOSITOS TODOS EL USUARIO PODRA FILTRARLOS.
    End If



    J& = U& 'verificar si siempre al final con el   Next u& viene con uno mas.
    POSICIONJ& = 0
    Call CIERRARCH("!TADEPOSI")
    CODEPO% = 0
    If valorMaximo% < 58 Then
       CODEPO% = valorMaximo% + 1
MsgBox "1.-CODEPO% " & CODEPO%
    Else
       POSICIONJ& = 0
       CODEPO% = EncontrarPrimerSaltoNoCorrelativo%(numerosDepoTodosMayoresA30$)
       If CODEPO% < 1 Then
MsgBox "2.-CODEPO% " & CODEPO%
          CODEPO% = EncontrarValorMinimo(numerosDepoReemplazable$) 'de los nros de depositos reemplazables busca el valor  menor
          If CODEPO% > 0 Then
MsgBox "3.-CODEPO% " & CODEPO%
          
             POSICIONJ& = PosicionRegistroSegunDeposito&(CODEPO%)
          ElseIf CODEPO% < 1 Then
            Call COMUNI("No se Detectaron Depósitos Disponibles para Crear\Deberá Cerrar Alguna Orden de Fabricación\Que haya generado Depósito")
            OK% = -1
            GoTo 99
          End If
       End If
    End If
'    CODEPO% = 41
    If valorMaximo% < 30 Then CODEPO% = 30 'si el deposito a aplicar es menor a 30 inicializa con 30 (deberia ser solo la primera vez)
MsgBox "4.-CODEPO% " & CODEPO%
    
'    JJ& = J& 'RETENGO EL VALOR DE J& PARA LOS SETULTREG
'    If POSICIONJ& > 0 Then J& = POSICIONJ& 'si POSICIONJ& > 0 x q leyo la posicion para usar un deposito reemplazable
'
'    X$ = REGBLAN$("TADEPOSI")
'    Call REPLA(X$, Chr$(CODEPO%), 1, 1)
'    Call REPLA(X$, NUSERIE$, 2, 8)
'    Call REPLA(X$, NORFA$, 10, 12)
'
'    Y$ = REGBLAN$("SUCURSAL")
'    Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
'    Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
'    Call GRAREG("SUCURSAL", Y$, J&)
'    Y$ = REGBLAN$("ECODEP")
'    Call REPLA(Y$, Left$(X$, 21), 1, 21)
'    Call GRAREG("ECODEP", Y$, J&)
'    Call GRAREG("TADEPOSI", X$, J&)
'
'    Call SETULTREG("SUCURSAL", JJ&)
'    Call CIERRARCH("SUCURSAL")
'    Call SETULTREG("TADEPOSI", JJ&)
'    '
'    Y$ = REGBLAN$("ECODEP")
'    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
'    JJ& = JJ& + 1
'    Call GRAREG("ECODEP", Y$, JJ&)
'    Call SETULTREG("ECODEP", JJ&)
'    Call CIERRARCH("ECODEP")
'99  Call CIERRARCH("!TADEPOSI")
'    Call CIERRARCH("TADEPOSI")
'
'    Call FRESHECHO("TADEPOSI")
'    Call T_Espera(1 / 4)
'    Text4 = CODEPO% 'ASIGNO EL VALOR DEL NRO DE DEPOSITO *** OJO DESPUES LO LEO CUANDO GRABO EN SQL
'    J& = 0
99
    End Sub
Public Function DepositosPosibles(ByVal cadena_de_numeros As String, Optional ByVal modoInverso As Integer = 0) As String
    Dim arrDepositos() As String
    Dim i As Integer
    Dim DEPO As String
    Dim cadenaResultado As String
    Dim tieneRegistros As Boolean
    
    ' Dividir la cadena recibida en un array
    arrDepositos = Split(cadena_de_numeros, ",")
    
    cadenaResultado = ""
    
    ' Recorrer cada depósito
    For i = LBound(arrDepositos) To UBound(arrDepositos)
        DEPO = TRIM(arrDepositos(i))
        DEPONUM% = XVALO(DEPO)
        ' Verificar si el depósito tiene registros
        tieneRegistros = ExisteRegistroEnDeposito(DEPONUM%)
        'If (modoInverso > 0 entra por este if si tiene registros con saldo
        If (modoInverso > 0 And tieneRegistros) Then
             If cadenaResultado <> "" Then cadenaResultado = cadenaResultado & ","
             cadenaResultado = cadenaResultado & DEPO
        Else 'si no viene aca y  Si NO tiene registros, agregar a la cadena resultado
            If Not tieneRegistros Then
                If cadenaResultado <> "" Then cadenaResultado = cadenaResultado & ","
                cadenaResultado = cadenaResultado & DEPO
            End If
        End If
    Next i
    
    ' Devolver cadena resultado (vacía si no hay depósitos posibles)
    DepositosPosibles = cadenaResultado
End Function

Private Function ExisteRegistroEnDeposito(nroDeposito%) As Boolean
    
    SQLX$ = "SELECT  SUM(ISNULL(CANTINGRE, 0) - ISNULL(CANTSALID, 0))  AS RESULTADO,DEPOMOVI,Cod_Exte  "
    SQLX$ = SQLX$ + " From Movisto WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where DEPOMOVI = " & nroDeposito%
    SQLX$ = SQLX$ + " GROUP BY DEPOMOVI, COD_EXTE"
    SQLX$ = SQLX$ + " Having Abs(SUM(CANTINGRE - CANTSALID)) > 0.5"
    
    
    Set RS = FLEXCONN.Execute(SQLX$)
    
    ExisteRegistroEnDeposito = Not RS.EOF ' True si hay registros
    
    RS.Close
    Set RS = Nothing
End Function

Sub Command8_Click()
    Command8.Enabled = False
    'ESTE CODIGO VIENE DE LA ASIGNACION DE DEPOSITO POR NRO DE SERIE
    'recargo los depositos en tadeposi local
    Call TRALOCAL("TADEPOSI", "")
    For U& = 1 To ULTREG&("TADEPOSI")
      XX$ = REGLEIDO$("!TADEPOSI", U&)
      If ECOARCH$("PROVED1", Mid$(XX$, 61, 2)) = "" Then
         Call REPLA(XX$, MKI$(0), 61, 2)
         Call GRAREG("!TADEPOSI", XX$, U&)
      End If
    Next U&
    Call CIERRARCH("!TADEPOSI")
    '

    valorMaximo% = 0

    For U& = 1 To ULTREG&("!TADEPOSI")
        XX$ = REGLEIDO$("!TADEPOSI", U&)
        NroDepositoLeido% = Asc(Left$(XX$, 1))
        'me quedo con el nroDeposito maximo
        If NroDepositoLeido% > valorMaximo% Then
          valorMaximo% = NroDepositoLeido%
        End If
        OfLeida$ = Mid$(XX$, 10, 12)

        '
        'me quedo con el primer deposerie que encuentre con la o.f cerrada, por si lo tengo que usar para reemplazar
        If NroDepositoLeido% >= 30 Then
             XZ = AgregarValorACadena(NroDepositoLeido%, numerosDepoTodosMayoresA30$) 'guardo todos los depo > 30 por si busco huecos
             ST% = STATORFA%(OfLeida$)
             If ST% = 3 Then
               PrimerOFabricacionCerrada$ = OfLeida$
               NSerieAReemplazar$ = NseriedLeido$
               NDepoAReemplazable% = NroDepositoLeido%
               'genero una cadena con los Nros de  deposerie reemplazables, por si luego tengo que usar el primer hueco
               XR = AgregarValorACadena(NDepoAReemplazable%, numerosDepoReemplazable$)
             End If
        End If
    Next U&
    TraerRegistrosConSaldo% = 1 'asigno a una variable para que quede mas claro el valor 1
    DepoReemplazableConSaldoStock$ = DepositosPosibles(numerosDepoReemplazable$, TraerRegistrosConSaldo%)
    If TRIM$(DepoReemplazableConSaldoStock$) <> "" Then
       numerosDepoReemplazable$ = DepoReemplazableConSaldoStock$
       Call DepositosConSaldoArcon(numerosDepoReemplazable$)
    Else
       Call COMUNI("No se Detectaron Depósitos posibles con Saldo")
    End If

99  Command8.Enabled = True
End Sub
    Sub DepositosConSaldoArcon(CadenaDepositos$)
   'Stock Negativo considerando Lote indiv ventA
        Campos$ = "Codigo/A20" '1
        Campos$ = Campos$ + ";Descripción/A32" '2
        Campos$ = Campos$ + ";Saldo Stock/F16.1" '3
        Campos$ = Campos$ + ";Depósito/F12" '4
        J& = 0
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
        '
        SQLX$ = "SELECT SUM(COALESCE(CANTINGRE, 0) - COALESCE(CANTSALID, 0)) AS RESUL, DEPOMOVI, COD_INTE "
        SQLX$ = SQLX$ & "FROM MOVISTO WITH(NOLOCK) "
        SQLX$ = SQLX$ & "WHERE DEPOMOVI IN (" & CadenaDepositos$ & ") "
        SQLX$ = SQLX$ & "GROUP BY DEPOMOVI, COD_INTE  "
        SQLX$ = SQLX$ & "HAVING ABS(SUM(COALESCE(CANTINGRE, 0) - COALESCE(CANTSALID, 0))) > 0.5"
        SQLX$ = SQLX$ & " ORDER BY DEPOMOVI, COD_INTE"
        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
             
'             If SALDO < XVALO(!CAMPO3) Then para que muestre todos los de la lista de StMinNeu
                J& = J& + 1
                CI1% = XVALO(!cod_inte)
                GRILLAGRAL.GRID.Rows = J& + 1
                GRILLAGRAL.GRID.TextMatrix(J&, 1) = CODEXT$(CI1%)
                GRILLAGRAL.GRID.TextMatrix(J&, 2) = DESCRIT0$(CI1%)
                GRILLAGRAL.GRID.TextMatrix(J&, 3) = FORMATNUM$(XVALO(!RESUL), "F12.1")
                GRILLAGRAL.GRID.TextMatrix(J&, 4) = XVALO(!DEPOMOVI)
'             End If
            .MoveNext
          Loop
        End With
        RS.Close
        Set RS = Nothing
        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        TX$ = "Depósitos Posibles C/Items con Saldo "
        GRILLAGRAL.Caption = TX$
        GRILLAGRAL.SinEfectoClickEnGrilla = 1
        Screen.MousePointer = 1
        GRILLAGRAL.Show 1
        
        
        Dim MSF As msFlexGrid
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MENU1 > 0 Then
             If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
             Call CARGA_TABLA_IMPRE(MSF, Campos$, "DEPCSALD", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
             Call FINAL("?" & RET_NAMECONS$)
             GRILLAGRAL.MENU1 = 0
             Screen.MousePointer = 1
             Exit Sub
        End If

   
End Sub

Private Sub Form_Load()
    '
    CamposOriginal$ = "CI1/Aa0"                              '1
    CamposOriginal$ = CamposOriginal$ + ";CodXEelem/A15"    '2
    CamposOriginal$ = CamposOriginal$ + ";Descripción/A32"  '3
    CamposOriginal$ = CamposOriginal$ + ";U.M./A3"           '4
    CamposOriginal$ = CamposOriginal$ + ";Uso/F9.4"         '5
    CamposOriginal$ = CamposOriginal$ + ";Cant.Asig./F10.4" '6
    CamposOriginal$ = CamposOriginal$ + ";Dep.Salida/F10"   '7
    CamposOriginal$ = CamposOriginal$ + ";Fecha/D8"         '8
    CamposOriginal$ = CamposOriginal$ + ";ID/A10"           '9
    '
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   Me.MSF.Width = Me.Width - 1000
'   Me.MSF_2.Height = Me.Height - 1600
'   Frame1.Top = Me.Height - 1300
'   Frame2.Top = Me.Height - 1300
   On Error GoTo 0
End Sub

Private Sub Label4_Change()
   '
   Command6.Visible = True
   MSF.Rows = 1
   NORFA$ = TRIM$(Label4)
   CANTI = CantiProyectadaOF(NORFA$)
   If CANTI > 1 Then
      Call COMUNI("Orden de Fabricación Seleccionada Tiene más de 1 Item")
      Exit Sub
   End If
   CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
   NUSERIE$ = TRIM$(VALOBADA("CONTROSERIE", "NUMEROSERIE", CONDI$, "NOMESS"))
   If NUSERIE$ = "" Then Command6.Visible = True
   If NORFA$ = "" Then Exit Sub
   CI1% = CODIFAB%(NORFA$)
   Label2 = RESP_ZELE_VECT(2)
   Text10 = CODEXT$(CI1%)
   Label3 = DESCRIT0$(CI1%)
   CIFAB% = CODINT%(Text10)
   Label3 = DESCRIT$(CI1%)
   CODDX$ = Label4
   Text4 = ""
   Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$)
   Text14 = CANTI
   Call CARGA_ENCABEZADO(MSF, CamposOriginal$, CONTRCOM)
   Call CARGA_ENCABEZADO(MSF2, CamposOriginal$, CONTRCOM)
   If CANTI > 1 Then
      'call cargar la lista de las of con los n.serie.
   End If
   Text1 = ""
   If NUSERIE$ <> "" And NORFA$ <> "" Then
      Text1 = NUSERIE$
   Else
     TXTNIVELES = 100
     Call CargaEstructuraMultinivel(Me.MSF, CI1%, XVALO(TXTNIVELES), 0)
   End If
   
End Sub
Sub CargarGrilla(CI1%)
   '
   Screen.MousePointer = 11
   '
   cantidadSimulada = XVALO(Text3)
   If cantidadSimulada < 1 Then cantidadSimulada = 1
   COEFCA = 1: If PORMILLAR% = 1 Then COEFCA = 1000
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     RESUL = NIVE%(KKI%)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     SP1$ = String$(RESUL, ">")
     If RESUL > 1 And RESUL_ANT <> RESUL Then IDENTIFICADORDENIVEL = IDENTIFICADORDENIVEL & "." & TRIM$(CStr(RESUL))
     MSF.Rows = KKI%
     
     MSF.TextMatrix(KKI% - 1, 1) = CIK%(KKI%)
     MSF.TextMatrix(KKI% - 1, 2) = CODEXT$(CIK%(KKI%))
     MSF.TextMatrix(KKI% - 1, 3) = SP$ & SP$ & DESCRIT0$(CIK%(KKI%))
     MSF.TextMatrix(KKI% - 1, 4) = Unimed$(CIK%(KKI%))
     MSF.TextMatrix(KKI% - 1, 5) = ROUND(COEFCA * CANT(KKI%), 4)
     MSF.TextMatrix(KKI% - 1, 6) = ""
     MSF.TextMatrix(KKI% - 1, 7) = ""
     MSF.TextMatrix(KKI% - 1, 8) = ""
   Next KKI%
   Screen.MousePointer = 1
End Sub



Private Sub List1_Change()

End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Exit Sub
   REG& = MSF.MouseRow
   COL& = MSF.MouseCol
   If REG& = 0 Or REG& > MSF.Rows Then Exit Sub
   
   If Button = 2 Then
       '
       CI1% = XVALO(MSF.TextMatrix(REG&, 1))
       LOTE$ = ""
       NORFA$ = TRIM$(Label4)
       NUSERIE$ = TRIM$(Text1)
       ID = XVALO(MSF.TextMatrix(REG&, 9))
       MSF.row = REG&
       MSF.COL = COL&
       MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
       DEPO1$ = MSF.TextMatrix(0, COL&)
       CANTIX$ = MSF.TextMatrix(REG&, COL&)
       If XVALO(CANTIX$) <= 0 Then
          Call COMUNI("Cantidad Insuficiente Para Transferir")
          Exit Sub
       End If
       If XVALO(Text4) < 1 Then
         Call COMUNI("Falta Informar Depósito de Ingreso de Reserva del Material")
         Exit Sub
       End If
       
       CANTI = XVALO(InputBox$("Stock Actual en : " & DEPO1$ & "=" & CANTIX$, "Ingrese Cantidad a Transferir"))
       If CANTI <= 0 Then Exit Sub
       
       FCHA% = HOY(HO$)
       CMOV$ = "TR"

       DEPOSAL% = XVALO(Mid$(DEPO1$, 4))
       DEPOINGRE% = XVALO(Text4)

       NCOMP& = ProNroComprobante(CMOV$)
       DOPRI$ = "TR." + TRIM$(Str$(NCOMP&))
       DOSEC$ = NORFA$
      
       DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOSAL%)))
       REFE$ = "Tr. " & NUSERIE$ & " - " & NORFA$
       NCOMP& = ProNroComprobante(CMOV$)
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans

'       Call MOVISTOK(FCHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOSAL%, (-1) * CANTI)
'       Call T_Espera(1 / 10) 'PARA QUE HAYA DIFERENCIA ENTRE UNA GRABACION Y LA OTRA
       DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOINGRE%)))
'       REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TRIM$(NORFA$)
'       Call MOVISTOK(FCHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOINGRE%, CANTI)
       Call GeneraIngresoDeMaterialAControserie(CI1%, NORFA$, NUSERIE$, ID, DEPOINGRE%, FORMATNUM$(CANTI, "F15.4"))
       
       Call CargaEstructuraMultinivel(Me.MSF, CI1%, 100)
       Call ActualizarInfo

ERROR_GUARDAR:
    Screen.MousePointer = 1
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Transferencia Realizada")
    End If

   End If

End Sub

Private Sub Text1_Change()
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then
      GoTo 99
   End If
   '
   NUSERIE$ = TRIM$(Text1)
   If NUSERIE$ = "" Then
      GoTo 99
   End If
   '
   CI1% = CODINT%(Text10)
   If CI1% < 1 Then
      GoTo 99
   End If

   REG& = 0
   'SI YA EXISTE CON EL N.SERIE TRAE TODA LA INFO DE LA TABLA CONTROSERIE    - PARA QUE TRAIGA LOS REGISTROS ORIGINALES Y NO LOS QUE TIENEN INGRESOS
   CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND (CANTINGRE <= 0 AND CANTSALID <= 0)  " ' LOS APLICA POR VALOBADA CON UN SUM
   If CantRegistros("CONTROSERIE", CONDI$, "NOMESS") > 0 Then
      Call CargarDatosControSerie(NORFA$, NUSERIE$)
      GoTo 99
      
   Else 'SI NO TRAE LA ESTRUCTURA
      Call CargaEstructuraMultinivel(Me.MSF, CI1%, 100)
      Call ActualizarInfo
   End If
      
   
   
99 Command6.Enabled = True
   
  
End Sub

Sub ActualizarInfo()

   Campos$ = CamposOriginal$
   Erase CADENATEX$
   If TRIM$(TXTDEPOS) <> "" Then
      Call TEXTOLOTES(TXTDEPOS, ",")
        For i& = 1 To UBound(CADENATEX$)
           Campos$ = Campos$ + ";Dep." & CADENATEX$(i&) & "/A12"
        Next i&
   End If
   Call CARGA_ENCABEZADO(MSF, Campos$, CONTRCOM)
   Call CARGA_ENCABEZADO(MSF2, Campos$, CONTRCOM)
   CI1% = CODINT%(Text10)
   NORFA$ = TRIM$(Label4)
   NUSERIE$ = TRIM$(Text1)
   If NUSERIE$ <> "" And NORFA$ <> "" Then
    CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "' AND IDSUBOR = '" & NORFA$ & "' AND (CANTINGRE <= 0 AND CANTSALID <= 0)  " ' LOS APLICA POR VALOBADA CON UN SUM
    If CantRegistros("CONTROSERIE", CONDI$, "NOMESS") > 0 Then
       Call CargarDatosControSerie(NORFA$, NUSERIE$)
    Else
      Call CargaEstructuraMultinivel(Me.MSF, CI1%, 1)
    End If
   End If
   indDepo% = 0
   'Call CargarGrilla(CI1%)
   If TRIM$(TXTDEPOS) = "" Then Exit Sub
   CantEnDepositoAcumulado = 0
   cantidadColumnasOriginal = 9
   cantidadColumnasAgregadas = XVALO(UBound(CADENATEX$))
   totalColumnas = cantidadColumnasOriginal + cantidadColumnasAgregadas
   For REG& = 1 To MSF.Rows - 1
   CantEnDepositoAcumulado = 0
    For COL& = cantidadColumnasOriginal + 1 To totalColumnas
        CI1% = XVALO(MSF.TextMatrix(REG&, 1))
        indDepo% = indDepo% + 1
        If Int(CADENATEX$(indDepo%)) = 5 Then 'esto es por que arcon en el deposito 5 tiene lo pendiente de C.C.
           CantEnDeposito = DepositoRecepcionArcon(CI1%)
        Else
           CantEnDeposito = STODEPOS(CI1%, Int(CADENATEX$(indDepo%))) ' Y quiere que lo presente
        End If
        
        MSF.TextMatrix(REG&, COL&) = FORMATNUM$(CantEnDeposito, "F9.1")
'        CantEnDepositoAcumulado = CantEnDepositoAcumulado + CantEnDeposito
    Next COL&
    'MSF.TextMatrix(REG&, 7) = FORMATNUM$(CantEnDepositoAcumulado, "F12.1")
    indDepo% = 0
   Next REG&

End Sub
Function DepositoRecepcionArcon(CI1%)
        TORECEP = 0
        
     SQLX$ = " SELECT CANTNOM, NumProv, NUMEBORE , FECHRECEP, NumRem_Prov, CanToApro  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + " AND Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '

     Set COBRECEPTMP = READSET(SQLX$)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         k& = 0
         Do While (.EOF = False)
            CARECEP = XVALO(!CantNom) - XVALO(!CANTOAPRO)
            If CARECEP > 0 Then
              NCI% = XVALO(!NumProv)
              REMIPRO$ = SAFETEXT$(!NumRem_Prov)
              NROBOL& = XVALO(!NUMEBORE)
              FEX = CVINT(!FechRecep)
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     COBRECEPTMP.Close
     Set COBRECEPTMP = Nothing
     
    Call COPYSTRU("BRECEPEN", "BRECEPEL")
    Call FILESORT("BRECEPEN", "BRECEPEL", 1, 1)
    FIN& = ULTREG&("BRECEPEL")
    
    For U& = 1 To FIN&
      TXTX$ = REGLEIDO$("BRECEPEL", U&)
      CIII% = CVI(Mid$(TXTX$, 3, 2))
      If CIII% = CI1% Then
           DOPRI$ = TRIM$(Mid$(TXTX$, 5, 16))
           CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
           CANTSALID = CVD(Mid$(TXTX$, 96, 8))
           CANTI = CANTINGRE - CANTSALID
           If XVALO(CANTI) > 0.05 Then
             TORECEP = TORECEP + XVALO(CANTI)
           End If
      End If
    Next U&
    DepositoRecepcionArcon = TORECEP

End Function
Private Sub Text3_Change()
   If XVALO(Text3) < 1 And TRIM$(Text3) <> "" Then
      Call COMUNI("Cantidad Ingresada No Válida")
      Exit Sub
   End If
   Call T_Espera(1)

   Call ActualizarInfo
End Sub

Private Sub Text1_DblClick()
   SerieOriginal$ = TRIM$(Text1)
   If SerieOriginal$ <> "" Then Exit Sub 'EN PRINCIPIO SI TIENE N.SERIE, NO SE DEBERIA PODER EDITAR.
   NUSERIE$ = TRIM$(InputBox$("Ingrese N° Serie", "Número de Serie", SerieOriginal$))
   If SerieOriginal$ <> "" And SerieOriginal$ <> NUSERIE$ Then
      OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Desea Cambiar el N° De Serie Actual " & vbCrLf & " Esta Acción Cambia el N° de Serie Para Todos los Componentes")
      If OPX% <= 1 Then Exit Sub
   End If
   Text1 = NUSERIE$

End Sub

Private Sub Text4_Change()
    Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text4)))
End Sub

Private Sub TXTDEPOS_Change()
Call ActualizarInfo
End Sub


Private Sub TXTNIVELES_Change()
   Call T_Espera(1)

   Call ActualizarInfo
End Sub
