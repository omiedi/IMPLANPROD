VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECONCI09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reconciliacion para Orden de Produccion"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   1050
      Left            =   10980
      ScaleHeight     =   990
      ScaleWidth      =   3435
      TabIndex        =   34
      Top             =   315
      Width           =   3495
      Begin VB.CommandButton Command32 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         Picture         =   "RECONCI09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Listados de Balance de Requerimientos"
         Top             =   150
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         Height          =   690
         Left            =   2595
         Picture         =   "RECONCI09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Salir"
         Top             =   150
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   915
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   37
         Top             =   240
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -135
            TabIndex        =   38
            Top             =   -60
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Save"
         Height          =   1005
         Left            =   105
         Picture         =   "RECONCI09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Guardar Cierre de O/F y Reingreso de Materiales a O/F"
         Top             =   6360
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         Height          =   690
         Left            =   1845
         Picture         =   "RECONCI09.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   150
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   510
         Picture         =   "RECONCI09.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   420
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   7425
      TabIndex        =   22
      Top             =   90
      Width           =   3120
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realizado:"
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
         Left            =   -75
         TabIndex        =   32
         Top             =   705
         Width           =   1365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proyectado:"
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
         Left            =   -90
         TabIndex        =   31
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Realizado:"
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
         Left            =   -90
         TabIndex        =   30
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1350
         TabIndex        =   29
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1350
         TabIndex        =   28
         Top             =   645
         Width           =   1560
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Prod.:"
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
         Left            =   -285
         TabIndex        =   27
         Top             =   1875
         Width           =   1560
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PCP:"
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
         Left            =   -75
         TabIndex        =   26
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1335
         TabIndex        =   25
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1335
         TabIndex        =   24
         ToolTipText     =   "Proceso Control de Producto"
         Top             =   1470
         Width           =   1185
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1335
         TabIndex        =   23
         Top             =   1815
         Width           =   1200
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7470
      TabIndex        =   21
      Top             =   1170
      Width           =   3045
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame3"
      Height          =   3720
      Left            =   3615
      TabIndex        =   20
      Top             =   2055
      Width           =   30
   End
   Begin VB.CommandButton Command5 
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
      Left            =   3180
      TabIndex        =   19
      Top             =   210
      Width           =   175
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   3675
      Left            =   11490
      TabIndex        =   18
      Top             =   2100
      Width           =   30
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame3"
      Height          =   3720
      Left            =   9915
      TabIndex        =   17
      Top             =   2055
      Width           =   30
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Materiales de la Orden de Produccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   120
      TabIndex        =   10
      Top             =   1665
      Width           =   14400
      Begin VB.Frame Frame10 
         Caption         =   "Frame3"
         Height          =   3675
         Left            =   12945
         TabIndex        =   33
         Top             =   435
         Width           =   30
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3735
         Left            =   1935
         TabIndex        =   14
         Top             =   375
         Width           =   30
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame3"
         Height          =   3705
         Left            =   6630
         TabIndex        =   13
         Top             =   420
         Width           =   30
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   3720
         Left            =   8220
         TabIndex        =   12
         Top             =   390
         Width           =   30
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame3"
         Height          =   3735
         Left            =   5070
         TabIndex        =   11
         Top             =   390
         Width           =   30
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         ItemData        =   "RECONCI09.frx":2AC2
         Left            =   240
         List            =   "RECONCI09.frx":2AC9
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   14040
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"RECONCI09.frx":2B5A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   225
         TabIndex        =   16
         Top             =   375
         Width           =   14055
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "RECONCI09.frx":2C15
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "RECONCI09.frx":2E49
      TabIndex        =   42
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   3270
      TabIndex        =   9
      Top             =   675
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O/Prod. Nº:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   270
      Width           =   1125
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   5820
      TabIndex        =   7
      Top             =   210
      Width           =   1020
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Movimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4485
      TabIndex        =   6
      Top             =   90
      Width           =   1155
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   4875
      TabIndex        =   5
      Top             =   615
      Width           =   1965
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1365
      TabIndex        =   4
      Top             =   1110
      Width           =   5475
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   150
      TabIndex        =   3
      Top             =   1170
      Width           =   1125
   End
   Begin VB.Label Label32 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1380
      TabIndex        =   2
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   195
      Width           =   1800
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1380
      TabIndex        =   1
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   660
      Width           =   1800
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro S/F:"
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
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   1125
   End
End
Attribute VB_Name = "RECONCI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub
Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
   Call CalculoRecon
End Sub

Private Sub Command2_Click()
 
10 Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()
    CONDI$ = "StatuOrf = 3 "
    CONDI$ = CONDI$ + " and NuOrProd <> '' "
    CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
    Call CommSel("ORDPEN/Indice de Ordenes de Produccion Pendientes", Command5, Label32)
    
    If Label32 <> "" Then
       On Error Resume Next
       Text8.SetFocus
       On Error GoTo 0
    End If
End Sub

Private Sub Form_Load()
  Call APLICACIONSKINS(Me)
  List1.Clear
End Sub

Private Sub Label32_Change()
    NORFA$ = Label32
    Label24 = HOS$
    'LANCO TODOS LOS RESTANTES CAMPOS Y LIST'
    Label28 = ""
    Label30 = ""
    Text5 = ""
    Text6 = ""
    Text2 = ""
    Text3 = ""
    Label34 = ""
    Label37 = ""
    Label18 = ""
    Label26 = ""
    Label19 = ""
    Text7 = ""
    Text8 = ""
    
    List1.Clear
    '
    If NORFA$ <> "" Then
         ' presenta datos de la orden de fabricación
         Call ABREORFAB
         SQLX$ = " SELECT * FROM ORDEFABR "
         SQLX$ = SQLX$ + " WHERE NuOrProd = '" & NORFA$ & "'"
         '
         Dim ORDEFABTMP As ADODB.Recordset
         Set ORDEFABTMP = New ADODB.Recordset
         Set ORDEFABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         With ORDEFABTMP
           On Error Resume Next
           .MoveFirst
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "Orden de Fabricación no Válida. Consulte.")
              Call FINAL("")
           End If
           On Error GoTo 0
         '
         'Datos del Encabezado
         CIXI% = XVALO(!Cod_Inte)
         Label28 = CODEXT$(CIXI%)
         Label30 = DESCRIT0$(CIXI%)
         Label34 = FORMATNUM$(!CanProy, "F13." + CNTDC$(CIXI%))
         Label26 = SAFETEXT(!IdSubOr)
         '
         CAREAL# = XVALO(!CanApro) + XVALO(!CanRech)
         Label37 = FORMATNUM$(CAREAL#, "F13." + CNTDC$(CIXI%))
         '
         List1.Clear
       End With
       '
       'Si esta registrado el calculo se levanta de la Tabla Reconcilia
       CONDI$ = "IDSUBOR ='" & TRIM$(Label26) & "'"
       If CantRegistros("RECONCILIA", CONDI$) > 0 Then
          Call LeerCalculo
        Else
          COMUNI ("No se ha Generado Calculo de Reconciliacion para OP. " & NORFA$)
          GoTo 99
       End If
       '
    End If
99  Screen.MousePointer = 1
End Sub

Sub CalculoRecon()
    DESTINO$ = Label26
    List1.Clear
    'CARGAR EL LIST CON LOS REQUERIMIENTOS
    
    'Calcular Coeficiente
    CANTPROY# = XVALO(Label34)
    TOTPRODUCIDO# = XVALO(Label37)       '   CONFORME + NO CONF + PCP
    On Error Resume Next
    COEFIC# = TOTPRODUCIDO# / CANTPROY#
    On Error GoTo 0
    '
    CONDI$ = "DESTINO='" & DESTINO$ & "' AND KILO_ASIG > 0 "
    SQLX$ = "SELECT cod_inte, idenlote, SUM(KILO_ASIG) AS KASIG FROM BALAREQMO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " + CONDI$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, IDENLOTE "
    Set Reque = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With Reque
      Do While Not (.EOF)
       LINEA$ = ""
       CI% = XVALO(!Cod_Inte)
       CODIGO$ = CODEXT$(CI%)
       LOTE$ = SAFETEXT(!idenLote)
       CANTASIG# = XVALO(!KASIG)
       CANTESTI# = CANTASIG# * COEFIC#
       '
       Call REPLA(LINEA$, CODIGO$, 1, 16)
       Call REPLA(LINEA$, LOTE$, 17, 16)
       CANESTI$ = FORMATSTAN(CANTESTI#, "F13", CI%)
       Call REPLA(LINEA$, CANESTI$, 92, 13)
       '
       List1.AddItem LINEA$
       .MoveNext
      Loop
    End With
    '
    For U& = 0 To List1.ListCount - 1
      'Cantidad Consumida para este lote
      LINEA$ = List1.List(U&)
      CODIGO$ = TRIM(Mid$(LINEA$, 1, 16))
      CIXI% = CODINT%(CODIGO$)
      LOTE$ = TRIM$(Mid$(LINEA$, 17, 16))
      
      CONDI$ = "DoSecLar = '" & DESTINO$ & "' "
      CONDI$ = CONDI$ + " AND CODIMOVI = 'CF' "
      CONDI$ = CONDI$ + " AND IDENLOTE ='" & LOTE$ & "'"
      CantCons# = XVALO(VALOBADA("MOVISTO", "  SUM(CANTSALID) - SUM(CANTINGRE)", CONDI$))
      
      'Cantidad Merma para este lote
      CONDI$ = "DoSecLar = '" & DESTINO$ & "' "
      CONDI$ = CONDI$ + " AND CODIMOVI = 'MP' "
      CONDI$ = CONDI$ + " AND IDENLOTE ='" & LOTE$ & "'"
      CANTMERMA# = XVALO(VALOBADA("MOVISTO", "CANTSALID", CONDI$))
      '
      'Cantidad Devolucion para este lote
      CONDI$ = "DoSecLar = '" & DESTINO$ & "' "
      CONDI$ = CONDI$ + " AND CODIMOVI = 'DV' "
      CONDI$ = CONDI$ + " AND IDENLOTE like '" & LOTE$ & "%'"
      CANTDEVOL# = XVALO(VALOBADA("MOVISTO", "CANTINGRE", CONDI$))
      '
      'Consumo Neto
      CONSNETO# = CantCons# + CANTMERMA# - CANTDEVOL#
      '
      Call REPLA(LINEA$, FORMATSTAN(CantCons#, "F13.", CIXI%), 32, 13)
      Call REPLA(LINEA$, FORMATSTAN(CANTMERMA#, "F13.", CIXI%), 47, 13)
      Call REPLA(LINEA$, FORMATSTAN(CANTDEVOL#, "F13.", CIXI%), 62, 13)
      Call REPLA(LINEA$, FORMATSTAN(CONSNETO#, "F13.", CIXI%), 77, 13)
      '
      'Calcular Diferencia
      CONDI$ = "DESTINO='" & DESTINO$ & "' AND IDENLOTE = '" & LOTE$ & "'"
      CANTASIG# = VALOBADA("BALAREQMO", "SUM(KILO_ASIG)", CONDI$)
      CANTESTI# = CANTASIG# * COEFIC#
      DIFERENCIA# = CONSNETO# - CANTESTI#
      Call REPLA(LINEA$, FORMATSTAN(DIFERENCIA#, "F13.", CIXI%), 107, 13)
      '
      'Calcular Porcentaje
      
      TOT# = CANTESTI#
      PARTE# = DIFERENCIA#
      On Error Resume Next
      PORCENT# = (PARTE# * 100) / TOT#
      If Err > 0 Then
         PORCENT# = 0
      End If
      On Error GoTo 0
      TPOR$ = FORMATNUM(PORCENT#, "F7.2")
      TPOR$ = TPOR$ & " %"
      Call REPLA(LINEA$, TPOR$, 122, 7)
      '
      List1.List(U&) = LINEA$
      '
    Next U&
    Call GuardarCalculo
    
         
End Sub
Sub GuardarCalculo()

    SF$ = TRIM$(Label26)
    CONDI$ = "IDSUBOR ='" & SF$ & "'"
    Call BORRAREGISTROS("RECONCILIA", CONDI$, REG&)
    SQLX$ = "SELECT * FROM RECONCILIA WHERE IDSUBOR ='" & SF$ & "'"
    '
    Dim RECONCI As ADODB.Recordset
    Set RECONCI = New ADODB.Recordset
    '
    RECONCI.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic
    '
    With RECONCI
     For U& = 0 To List1.ListCount - 1
      
      LINEA$ = List1.List(U&)
      '
      'SE OBTINEN LOS DATOS DEL LIST
      CODIGO$ = TRIM(Mid$(LINEA$, 1, 16))
      CIXI% = CODINT%(CODIGO$)
      LOTE$ = TRIM$(Mid$(LINEA$, 17, 16))
      CantCons# = XVALO(Mid$(LINEA$, 32, 13))
      CANTMERMA# = XVALO(Mid$(LINEA$, 47, 13))
      CANTDEVOL# = XVALO(Mid$(LINEA$, 62, 13))
      CONETO# = XVALO(Mid$(LINEA$, 77, 13))
      DIFER# = XVALO(Mid$(LINEA$, 107, 13))
      COFORM# = XVALO(Mid$(LINEA$, 92, 13))
      '
      'SE AGREGA EL REGISTRO
      .AddNew
      !IdSubOr = SF$
      !Cod_Inte = CIXI%
      !idenLote = LOTE$
      !CONSUMIDO = CantCons#
      !MERMAS = CANTMERMA#
      !DEVOLUCION = CANTDEVOL#
      !CONSUNETO = CONETO#
      !CONSUFORMU = COFORM#
      !DIFERENCIA = DIFER#
      .Update
    Next U&
    
    End With
End Sub

Sub LeerCalculo()
    '
    SF$ = TRIM$(Label26)
    SQLX$ = "SELECT * FROM RECONCILIA WHERE IDSUBOR ='" & SF$ & "'"
    '
    Dim RECONCI As ADODB.Recordset
    Set RECONCI = New ADODB.Recordset
    Set RECONCI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With RECONCI
      Do While Not (.EOF)
      LINEA$ = ""
      CIXI% = XVALO(!Cod_Inte)
      CODIGO$ = CODEXT$(CIXI%)
      LOTE$ = SAFETEXT$(!idenLote)
      CantCons# = XVALO(!CONSUMIDO)
      CANTMERMA# = XVALO(!MERMAS)
      CANTDEVOL# = XVALO(!DEVOLUCION)
      CONETO# = XVALO(!CONSUNETO)
      COFORM# = XVALO(!CONSUFORMU)
      DIFER# = XVALO(!DIFERENCIA)
      '
      'Calcular Porcentaje
      TOT# = COFORM#
      PARTE# = DIFER#
      On Error Resume Next
      PORCENT# = (PARTE# * 100) / TOT#
      If Err > 0 Then
         PORCENT# = 0
      End If
      On Error GoTo 0
      '
      TPOR$ = FORMATNUM(PORCENT#, "F7.2")
      TPOR$ = TPOR$ & " %"
      '
      Call REPLA(LINEA$, CODIGO$, 1, 16)
      Call REPLA(LINEA$, LOTE$, 17, 16)
      '
      AX$ = FORMATSTAN(CantCons#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 32, 13)
      '
      AX$ = FORMATSTAN(CANTMERMA#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 47, 10)
      '
      AX$ = FORMATSTAN(CANTDEVOL#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 62, 13)
      '
      AX$ = FORMATSTAN(CONETO#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 77, 13)
      '
      AX$ = FORMATSTAN(COFORM#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 92, 13)
      '
      AX$ = FORMATSTAN(DIFER#, "F13", CIXI%)
      Call REPLA(LINEA$, AX$, 107, 13)
      
      Call REPLA(LINEA$, TPOR$, 122, 7)
      List1.AddItem LINEA$
      '
      .MoveNext
      Loop
      '
    End With

End Sub
