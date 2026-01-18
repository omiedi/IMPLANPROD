VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INFOUBIC 
   BackColor       =   &H00E2D3C0&
   Caption         =   "Ubicaciones de Materiales en Proceso"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   7920
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Informar Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8800
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15195
      Begin VB.TextBox lblidregistro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   " "
         Top             =   600
         Width           =   810
      End
      Begin VB.TextBox TXTMENSAJE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "INFOUBIC.frx":0000
         Top             =   7680
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13440
         TabIndex        =   23
         Text            =   " "
         Top             =   240
         Width           =   940
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   670
         Left            =   14475
         Picture         =   "INFOUBIC.frx":0073
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   600
      End
      Begin VB.CommandButton Command3 
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
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTCANTICORTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   600
         Width           =   696
      End
      Begin VB.TextBox TXTCORTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9150
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   600
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox TXTCODELEM 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   240
         Width           =   1656
      End
      Begin VB.TextBox TXTLOTE 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Width           =   1770
      End
      Begin VB.CommandButton Command2 
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
         Height          =   315
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "O.K."
         Top             =   600
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   600
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "."
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
         Left            =   3990
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   200
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   7650
         Left            =   3960
         TabIndex        =   22
         ToolTipText     =   "Doble Clic para Ubicar Lote o Quitar Lote-"
         Top             =   1080
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   13494
         _Version        =   393216
         Rows            =   26
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   6555
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Clic para Filtrar por Rack"
         Top             =   1080
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   11562
         _Version        =   393216
         Cols            =   4
         BackColor       =   8421504
         ForeColor       =   16777152
         BackColorFixed  =   16501405
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Id Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   10920
         TabIndex        =   31
         Top             =   675
         Width           =   855
      End
      Begin VB.Label LBLNUMEORDENINFO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4200
         TabIndex        =   28
         Top             =   8640
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro x Rack"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12240
         TabIndex        =   24
         Top             =   315
         Width           =   1125
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   4530
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   19
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. Cortes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   12480
         TabIndex        =   16
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   8235
         TabIndex        =   15
         Top             =   675
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   5640
         TabIndex        =   12
         Top             =   315
         Width           =   765
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   8040
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4200
         TabIndex        =   4
         Top             =   600
         Width           =   2970
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   3
         Top             =   660
         Width           =   1125
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFOUBIC.frx":01BD
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFOUBIC.frx":03F1
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuCtaListado 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver "
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuOcupada 
         Caption         =   "Marcar/Desmarcar Ubicaciones ocupadas"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTabAux 
         Caption         =   "Tablas Auxiliares"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "INFOUBIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NOUBICAR%




Sub CARGARGRILARACK()
   REG& = 0
   SQLX$ = "SELECT  CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'CONFUBIX' ORDER BY CAMPO1"
   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
       REG& = REG& + 1
       MSF2.Rows = REG& + 1
       MSF2.TextMatrix(REG&, 1) = XVALO(!CAMPO1)
       MSF2.TextMatrix(REG&, 2) = SAFETEXT$(!campo2)
       .MoveNext
      Loop
   End With
End Sub

Sub GRABARUBICA(NORFA$, LOTE$, CANTI, CODIEXLM$, RK, FL, KL, NOP%, DESOPER$, AGREGACANT, AGREGALOTE$, DESCRACK$, Optional LblId1&)
    NPED& = NroPedidoSegunOf&(NORFA$)
    NC = XVALO(VALOBADA("PEDENCA", "NroClie", "NROPED = " & NPED&, "NOMESS"))
    rs$ = rasocli$(NC)
    CONDI$ = "RACK = " & RK
    CONDI$ = CONDI$ + " AND FILA = " & FL
    CONDI$ = CONDI$ + " AND COLUMNA = " & KL
    EXISTEFILA% = CantRegistros&("UBILOTE_ENVAP", CONDI$, "NOMESS")
'    AGREGALOTE$ = TRIM$(AGREGALOTE$)
'    If AGREGALOTE$ <> "" Then
'       LOTE$ = AGREGALOTE$ & "; " & LOTE$
'    End If
    '
'    On Error GoTo ERROR_GUARDAR
'    FLEXCONN.BeginTrans
'    '
'    If EXISTEFILA% > 0 And AGREGACANT < 1 Then
'       SQLX$ = "UPDATE UBILOTE_ENVAP SET LOTE = '" & LOTE$ & "'"
'       SQLX$ = SQLX$ + ", CANTIDAD = " & CANTI + AGREGACANT
'       SQLX$ = SQLX$ + ", MARBORRA = 0"
'       SQLX$ = SQLX$ + ", CODINT = " & CODINT%(CODIEXLM$)
'       SQLX$ = SQLX$ + ", NUCLIEN = " & NC
'       SQLX$ = SQLX$ + ", FECHEMISI = '" & FETEXCOR1$(HOY(HO$)) & "'"
'       SQLX$ = SQLX$ + ", IDSUBOR =   '" & NORFA$ & "'"
'       SQLX$ = SQLX$ + ", RASO_CLI = '" & RS$ & "'"
'       SQLX$ = SQLX$ + ", DescOper = '" & DESOPER$ & "'"
'       SQLX$ = SQLX$ + ", NumeOper = " & NOP%
'       SQLX$ = SQLX$ + ", NROPED = " & NPED&
'       SQLX$ = SQLX$ + ", CodiEmpr = " & CodiEmp%
'       SQLX$ = SQLX$ + ", DescriRack = '" & DescRack$ & "'"
'       SQLX$ = SQLX$ + " WHERE " & Condi$
'       FLEXCONN.Execute (SQLX$)
'    Else
'       SQLX$ = "INSERT INTO UBILOTE_ENVAP "
'       SQLX$ = SQLX$ + " (LOTE,CANTIDAD,MARBORRA,CODINT,NUCLIEN,FECHEMISI"
'       SQLX$ = SQLX$ + " ,IDSUBOR,RASO_CLI,DescOper,NumeOper,NROPED,CodiEmpr,RACK,FILA,COLUMNA,DescriRack)"
'       SQLX$ = SQLX$ + " VALUES('" & LOTE$ & "'"
'       SQLX$ = SQLX$ + " ," & CANTI '+ AGREGACANT
'       SQLX$ = SQLX$ + " ,0"
'       SQLX$ = SQLX$ + " ," & CODINT%(CODIEXLM$)
'       SQLX$ = SQLX$ + " ," & NC
'       SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'"
'       SQLX$ = SQLX$ + " ,'" & NORFA$ & "'"
'       SQLX$ = SQLX$ + " ,'" & RS$ & "'"
'       SQLX$ = SQLX$ + " ,'" & DESOPER$ & "'"
'       SQLX$ = SQLX$ + " ," & NOP%
'       SQLX$ = SQLX$ + " ," & NPED&
'       SQLX$ = SQLX$ + " ," & CodiEmp%
'       SQLX$ = SQLX$ + " ," & RK
'       SQLX$ = SQLX$ + " ," & FL
'       SQLX$ = SQLX$ + " ," & KL
'       SQLX$ = SQLX$ + " ,'" & DescRack$ & "'"
'
'       SQLX$ = SQLX$ + ")"
'       FLEXCONN.Execute (SQLX$)
'    End If
    '
    'Guarda las Cordenadas de ubicacion in infofab
    CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
    'Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
    CONDI$ = CONDI$ + " AND  NUORINF = " & XVALO(LBLNUMEORDENINFO)
    CONDI2$ = CONDI$ + " AND COORDUBIC = '" & CoordenadasUbicacion$ & "'"
    'ESTO SE AGREGO DESPUES SI SE LEVANTA DE LA PANTALLA DE INFOCORTE VIENE CON EL ID PARA ASEGURAR LA TRANSACCION
    If LblId1& > 0 Then
       CONDI$ = "INFOFAB_ID =" & LblId1&
    End If
    
    If CantRegistros("INFOFAB", CONDI2$, "NOMESS") > 0 Then
         Call COMUNI("Atención !!! Imposible Ingresar Misma Operación y Lote en en Misma Ubicación\Libere la Ubicación y Vuelva a Intentar o Seleccione una Nueva Ubicación")
       Exit Sub
    End If
    Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, CoordenadasUbicacion$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, 0, REGAF&, "NOMESS")


'    Call ACTUREGISTRO("INFOFAB", "IDENLOTE", Condi$, LOTE$, REGAF&, "NOMESS")
    
'ERROR_GUARDAR:
'    If Err <> 0 Then
'       On Error Resume Next
'       On Error GoTo 0
'       AAA$ = Err.Description
'       Call FRATEXTO(AAA$, 50)
'       TEXT$ = ""
'       For KU& = 1 To CARETEX%
'         TEXT$ = TEXT$ + RETEX$(KU&) + "\"
'       Next KU&
'
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION\" & TEXT$)
'       FLEXCONN.RollbackTrans
'       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
'      Call COMUNI("Proceso Realizado Exitosamente")
'    End If

End Sub

Sub MARCAROCUPADAS(RK, FL, KL)
    CONDI$ = "RACK = " & RK
    CONDI$ = CONDI$ + " AND FILA = " & FL
    CONDI$ = CONDI$ + " AND COLUMNA = " & KL
    EXISTEFILA% = CantRegistros&("UBILOTE_ENVAP", CONDI$, "NOMESS")
    If XVALO(MSF1.TextMatrix(FL, 5)) < 1 Then
    
    
    End If

    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    '
    If EXISTEFILA% > 0 And AGREGACANT < 1 Then
       SQLX$ = "UPDATE UBILOTE_ENVAP SET LOTE = 'OCUPADO'"
       SQLX$ = SQLX$ + ", IDSUBOR =   'OCUPADO'"
       SQLX$ = SQLX$ + ", CodiEmpr = " & CodiEmp%
       SQLX$ = SQLX$ + ", DescriRack = '" & DESCRACK$ & "'"
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
    Else
       SQLX$ = "INSERT INTO UBILOTE_ENVAP "
       SQLX$ = SQLX$ + " (LOTE,CodiEmpr,IDSUBOR,DescriRack"
       SQLX$ = SQLX$ + " VALUES('" & LOTE$ & "'"
       SQLX$ = SQLX$ + " ," & CodiEmp%
       SQLX$ = SQLX$ + " ,'OCUPADO'"

       SQLX$ = SQLX$ + ")"
       FLEXCONN.Execute (SQLX$)
    End If
    
ERROR_GUARDAR:
    If Err <> 0 Then
       On Error Resume Next
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       FLEXCONN.RollbackTrans
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Proceso Realizado Exitosamente")
    End If
    
    
End Sub

Private Sub Command2_Click()
     Command2.Enabled = False
    '
     NROOF$ = TRIM$(Label4)

     '
     CIELEM% = CODINT(TXTCODELEM)
     If CIELEM% < 1 Then
        Call MENSERR%(24, "Codigo Componente Inválido")
        GoTo 99
     End If

     If NROOF$ = "" Then
        Call MENSERR%(24, "Debe Seleccionar la Orden de Fabricación.")
        GoTo 99
     End If
     '
     Condi1$ = "CodIConj = " & CI% & " and  NumeOper = " & XVALO(Text3)
     If VALOBADA("SECOPER", "DescOper", Condi1$) = "" Then
         Call MENSERR(24, "Operación Incorrecta para el Articulo Seleccionado.")
         GoTo 99
     End If
     '
     If TRIM$(TXTLOTE) = "" Then
         Call MENSERR(24, "Falta Lote Para el Registro Seleccionado.")
         GoTo 99
     End If

99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
    End If
    '
    CONDI$ = "COD_INTE = " & CIXI% & " AND StatuOrf < 3"
    REG& = CantRegistros&("ORDEFABR", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Ordenes de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!ECORDEF1/Ordenes de Fabricacion Cerradas")
    If VALACT1$("!ECORDEF1") <> "" Then
     Label4 = VALACT1$("!ECORDEF1")
     Label2 = VALACT2$("!ECORDEF1")
    End If
    '
End Sub

Private Sub Command4_Click(Index As Integer)
  
End Sub '

Private Sub Command5_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
    Text4.TEXT = VDEV$
  End If
End Sub

Private Sub Command6_Click()
    '
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
        If TRIM$(Label4) = "" Then
            Call MENSERR(24, "Primero debe Seleccionar la Orden de Fabricacion")
            Exit Sub
        End If
    End If
    '
    CONDI$ = "CodIConj = " & CIXI%
    REG& = CantRegistros&("SECOPER", CONDI$)
    If REG& < 1 Then
        Call MENSERR(24, "No Existen Operaciones de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!OPERINF1/Operaciones de Ordenes de Fabricacion")
    If TRIM$(VALACT1$("!OPERINF1")) <> "" Then
      Text3.TEXT = TRIM$(VALACT1$("!OPERINF1"))
'      Label7.Caption = TRIM$(VALACT2$("!OPERINF1"))
    End If
    '
End Sub

Private Sub Command7_Click()
   '
   Call SELECHO("PADMAQ")
   VDEVU$ = VALACT1$("PADMAQ")
   If VDEVU$ <> "" Then
      Text5.TEXT = TRIM$(VDEVU$)
      Label10.Caption = TRIM$(VALACT2$("PADMAQ"))
   End If
   '
End Sub

Private Sub Command8_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub



Private Sub Command9_Click()
   
End Sub

Private Sub Form_Load()
    '
    Call CARGA_ENCABEZADO(MSF1, "Rack/F3;Descripción/A24;Piso/F4;Picking/F7;Cant./F5;Lote-s/A24;O.F./A12;ORDEN/A0;INFOFAB_ID/A0;NC/A0;Razon Social/A24;Desc.Oper./A16", Me)
    MSF1.FixedCols = 4
    MSF1.ScrollTrack = True
    MSF1.ScrollBars = flexScrollBarBoth
    MSF1.MergeCells = flexMergeFree
    MSF1.MergeCol(1) = True
    MSF1.MergeCol(2) = True
    MSF1.MergeCol(3) = True
    Text4 = FECHATEX$(HOY(HO$))
    Call ARMARGRILLAUBICA(NRACK%)
    Call CARGA_ENCABEZADO(MSF2, "Rack/F4;Descripción/A28;Nada/A0", Me)
    Call CARGARGRILARACK
    
End Sub

Private Sub Label19_Click()

End Sub

Sub CargaMatTranfXOrdenSoloIngresos()
 '********************************************************************
 'VER ACA SI ESTA BIEN QUE MUESTRE SOLO CANTINGRE EN VEZ DE CANTINGRE - CANTSALID
  '********************************************************************

  MSF1.Rows = 1: REG& = 0
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  
  SQLX$ = "SELECT sum(CANTINGRE) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION "
  SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where doprilar = 'TR." & TRIM(Label4) & "'  AND DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
  SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  Set rs = READSET(SQLX$)
  With rs
    Do While Not .EOF
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!Resultado), "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       .MoveNext
    Loop
 End With
 rs.Close
 Set rs = Nothing
 
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Frame1.Width = Me.Width - 200
   Frame1.Height = Me.Height - 800
   
   Me.MSF1.Width = Frame1.Width - 4060
   Me.MSF1.Height = Me.Height - 2000
'   Frame1.Top = Me.Height - 1300
   
11145
   On Error GoTo 0
End Sub

Private Sub Label3_Change()
   Call ARMARGRILLAUBICA(XVALO(Label3))
End Sub

Private Sub Label3_DblClick()
  Label3 = ""
End Sub

Private Sub Label4_Change()
'   Call CargaMatTranfXOrdenSoloIngresos
End Sub


Private Sub mnuOcupada_Click()
   If txtMarcarOcupadas.Visible = True Then
      txtMarcarOcupadas.Visible = False
   Else
      txtMarcarOcupadas.Visible = True
   End If
    
End Sub

Private Sub mnuTabAux_Click()

' If DERACCE%("TABAUXGE", "ACCESO A TABLAS AUXILIARES GENERALES") < 2 Then Exit Sub
' Call MOSTRARTABLASCONFIGURADAS(Me.Name)
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)

End Sub



Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
    RK = XVALO(MSF1.TextMatrix(REG&, 1))
    DRK$ = MSF1.TextMatrix(REG&, 2)
    FL = XVALO(MSF1.TextMatrix(REG&, 3))
    KL = XVALO(MSF1.TextMatrix(REG&, 4))
'    If txtMarcarOcupadas.Visible = True Then
'       Call MARCAROCUPADAS(RK, FL, KL)
'       Exit Sub
'    End If
    
    CANTI = XVALO(MSF1.TextMatrix(REG&, 5))
    LOTE$ = MSF1.TextMatrix(REG&, 6)
    CANTIAGREGAR = XVALO(TXTCANTICORTES)
    LOTEAGREGAR$ = TXTLOTE.TEXT
    CODIEXLM$ = TRIM$(TXTCODELEM)
    NORFA$ = Label4
    NOP% = XVALO(Text3)
    NORDEN% = XVALO(MSF1.TextMatrix(REG&, 8))
    LblId1& = XVALO(MSF1.TextMatrix(REG&, 9))

    TXXK$ = "RACK: " & RK & " (" & DRK$ & ")"
    TXXK$ = TXXK$ + " -Piso: " & FL
    TXXK$ = TXXK$ + " -Picking: " & KL
    OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Quitar Lote\Agregar Lote")
    If OPX% < 2 Then Exit Sub

    'OPCIÓN DE QUITAR LOTE DE UBICACION
    If OPX% = 2 Then
       If CANTI < 1 Then
          Call COMUNI("No se Detecta Cantidad en Ubicación Seleccionada")
          Exit Sub
       Else
          Call ACTUALIZARUBICACION(RK, DRK$, FL, KL, LblId1&)
          GoTo 60
       End If
      End If
''          MSF1.TextMatrix(REG&, 5) = ""
''          MSF1.TextMatrix(REG&, 6) = ""
''          MSF1.TextMatrix(REG&, 7) = ""
''          MSF1.TextMatrix(REG&, 8) = ""
''          Condi$ = "IDSUBOR = '" & NORFA$ & "'"
''          Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
''          Condi$ = Condi$ + " AND NUORINF = " & NORDEN%
''          Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
''          Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
''          GoTo 60
'''          Condi$ = "RACK = " & RK
'''          Condi$ = Condi$ + " AND FILA = " & FL
'''          Condi$ = Condi$ + " AND COLUMNA = " & KL
'''          Call BORRAREGISTROS("UBILOTE_ENVAP", Condi$, REGAF&, "NOMESS")
''          CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
''          Condi$ = "IDSUBOR = '" & NORFA$ & "'"
''          Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
''          Condi$ = Condi$ + " AND CANTREAL = " & CANTI
''          Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
''          Condi$ = Condi$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
''          Condi$ = Condi$ + " AND  COORDUBIC = '" & CoordenadasUbicacion$ & "'"
''          Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
''    End If
''
''    OPCION DE INGRESAR LOTE A UBICACION SELECCIONADA.
    If OPX% = 3 Then
'       If NOUBICAR% > 0 Then Call COMUNI("Imposible Ubicar Pallet!\Debe Ingresarlo Mediante la Selección de un informe Activo")
       If CANTI > 0 Then
          OPX% = COMALTER("Se Detecta un lote en Ubicación Seleccionada, (Lote: " & LOTE$ & ")\\Cancelar\Sumar Lote en Posición Seleccionada")
          If OPX% < 2 Then Exit Sub
          If OPX% = 2 Then
              Call GRABARUBICA(NORFA$, LOTEAGREGAR$, CANTIAGREGAR, CODIEXLM$, RK, FL, KL, XVALO(Text3), Label7, CANTI, LOTE$, DRK$, XVALO(lblidregistro))
          End If
       Else
          Call GRABARUBICA(NORFA$, LOTEAGREGAR$, CANTIAGREGAR, CODIEXLM$, RK, FL, KL, XVALO(Text3), Label7, 0, "", DRK$, XVALO(lblidregistro))
       End If
    End If

60

    Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(MSF1, 5) 'limpia los registros solo borra el contenido a paritr de la columna 5
    Call MostrarUbicacionesGuardadas 'refresca la info guardada en bd.

    Unload Me
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call Command4_Click(Index)
End Sub





Private Sub Text19_Change()
    On Error Resume Next
    TXT$ = TRIM$(ECOARCH$("TAMOTIV", Chr$(Text19)))
    On Error GoTo 0
    If TXT$ <> "" Then
       Label55 = TXT$
    Else
       Label55 = ""
    End If
End Sub

Private Sub Text2_LostFocus()
  Text2 = HORATEX$(HORANUM(Text2.TEXT))
End Sub

Sub ACTUALIZARUBICACION(RK, DESCRACK$, FL, KL, Optional LblId1&)

    '
    TXXK$ = "RACK: " & RK & " (" & DESCRACK$ & ")"
    TXXK$ = TXXK$ + " -Piso: " & FL
    TXXK$ = TXXK$ + " -Picking: " & KL
'    OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Quitar Lote\Agregar Lote")
'    If OPX% < 2 Then Exit Sub
    '
    'OPCIÓN DE QUITAR LOTE DE UBICACION
    If OPX% = 2 Then
       If CANTI < 1 Then
          Call COMUNI("No se Detecta Cantidad en Ubicación Seleccionada")
          Exit Sub
       End If
    ElseIf OPX% = 3 Then
       CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
    End If
    
    NORFA$ = SELECCIONLOTEENUBICACION$(RK, FL, KL) 'PRESENTA LISTA DE SELECCION
    
    If NORFA$ = "" Then Exit Sub
    CODICOMP$ = RESP_ZELE_VECT(2)
    CiComp = RESP_ZELE_VECT(3)
    LOTE$ = RESP_ZELE_VECT(4)
    NOP% = XVALO(RESP_ZELE_VECT(5))
    DESCOP$ = RESP_ZELE_VECT(6)
    CANTI = XVALO(RESP_ZELE_VECT(7))
    ORDEN = XVALO(RESP_ZELE_VECT(8))
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    CONDI$ = CONDI$ + " AND NUMEOPER = " & NOP%
    CONDI$ = CONDI$ + " AND CANTREAL = " & CANTI
    CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND  CODIELEM = " & CiComp
    CONDI$ = CONDI$ + " AND  NUORINF = " & ORDEN
    If LblId1& > 0 Then
       CONDI$ = "INFOFAB_ID =" & LblId1&
    End If

    Call ACTUREGISTRO("INFOFAB", "COORDUBIC", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("INFOFAB", "IDCOORDUBIC", CONDI$, 0, REGAF&, "NOMESS")

    Unload Me
    
    
'    If NUID& < 1 Then Exit Sub
'
'    NOP% = XVALO(RESP_ZELE_VECT(3))
'    NORFA$ = TRIM$(RESP_ZELE_VECT(5))
'    LOTE$ = TRIM$(RESP_ZELE_VECT(6))
'    CANTI = XVALO(RESP_ZELE_VECT(7))
'    CODIEXLM$ = TRIM$(TXTCODELEM)
'
'
'    OPX% = COMALTER%("Opción de Ubicaciones\\Canelar\Quitar Solo Lote Seleccionado")
'    If OPX% < 2 Then Exit Sub
'    If OPX% = 2 Then
'        Call BORRAREGISTROS("UBILOTE_ENVAP", "UBILOTE_ENVAP_ID = " & NUID&, REGAF&, "NOMESS")
'        CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
'        Condi$ = "IDSUBOR = '" & Label4 & "'"
'        Condi$ = Condi$ + " AND NUMEOPER = " & NOP%
'        Condi$ = Condi$ + " AND CANTREAL = " & CANTI
'        Condi$ = Condi$ + " AND IDENLOTE = '" & LOTE$ & "'"
'        Condi$ = Condi$ + " AND  CODIELEM = " & CODINT%(CODIEXLM$)
'        Condi$ = Condi$ + " AND  COORDUBIC = '" & CoordenadasUbicacion$ & "'"
'        Call ACTUREGISTRO("INFOFAB", "COORDUBIC", Condi$, "", REGAF&, "NOMESS")
'nload Me
'      U
'    End If

End Sub
Function SELECCIONLOTEENUBICACION$(RK, FL, KL)
    'SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED FROM UBILOTE_ENVAP
    Campos$ = "O.F/a10" '1
    Campos$ = Campos$ + " ;Componente/A32" '2
    Campos$ = Campos$ + " ;CiComp/A0" '3
    Campos$ = Campos$ + " ;Lote/a12" '4
    Campos$ = Campos$ + " ;N°OP/f5" '5
    Campos$ = Campos$ + " ;Desc./A14" '6
    Campos$ = Campos$ + " ;Cant./F7.1" '7
    Campos$ = Campos$ + " ;Orden/A5" '8

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    CoordenadasUbicacion$ = RK & ";" & FL & ";" & KL
    CONDI$ = " COORDUBIC = '" & CoordenadasUbicacion$ & "'"
    SQLX$ = "SELECT IDSUBOR,CODXELEM,CODIELEM,IDENLOTE,NUMEOPER,DESCOPER,CANTREAL,NUORINF FROM INFOFAB WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Dim obj As New RECORXET
    Set rs1 = READSET(SQLX$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing

    FRMZELECHO.Caption = "Rack: " & RK & " Piso: " & FL & " Picking: " & KL
    FRMZELECHO.MSF1.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
    FRMZELECHO.MSF2.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
    FRMZELECHO.Width = 13000
    FRMZELECHO.Show 1
    
    SELECCIONLOTEENUBICACION$ = RESP_ZELE_VECT(1)

'    'SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED FROM UBILOTE_ENVAP
'    Campos$ = "Cta./F7;R.Social/A24;OP/f3;Desc./A12;O.F/a10;Lote/a12;Cant./F7.1;Ped./F7;ID/A0"
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'    Condi$ = " RACK = " & RK
'    Condi$ = Condi$ + " AND FILA = " & FL
'    Condi$ = Condi$ + " AND COLUMNA = " & KL
'
'    SQLX$ = "SELECT NUCLIEN,RASO_CLI,NUMEOPER,DESCOPER,IDSUBOR,LOTE,CANTIDAD,NROPED,UBILOTE_ENVAP_ID FROM UBILOTE_ENVAP WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE " & Condi$
'    Dim obj As New RECORXET
'    Set RS1 = READSET(SQLX$)
'    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
'    Set obj = Nothing
'    RS1.Close
'    Set RS1 = Nothing
'
'    FRMZELECHO.Caption = "Rack: " & RK & " Piso: " & FL & " Columna: " & KL
'    FRMZELECHO.MSF1.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
'    FRMZELECHO.MSF2.ToolTipText = "Doble Clic Para Borrar Lote en Ubicación Seleccionada"
'    FRMZELECHO.Width = 13000
'    FRMZELECHO.Show 1
'
'    SELECCIONLOTEENUBICACION$ = XVALO(RESP_ZELE_VECT(9))
End Function
Private Sub MSF2_Click()
    REG& = MSF2.MouseRow
    COL& = MSF2.MouseCol
    If REG& = 0 Or REG& >= MSF2.Rows Then Exit Sub
    RACK% = XVALO(MSF2.TextMatrix(REG&, 1))
    Label3 = RACK%

End Sub

Private Sub TEXT3_Change()
   Label7.Caption = TRIM$(ECOARCH$("!OPERINF1", MKI$(XVALO(Text3))))
End Sub

Private Sub Text3_DblClick()
    Call Command6_Click
End Sub

Private Sub Text4_DblClick()
    Call Command5_Click
End Sub

Private Sub Text5_DblClick()
    Call Command7_Click
End Sub

Private Sub Text7_LostFocus()
    Text7 = FORMATNUM$(XVALO(Text7), "F8.1")
End Sub

Sub ARMARGRILLAUBICA(NRACK%)
    ' **** SE ARMA LA GRILLA CON TODAS LAS UBICACIONES EN BLANCO ****
    MSF1.Rows = 1
    J& = 0
    VUELTA% = 0
    SQLX$ = "SELECT * FROM DATASQL WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE LISTA = 'CONFUBIX' AND CONVERT(int, campo1) > 0 "
    If NRACK% > 0 Then SQLX$ = SQLX$ + " AND CONVERT(int, campo1)= " & NRACK%
    SQLX$ = SQLX$ + " ORDER BY CAMPO1,CAMPO3,CAMPO4"
    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
         VUELTA% = VUELTA% + 1
         RACK% = XVALO(!CAMPO1)
         FIL% = XVALO(!CAMPO3)
         COL% = XVALO(!CAMPO4)
         For F% = 1 To FIL%
            For i% = 1 To COL%
               J& = J& + 1
               MSF1.Rows = J& + 1
               Call NEGRITA_SOLO_UNA_CELDA(MSF1, J&, 2)
               MSF1.TextMatrix(J&, 1) = XVALO(!CAMPO1)
               MSF1.TextMatrix(J&, 2) = SAFETEXT$(!campo2)
               MSF1.TextMatrix(J&, 3) = F% 'FILA
               MSF1.TextMatrix(J&, 4) = i% 'COLUMNA
               If (VUELTA% Mod 2) = 0 Then
                 Call COLOREAR_GRILLA_SOLAFILA(MSF1, &HE2D3C0, J&, , 1, 1)
               End If
             Next i%
         Next F%
         .MoveNext
       Loop
    End With
    rs.Close: Set rs = Nothing
    
    'FIN **** SE ARMA LA GRILLA CON TODAS LAS UBICACIONES EN BLANCO ****
    '
    Call MostrarUbicacionesGuardadas
'    'RECORRO LAS UBICACIONES ASIGNADAS Y LO IMPRIMO EN GRILLA
'    SQLX$ = "SELECT *  FROM UBILOTE_ENVAP WITH(NOLOCK)"
'    SQLX$ = SQLX$ + "ORDER BY RACK,FILA,COLUMNA,LOTE,IDSUBOR"
'    Set RS = READSET(SQLX$)
'    With RS
''    .CellForeColor = vbRed
'      Do While Not .EOF
'         RACK% = XVALO(!RACK)
'         FIL% = XVALO(!fila)
'         COL% = XVALO(!COLUMNA)
'         CAN = XVALO(!CANTIDAD)
'         LOTE$ = XVALO(!LOTE)
'         Call UBICAR_UBICACION_EN_GRILLA(Me.MSF1, RACK%, FIL%, COL%, CAN, LOTE$)
'        .MoveNext
'      Loop
'    End With
'    RS.Close: Set RS = Nothing
End Sub
Sub MostrarUbicacionesGuardadas()
    'RECORRO LAS UBICACIONES ASIGNADAS Y LO IMPRIMO EN GRILLA
    SQLX$ = "SELECT *  FROM INFOFAB WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COORDUBIC <> ''"
    SQLX$ = SQLX$ + " ORDER BY COORDUBIC"
    Set rs = READSET(SQLX$)
    With rs
'
      Do While Not .EOF
         COORDENADA$ = TRIM$(SAFETEXT$(!COORDUBIC))
         If XVALO(Left$(COORDENADA$, 1)) > 0 Then
             On Error Resume Next
             strArray = Split(COORDENADA$, ";")
             RK% = XVALO(strArray(0))
             FL% = XVALO(strArray(1))
             KL% = XVALO(strArray(2))
             If Err.Number = 0 Then
               On Error GoTo 0
               LOTE$ = SAFETEXT$(!idenlote)
               NORFA$ = SAFETEXT$(!idsubor)
               CAN = XVALO(!CANTREAL)
               NUORD% = XVALO(!NuOrInf)
               LblId1& = XVALO(!INFOFAB_ID)
               DESCOP$ = SAFETEXT$(!descoper)
               Call UBICAR_UBICACION_EN_GRILLA(Me.MSF1, RK%, FL%, KL%, CAN, LOTE$, NORFA$, NUORD%, LblId1&, DESCOP$)
             Else
               On Error GoTo 0
             End If
             
        End If
        .MoveNext
      Loop
    End With
    rs.Close: Set rs = Nothing
'ORIGINAL DESDE TABLA UBILOTE_ENVAP
'    SQLX$ = "SELECT *  FROM UBILOTE_ENVAP WITH(NOLOCK)"
'    SQLX$ = SQLX$ + "ORDER BY RACK,FILA,COLUMNA,LOTE,IDSUBOR"
'    Set RS = READSET(SQLX$)
'    With RS
''
'      Do While Not .EOF
'         RACK% = XVALO(!RACK)
'         FIL% = XVALO(!fila)
'         COL% = XVALO(!COLUMNA)
'         CAN = XVALO(!CANTIDAD)
'         LOTE$ = SAFETEXT$(!LOTE)
'         NORFA$ = SAFETEXT$(!IDSUBOR)
'         Call UBICAR_UBICACION_EN_GRILLA(Me.MSF1, RACK%, FIL%, COL%, CAN, LOTE$, NORFA$)
'        .MoveNext
'      Loop
'    End With
'    RS.Close: Set RS = Nothing
End Sub
Sub UBICAR_UBICACION_EN_GRILLA(MSF As msFlexGrid, RACK%, FIL%, COL%, CAN, LOTE$, NORFA$, NUORDEN%, Optional LblId1&, Optional DESCOP$)
    For i% = 1 To MSF.Rows - 1
       RACKENGRILLA% = XVALO(MSF.TextMatrix(i%, 1))
       If RACKENGRILLA% = RACK% Then
          FILAENGRILLA% = XVALO(MSF.TextMatrix(i%, 3))
          If FILAENGRILLA% = FIL% Then
             COLUENGRILLA% = XVALO(MSF.TextMatrix(i%, 4))
             If COLUENGRILLA% = COL% Then
               NROPEX& = NroPedidoSegunOf&(NORFA$)
               NC = ClienteSegunNroPedido(NROPEX&)

               CantiAsignar = CAN + XVALO(MSF.TextMatrix(i%, 5))
               LOTEENGRILLA$ = TRIM$(MSF.TextMatrix(i%, 6))
               PCOMA$ = ""
               If LOTEENGRILLA$ <> "" Then PCOMA$ = ";"
               LoteAsignar$ = LOTEENGRILLA$ & PCOMA$ & LOTE$
               MSF.TextMatrix(i%, 5) = CantiAsignar
               MSF.TextMatrix(i%, 6) = LoteAsignar$
               MSF.TextMatrix(i%, 7) = NORFA$
               
               If PCOMA$ = ";" Then 'SI YA TIENE INFO EN UNA MISMA UBICACION QUITA EL NRO DE INDICE Y PONE VALOR EN -1
                  NUORDEN% = -1
               End If
               MSF.TextMatrix(i%, 8) = NUORDEN%
               MSF.TextMatrix(i%, 9) = LblId1&
               MSF.TextMatrix(i%, 10) = NC
               MSF.TextMatrix(i%, 11) = rasocli$(NC)
               MSF.TextMatrix(i%, 12) = DESCOP$

               
               MSF.Row = i%
               MSF.COL = 5: MSF.CellForeColor = vbRed
               MSF.COL = 6: MSF.CellForeColor = vbRed
               MSF.COL = 7: MSF.CellForeColor = vbRed
               Exit For
             End If
          End If
       End If
    Next i%
End Sub
Sub BLANFORMU()
   '
    Label4 = ""
    Text3 = ""
    
    TXTCODELEM = ""
    TXTCORTE = ""
    TXTLOTE = ""
    TXTCANTICORTES = ""
   '
End Sub

