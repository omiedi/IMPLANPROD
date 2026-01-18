VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ASINUMSE 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Números de Serie a Pedidos"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE DE TRAZABLES POR LOTE  ( '*' Items NO Trazables)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   25
      Top             =   4800
      Width           =   11805
      Begin VB.CommandButton Command5 
         Caption         =   "Transferir Items No Trazables"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8160
         TabIndex        =   38
         ToolTipText     =   "Solo Transfiere los Items No Trazables, Los Trazables se Transfieren al Asignarle el Lote"
         Top             =   2760
         Width           =   3540
      End
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2400
         Left            =   30
         TabIndex        =   26
         ToolTipText     =   "Doble Clic Para Asignar Lote\Editar Cantidad Para Items NO Trazables"
         Top             =   240
         Width           =   11710
         _ExtentX        =   20664
         _ExtentY        =   4233
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   13490906
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   30
         Top             =   2850
         Width           =   810
      End
      Begin VB.Label Label14 
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
         Left            =   1215
         TabIndex        =   29
         Top             =   2760
         Width           =   1890
      End
      Begin VB.Label Label13 
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
         Left            =   3195
         TabIndex        =   28
         Top             =   2760
         Width           =   4920
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   27
         Top             =   2760
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   180
      Index           =   0
      Left            =   2520
      ScaleHeight     =   120
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   300
      Left            =   11280
      TabIndex        =   0
      Top             =   50
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   0
      TabIndex        =   9
      Top             =   3195
      Width           =   11805
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   9360
         TabIndex        =   32
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   8985
         TabIndex        =   31
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Asignar y Transferir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9360
         TabIndex        =   2
         Top             =   720
         Width           =   2340
      End
      Begin VB.CommandButton Command4 
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
         Height          =   270
         Left            =   3480
         TabIndex        =   1
         Top             =   720
         Width           =   175
      End
      Begin VB.Label Label6 
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
         Height          =   300
         Left            =   9720
         TabIndex        =   37
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   1920
         TabIndex        =   36
         Top             =   1080
         Width           =   1860
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   1440
         TabIndex        =   35
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8115
         TabIndex        =   34
         Top             =   1185
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "De Depósito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   400
         TabIndex        =   33
         Top             =   1185
         Width           =   915
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   17
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label10 
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
         Left            =   4200
         TabIndex        =   16
         Top             =   315
         Width           =   7485
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   15
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label8 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   315
         Width           =   2010
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
         Left            =   1440
         TabIndex        =   3
         Top             =   705
         Width           =   2025
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Número Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   10
         Top             =   800
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE DE TRAZABLES POR N° DE SERIE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   11805
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2400
         Left            =   30
         TabIndex        =   24
         Top             =   240
         Width           =   11710
         _ExtentX        =   20664
         _ExtentY        =   4233
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   13490906
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8280
      Left            =   11880
      ScaleHeight     =   8250
      ScaleWidth      =   1590
      TabIndex        =   4
      Top             =   0
      Width           =   1620
      Begin VB.CommandButton Command6 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   7200
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   110
         ScaleHeight     =   165
         ScaleWidth      =   585
         TabIndex        =   21
         Top             =   4560
         Width           =   620
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3840
         Width           =   600
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Desasignar N.Serie a Pedido"
         Top             =   3120
         Width           =   600
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Acceso Directo a Transferencia de Equipos"
         Top             =   2370
         Width           =   600
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consulta de Asignacion de N/Serie"
         Top             =   1620
         Width           =   600
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "ASINUMSE.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   600
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ASINUMSE.frx":0FF4
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "ASINUMSE.frx":1228
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label28 
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
      Height          =   300
      Left            =   9960
      TabIndex        =   8
      Top             =   45
      Width           =   1320
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Pedido"
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
      Left            =   8295
      TabIndex        =   7
      Top             =   150
      Width           =   1545
   End
   Begin VB.Menu mnuArchivo1 
      Caption         =   "Archivo"
      Begin VB.Menu mnuIngresoManualSerie 
         Caption         =   "Ingreso de N.Serie Forma Manual"
         Visible         =   0   'False
      End
      Begin VB.Menu S1 
         Caption         =   "_"
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu S 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesasignar 
         Caption         =   "Desasignar Lotes a Item Seleccionado"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "ASINUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_ASINUMSE$
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim MSFGLOBALFRM As msFlexGrid



Sub GrabarLotesPedido(NP&, ORDEX%, CI1%)

    CONDI$ = "NROPED = " & NP& & " AND NuOrItem = " & ORDEX% & " AND CODIINT = " & CI1%
    'LE GRABO UN IDENLOTE PARA IDENTIFICARLO, CUANDO LO TENGA QUE ELIMINAR.
    Call ACTUREGISTRO("PEDDETA", "IDENLOTE", CONDI$, "ABCDE99999", REGAF&)
    '
    CONDI$ = CONDI$ + " AND IDENLOTE = 'ABCDE99999'"
    'INSERTO LOS NUEVOS REGISTROS UNO PARA CADA ASIGNACION DE LOTE
    For i& = 1 To MSF2.Rows - 1
      If MSF2.TextMatrix(i&, 7) = ORDEX% Then
       CANTIX$ = MSF2.TextMatrix(i&, 4)
       CODI$ = MSF2.TextMatrix(i&, 2)
       LOTEX$ = MSF2.TextMatrix(i&, 6)
       SQLX$ = " INSERT INTO PEDDETA(PreList,CodiEmpr,CanPed,PreUnid,ImpoTot,ImpoAnti,CantEnt,CanSaldo,FeUltEnt,TPedItem,NroPed,TipoPed,FechEmis,NroClie,RaSoClie,NroOcom,FeSolEnt,FinVigen,Status,ListPre,CondPag,Vended,PorComVend,NuOrItem,CodiInt,CodiExt,Descrip,VenRep,MoneEmis,TipoCam,DescLista,NumeroSerie,IdenLote,LISTA_DE_ITEM,SIN_CARGO,CANASIG)"
       SQLX$ = SQLX$ + " SELECT PreList,CodiEmpr, '" & CANTIX$ & "',PreUnid,ImpoTot,ImpoAnti,CantEnt,'" & CANTIX$ & "',FeUltEnt,TPedItem,NroPed,TipoPed,FechEmis,NroClie,RaSoClie,NroOcom,FeSolEnt,FinVigen,Status,ListPre,CondPag,Vended,PorComVend,NuOrItem,CodiInt,CodiExt,Descrip,VenRep,MoneEmis,TipoCam,DescLista,NumeroSerie,'" & LOTEX$ & "',LISTA_DE_ITEM,SIN_CARGO,'" & CANTIX$ & "' FROM PEDDETA "
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
       TXTX$ = "Asignación de Lote a Código: " & CODI$ & vbCrLf
       TXTX$ = "Lote: " & LOTEX$
       Call AGREGA_ANOTAPED_NEW(NP&, TXTX$)
      End If
   Next i&
   'BORRO EL REGISTRO ORIGINAL DEL PEDDETA.
   Call BORRAREGISTROS("PEDDETA", CONDI$, REGAF&, "NOMESS")
End Sub

Sub DesasignaLotesPedido(NP&, ORDEX%, CI1%)
    'BORRA LOS REGISTROS Y ACUMULA EN 1 EL TOTAL PARA EL CODIGO
    
    'FUNCIONAMIENTO
    'RECORRE LOS REGISTROS SEGUN CONDICION
    'POR EJEMPLO SI LOS REGISTROS SON 3 ACUMULA LA CANTIDAD DE LOS PRIMEROS 2 Y LOS BORRA
    'CUANDO LLEGA AL TERCERO Y ULTIMO LE ASIGNA LA CANTIDAD ACUMULADA DE LOS PRIMEROS 2 + LA DEL REGISTRO 3
    'LIMPIA EL CASILLERO IDENLOTE, QUEDANDO UN SOLO REGISTRO CON LA CANTIDAD PEDIDA, ENTREGADA Y SALDO ORIGINAL Y SIN LOTE.
    
    CONDI$ = "NROPED = " & NP& & " AND NuOrItem = " & ORDEX% & " AND CODIINT = " & CI1%
    CANTREG% = CantRegistros("PEDDETA", CONDI$, "NOMESS")
    If CANTREG% < 1 Then Exit Sub
    '
    J& = 0
    SQLX$ = "SELECT * FROM PEDDETA WHERE " & CONDI$
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
      Do While Not .EOF
         J& = J& + 1
         If J& < CANTREG% Then   'LEO LAS CANTIDADES Y BORRO LOS PRIMEROS REGISTROS ANTES DE LLEGAR AL ULTIMO
            CANTIPED& = (CANTIPED& + XVALO(!CanPed))
            CANTIENT& = (CANTIENT& + XVALO(!CANTENT))
            .Delete
         Else 'CUANDO LLEGA AL ULTIMO ENTRA POR ACA
            !idenlote = ""
            CANTIORIGINAL = XVALO(!CanPed) + CANTIPED&
            TOT_CANTIENTREGADA& = CANTIENT& + XVALO(!CANTENT)
            If CANTIORIGINAL < 0.005 Then CANTIORIGINAL = 0
            SALDOAENTREGAR = CANTIORIGINAL - TOT_CANTIENTREGADA&: If SALDOAENTREGAR < 0.005 Then SALDOAENTREGAR = 0
            !CANTENT = TOT_CANTIENTREGADA&
            !CanPed = CANTIORIGINAL
            !CanSaldo = SALDOAENTREGAR
            !CanAsig = 0
            .Update
         End If
      
        .MoveNext
      Loop
    End With
    rs.Close: Set rs = Nothing
    TXTX$ = "DesAsignación de Lote a Código: " & CODI$
    Call AGREGA_ANOTAPED_NEW(NP&, TXTX$)


End Sub

Private Sub Command1_Click()
     '
     NPSEL& = XVALO(Label28)
     Me.MSF.Enabled = False
     Call SELECHO("INDIANSE")
     '
     Label28 = XVALO(VALACT1("INDIANSE"))
     If Label28 = 0 Then Label28 = NPSEL&
     
     Me.MSF.Enabled = True
     '
     'VALIDO SI EL PEDIDO TIENE MOTIVO DE VENTA
     If VALIDAOV%(XVALO(Label28)) < 1 Then
       Call BLANFORMU
       Exit Sub
     End If
     '
     TXTRAZA_ASINUMSE$ = ""
     TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 1 = " & Now & vbCrLf
     Call CARLISPED 'Carga la lista con el pedido seleccionado
     Me.MSF.Enabled = True
     TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 2 = " & Now & vbCrLf
     '
End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command11_Click()
     '
     Call CIERRARCH("*.*")
     Unload Me
     '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text1.TEXT = VDEV$
    End If
    '
End Sub

Private Sub Command23_Click()
    '
    OPX% = ALTERNA%("Pedidos Pendientes de Asignacion\Asignacion Por Numero de Serie\Numero de Serie por Deposito\Consulta por Numero de Serie")
    '
    Select Case OPX%
        Case 1: Call FINAL("?PEDPENASIG")
        Case 2: Call FINAL("?ASIGNUMSE")
        Case 3: Call FINAL("?NUMSEPORDEP")
        Case 4: Call FINAL("?TODOSNUMEROS")
    End Select
    '
End Sub

Private Sub Command26_Click()
    '
    FSTASNUSE.Show 1
    '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    TRANEQUI07.Show 1
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NPED& = XVALO(Label28)
    'ORDENITEM% = XVALO(Mid$(List1.TEXT, 70, 6))
    ORDENITEM% = XVALO(Label11)
    Ci% = CODINT%(Label8)
    NUMSE$ = TRIM$(Label1)
    DEPODE% = XVALO(Text1)
    DEPORI% = XVALO(Label4)
    '
    If Ci% < 1 Then
         Call MENSERR(24, "Codigo No Válido")
         GoTo 99
      ElseIf NUMSE$ = "" Then
         Call MENSERR(24, "Falta Numero de Serie.")
         GoTo 99
      ElseIf DEPODE% < 1 Then
         Call MENSERR(24, "Falta Depósito Destino.")
         GoTo 99
      ElseIf ECOARCH$("TADEPOSI", Chr$(DEPODE%)) = "" Then
         Call MENSERR(24, "Depósito Destino no Válido.")
         GoTo 99
      ElseIf DEPORI% < 1 Then
         Call MENSERR(24, "Falta Depósito Origen.")
         GoTo 99
      ElseIf ECOARCH$("TADEPOSI", Chr$(DEPORI%)) = "" Then
         Call MENSERR(24, "Depósito Origen no Válido.")
         GoTo 99
    End If
    '
    TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 3 = " & Now & vbCrLf

    Call TRANEQUI
    '

    Call PRINTASINUMSE
    '
    TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 17 = " & Now & vbCrLf
    
'    TX$ = List1.TEXT
'    i& = List1.ListIndex
'    Call REPLA(TX$, NUMSE$, 48, 16)
'    List1.List(i&) = TX$
'    List1.Refresh
    Call SETULTREG("!INUSEDIS", 0)
    Label1 = ""
    Label11 = ""
    Label4 = ""
    Label8 = ""
    Label10 = ""
    Text1 = ""
    Call CARLISPED
    
'GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TXTRAZA_ASINUMSE-" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_ASINUMSE$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Ci% = CODINT(Label8)
    jj& = 0
    If Ci% > 0 Then
      If ESTRAZABLE(Ci%) < 2 Then
         Call MENSERR(24, "Código '" & CODEXT(Ci%) & "' No Trazable. Verifique !")
         Exit Sub
      End If
      '
      Campos$ = "N° de Serie/A18;Código/A20;Descripción/A48;Dep./F4"
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
      FRMZELECHO.Caption = "N° Series Disponibles"

      Dim obj As New RECORXET
      Set rs1 = obj.RS_SERIESDISPONIBLES(Ci%)
      Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
      Set obj = Nothing
      rs1.Close
      Set rs1 = Nothing
     '
      FRMZELECHO.Width = 12000
      FRMZELECHO.Show 1
      NSERIE$ = RESP_ZELE_VECT(1)
      Label1 = NSERIE$
      Label4.Caption = XVALO(RESP_ZELE_VECT(4))
      If XVALO(Label4.Caption) < 1 And NSERIE$ <> "" Then
        Call RECUDEPO(NSERIE$, OKX%)
        If OKX% < 1 Then Call FINAL("")
        Call COMUNI("Recuperado Depósito del Equipo")
        Label4.Caption = XVALO(VALOBADA("TANUMSE", "DepoNumse", "NumeroSerie = '" & NSERIE$ & "'"))
      End If
   End If
'      If ULTREG&("!INUSEDIS") > 0 Then
'
'         Call SELECHO("!INUSEDIS")
'         Me.MSF2.Enabled = False
'         Me.MSF.Enabled = False
'         TTXX$ = TRIM$(VALACT1$("!INUSEDIS"))
'         If TTXX$ <> "" Then
'            Label1 = TTXX$
'15          Label4.Caption = XVALO(VALOBADA("TANUMSE", "DepoNumse", "NumeroSerie = '" & TTXX$ & "'"))
'            If XVALO(Label4.Caption) < 1 Then
'               Call RECUDEPO(TTXX$, OKX%)
'               If OKX% < 1 Then Call FINAL("")
'               Call COMUNI("Recuperado Depósito del Equipo")
'               GoTo 15
'            End If
'         End If
'         Me.MSF2.Enabled = True
'         Me.MSF.Enabled = False
'      Else
'         Call MENSERR(24, "Código " & CODEXT$(Ci%) & " Sin Números de  Serie para Asignar")
'      End If
'    End If
End Sub

Sub RECUDEPO(NUSEX$, OKX%)
    OKX% = 0
    NORFAORI$ = SAFETEXT$(VALOBADA("TANUMSE", "IdSubOr", "NumeroSerie = '" & NUSEX$ & "'", "NOMESS"))
    If NORFAORI$ = "" Then
       Call MENSERR(24, "Imposible Recuperar - Consulte")
       Exit Sub
    End If
    '
    DEPOCI = XVALO(VALOBADA("MOVISTO", "DEPOMOVI", "IDENLOTE = '" & NORFAORI$ & "' AND CantIngre >0 AND DepoMovi > 0"))
    If DEPOCI < 1 Then DEPOCI = XVALO(Control$("", "DEPOCIE"))
    '
    Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", "IDENLOTE = '" & NORFAORI$ & "' AND CantIngre >0 AND DepoMovi = 0", DEPOCI, RAFE&)
    Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "IDSUBOR = '" & NORFAORI$ & "' AND DEPONUMSE = 0", DEPOCI, RAFE&)
    If RAFE > 0 Then OKX% = 1
    '
End Sub
'
Private Sub Command43_Click()

      Call COPYSTRU("INDINUSE", "DESAGNUM")
      TTXX$ = REGBLAN$("DESAGNUM")
      Call SETULTREG("DESAGNUM", 0)
      '
10    Call SELECHO("MASTER")
      Ci% = CODINT%(VALACT1$("MASTER"))
      If Ci% < 1 Then Exit Sub
      If ESTRAZABLE(Ci%) < 2 Then
         Call MENSERR(24, "Código '" & CODEXT(Ci%) & "' No Trazable. Verifique !")
         Exit Sub
      End If
      jj& = 0
      SQLX$ = "SELECT * FROM TANUMSE "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & Ci% & ""
      SQLX$ = SQLX$ + " AND (NumRemito = '' OR NumRemito IS NULL)"
      Dim NUSERTMP As ADODB.Recordset
      Set NUSERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With NUSERTMP
         Do While Not .EOF
         NUSERR$ = SAFETEXT$(!NUMEROSERIE)
         NPED& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUSERR$ & "' AND Status < 8 "))
         NusePed$ = TRIM$(Str(NPED&))
         If NPED& > 0 Then
              Ci% = XVALO(!cod_inte)
              Call REPLA(TTXX$, NUSERR$, 1, 16)
              Call REPLA(TTXX$, CODEXT$(Ci%) + "  " + Left$(DESCRIT(Ci%), 20) + " - NºPed. " & NusePed$, 17, 47)
              jj& = jj& + 1
              Call GRAREG("DESAGNUM", TTXX$, jj&)
         End If
         .MoveNext
         Loop
      End With
      NUSERTMP.Close
      Set NUSERTMP = Nothing
      Call SETULTREG("DESAGNUM", jj&)
      Call FRESHECHO("DESAGNUM")
      Call CIERRARCH("DESAGNUM")

      If ULTREG&("DESAGNUM") > 0 Then
        Call SELECHO("DESAGNUM")
        NUMSE$ = TRIM$(VALACT1$("DESAGNUM"))
        TXTX$ = TRIM$(VALACT2$("DESAGNUM"))
        If NUMSE$ = "" Then GoTo 10
        '
        NN% = InStr(1, TXTX$, "NºPed.") + 6
        NPED& = XVALO(Mid$(TXTX$, NN%))
        CONDI$ = "NUMEROSERIE = '" & NUMSE$ & "'"
        Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REGAF&, "NOMESS") ' desasigna del pedido
        Call ACTUREGISTRO("PEDDETA", "NroRemi", CONDI$, "", REGAF&, "NOMESS")
        '
        HOII% = HOY(HOS$)
        TTXX$ = Space$(96)
        Call REPLA(TTXX$, HOS$ + "  " + "Desasignación de N/Serie", 1, 56)
        Call REPLA(TTXX$, USERTER$, 69, 24)
        Call REPLA(TTXYZ$, " - Código: ", 10, 11)
        Call REPLA(TTXYZ$, CODEXT$(Ci%), 22, 16)
        Call REPLA(TTXYZ$, "- N/Serie: ", 39, 11)
        Call REPLA(TTXYZ$, NUMSE$, 51, 12)
        TTXYZ1$ = TTXYZ1$ + TTXYZ$ + Chr$(10) + Chr$(13)
        Call AGREGA_ANOTAPED(NPED&, TTXX$)
        TTXYZ1$ = Left$(TTXYZ1$, (Len(TTXYZ1$) - 2))
        Call AGREGA_ANOTAPED(NPED&, TTXYZ1$)
        Call AGREGA_ANOTAPED(NPED&, String$(88, "="))
        '
        'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
        CUERPO$ = "Desasignación de Pedido: " & NPED&
        Call ANOTANUSE(NUMSE$, CUERPO$)
        '
        Call MENSERR(24, "Número de Serie Desasignado")
        '
        NPSEL& = XVALO(ASINUMSE.Label28)
        '
        Call CARLISPED 'Carga la lista con el pedido seleccionado
        '
      Else
        Call MENSERR(24, "No Hay Números de Serie para Desasignar")
         
      End If
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Screen.MousePointer = 11
   
   Call ABREPEDCLI
    '
   TTXYZ1$ = ""
   VUELTA% = 0
   For i& = 0 To List1.ListCount - 1
    '
    If XVALO(Mid$(List1.List(i&), 78, 1)) > 0 Then
       '
       NPED& = XVALO(Label28)
       ORDENITEM% = XVALO(Mid$(List1.List(i&), 70, 6))
       Ci% = CODINT%(TRIM$(Mid$(List1.List(i&), 1, 16)))
       NUMSE$ = TRIM$(Mid$(List1.List(i&), 48, 16))
       DEPONU% = XVALO(Mid$(List1.List(i&), 80, 2))
       DepOrig% = XVALO(Mid$(List1.List(i&), 83, 2))
       '
       CONDICION$ = " NroPed = " & NPED& & " AND CodiInt = " & Ci% & " AND NuOrItem = " & ORDENITEM%
       '
       If NUMSE$ = "" Then GoTo 10 ' SI DESASIGNA NO VALIDA
       Valiped& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUMSE$ & "'"))
       If Valiped& = 0 Then ' validacion que otro usuaario no haya utilizado el mismo numero de serie
10       VUELTA% = VUELTA% + 1
         Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDICION$, NUMSE$, REGAF&, "NOMESS")
         If DEPONU% > 0 Then
            '
            FECHH% = HOY(HO$)
            NUMCOM$ = ProNroComprobante("TR")
            NUMCOM$ = "TR." + TRIM$(NUMCOM$)
            REF$ = "Transferencia de Deposito"
            '
            'Registra la salida del articulo del deposito original
            Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DepOrig%, (-1) * 1, NPED&)
            'Registra el ingreso del articulo al deposito nuevo
            Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DEPONU%, 1, NPED&)
            '
            'REALIZA LA TRANSFERENCIA ACTUALIZANDO EL DEPOSITO EN TANUMSE
            Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE = '" & NUMSE$ & "'", DEPONU%, REGAF&)
            '
         End If
         '
         Call REPLA(TTXYZ$, " - Código: ", 10, 11)
         Call REPLA(TTXYZ$, CODEXT$(Ci%), 22, 16)
         Call REPLA(TTXYZ$, "- N/Serie: ", 39, 11)
         If TRIM$(NUMSE$) = "" Then NUMSE$ = "S/N"
         Call REPLA(TTXYZ$, NUMSE$, 51, 12)
         TTXYZ1$ = TTXYZ1$ + TTXYZ$ + Chr$(10) + Chr$(13)
         '
         TX$ = List1.List(i&) ' pongo en 0 los items leidos para evitar doble grabacion
         Call REPLA(TX$, "0", 78, 1)
         List1.List(i&) = TX$
       Else
         Call MENSERR(24, "El Número de Serie: " & NUMSE$ & " Ocupado en Pedido: " & Valiped&)
       End If
    End If
    '
   Next i&
   If VUELTA% > 0 Then ' SI GENERO ALGUNA MODIFICACION ANOTA EN EL ANOTADOR DE PEDIDOS
     HOII% = HOY(HOS$)
     TTXX$ = Space$(96)
     Call REPLA(TTXX$, HOS$ + "  " + "Asignación de N/Serie", 1, 56)
     Call REPLA(TTXX$, USERTER$, 69, 24)
     Call AGREGA_ANOTAPED(NPED&, TTXX$)
     TTXYZ1$ = Left$(TTXYZ1$, (Len(TTXYZ1$) - 2))
     Call AGREGA_ANOTAPED(NPED&, TTXYZ1$)
     Call AGREGA_ANOTAPED(NPED&, String$(88, "="))
     Call MENSERR(24, "Números de Series Asignados")
   End If
    '
   Command7.Enabled = True
   Screen.MousePointer = 1
  '
End Sub

Sub TRANEQUI()
   '
   Screen.MousePointer = 11
   'Call ABREPEDCLI
   Call ABRECONEXION
   '
   TTXYZ1$ = ""
   VUELTA% = 0
   
   NPED& = XVALO(Label28)
   Ci% = CODINT%(Label8)
   NUMSE$ = TRIM$(Label1)
   DEPONU% = XVALO(Text1)
   DepOrig% = XVALO(Label4)
   NUORD% = XVALO(Label11)
   '
   CONDICION$ = " NroPed = " & NPED&
   CONDICION$ = CONDICION$ & " AND CodiInt = " & Ci%
   CONDICION$ = CONDICION$ & " AND NuOrItem = " & NUORD% & ""
   CONDICION$ = CONDICION$ & " AND (NUMEROSERIE = '' or NUMEROSERIE IS NULL)"
   '
   Valiped& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUMSE$ & "'"))
   If Valiped& > 0 Then ' validacion que otro usuaario no haya utilizado el mismo numero de serie
       Call MENSERR(24, "Imposible Completar Transaccion Solicitada.\  \El Número de Serie: " & NUMSE$ & " Ocupado en Pedido: " & Valiped&)
       Call FINAL("")
   End If
   '
   'Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDICION$, NUMSE$, REGAF&, "NOMESS")
   SQLX$ = "SELECT NUMEROSERIE FROM PEDDETA WHERE " + CONDICION$
   Dim rstPEDDETA As New ADODB.Recordset
   Set rstPEDDETA = New ADODB.Recordset
25 On Error Resume Next
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 4 = " & Now & vbCrLf
   rstPEDDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 5 = " & Now & vbCrLf

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
   End If
   On Error GoTo 0
   '
   With rstPEDDETA
     If Not (.BOF And .EOF) Then
        !NUMEROSERIE = NUMSE$
        .Update
     End If
   End With
   rstPEDDETA.Close
   Set rstPEDDETA = Nothing
   '
'   If REGAF& < 1 Then
'      Call MENSERR(24, "La Transaccion no Pudo Completarse - Consulte !")
'      Call FINAL("")
'   End If
   '
   FECHH% = HOY(HO$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 6 = " & Now & vbCrLf
   NUMCOM$ = ProNroComprobante("TR")
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 7 = " & Now & vbCrLf
   NUMCOM$ = "TR." + TRIM$(NUMCOM$)
   REF$ = "Transferencia de Deposito"
   '
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 8 = " & Now & vbCrLf
   'Registra la salida del articulo del deposito original
   Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DepOrig%, (-1) * 1, NPED&)
   'Registra el ingreso del articulo al deposito nuevo
   Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DEPONU%, 1, NPED&)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 9 = " & Now & vbCrLf
   '
   'REALIZA LA TRANSFERENCIA ACTUALIZANDO EL DEPOSITO EN TANUMSE
   '
   'Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE = '" & NUMSE$ & "'", DEPONU%, REGAF&)
   SQLX$ = "SELECT DEPONUMSE FROM TANUMSE WHERE NUMEROSERIE = '" & NUMSE$ & "'"
   Dim rstTANUMSE As New ADODB.Recordset
   rstTANUMSE.CursorLocation = adUseServer
   Set rstTANUMSE = New ADODB.Recordset
35 On Error Resume Next
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 10 = " & Now & vbCrLf

   rstTANUMSE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 11 = " & Now & vbCrLf
   
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 35
   End If
   On Error GoTo 0
   '
   With rstTANUMSE
     If Not (.BOF And .EOF) Then
        !DepoNumSe = DEPONU%
        .Update
     End If
   End With
   rstTANUMSE.Close
   Set rstTANUMSE = Nothing
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 12 = " & Now & vbCrLf
   '
   If TRIM$(NUMSE$) = "" Then NUMSE$ = "S/N"
   TTXYZ$ = "-Código: "
   TTXYZ$ = TTXYZ$ + CODEXT$(Ci%)
   TTXYZ$ = TTXYZ$ + " - N/Serie: "
   TTXYZ$ = TTXYZ$ + " " & NUMSE$
   
   TTXX$ = "Asignación de N/Serie" & vbCrLf
   TTXX$ = TTXX$ & TTXYZ$
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 13 = " & Now & vbCrLf
   Call AGREGA_ANOTAPED_NEW(NPED&, TTXX$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 14 = " & Now & vbCrLf
   
   'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
   '
   CONDI$ = " NroPed = " & NPED& & " And CodiInt = " & Ci% & " And NUMEROSERIE = '" & NUMSE$ & "'"
   NCLIE = XVALO(VALOBADA("PEDDETA", "NROCLIE", CONDI$, "NOMESS"))
   CONDI$ = " NroPed = " & NPED&
   NUMOT% = NUMOTIVO%(NPED&)
   '
   CUERPO$ = "-Asignación de Pedido: " & NPED&
   If NUMOT% > 0 Then
     DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
     CUERPO$ = CUERPO$ & " Motivo: " & NUMOT% & "(" & TRIM$(DESCMOTI$) & ")"
   End If
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 15 = " & Now & vbCrLf
   CUERPO$ = CUERPO$ & " - Cliente: " & NCLIE
   CUERPO$ = CUERPO$ & " - " & rasocli(NCLIE)
   Call ANOTANUSE(NUMSE$, CUERPO$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 16 = " & Now & vbCrLf
   Screen.MousePointer = 1

End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   NUPED& = XVALO(Label28)
   'VALIDO NUEVAMENTE QUE HAYA STOCKS DE LOS ITEMS QUE SE VAN A TRANSFERIR
   Dim STOCK(32767)
   DEPOPAÑOL1% = XVALO(Control$("", "DEPOCEN"))
   HAYCANTIDADES = 0
   For i& = 1 To MSF2.Rows - 1
      CI1% = MSF2.TextMatrix(i&, 1)
      If ESTRAZABLE%(CI1%) < 1 Then
         STOCK(CI1%) = STODEPOS(CI1%, DEPOPAÑOL1) 'transfiere pero sin lote utilizo la misma aprovechando su amplio funcionamiento
         CANTIATRANSFERIR = XVALO(MSF2.TextMatrix(i&, 10))
         If CANTIATRANSFERIR <= 0 Then GoTo 20
         
         If CANTIATRANSFERIR > STOCK(CI1%) Then
            Call COMUNI("Atención!!! Falta de Stock Para Items Seleccionado")
            Call ResaltarFila(Me.MSF2, i&)
            GoTo 99
         Else
            HAYCANTIDADES = 1
         End If
      End If
20 Next i&
   If HAYCANTIDADES <= 0 Then
      Call COMUNI("No Hay Cantidades a Transferir informadas en Grilla")
      GoTo 99
   End If
   '
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   For i& = 1 To MSF2.Rows - 1
      ORD% = MSF2.TextMatrix(i&, 7)
      CI1% = MSF2.TextMatrix(i&, 1)
      If ESTRAZABLE%(CI1%) < 1 Then
        Call TransfReservaPedidoConLote(NUPED&, ORD%, CI1%, 0, 1) 'transfiere pero sin lote utilizo la misma aprovechando su amplio funcionamiento
      End If
   Next i&
   
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Screen.MousePointer = 1
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Screen.MousePointer = 1
      Call COMUNI("Transferencia Generada Correctamente Solo Para Items No Trazables")
      Call CARLISPED
    End If
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   FIKARDEX09.Text1 = Label14
   FIKARDEX09.Show 1
   Command6.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Height = 5500
    UTILIZA_SKIN% = 1
    INCLUYEITEMNOTRAZABLES% = XVALO(Control("ASINUMSE", "INCNOTRA"))
    Call APLICACIONSKINS(Me, Picture14)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "EYM") > 0 Then
       mnuIngresoManualSerie.Visible = True
    End If
    Campos$ = "CI/A0;Código/A20;Descripción/A30;Número de Serie/A16;Lote/A16;NORITEM/A0;;A Transf./F0"
    Call CARGA_ENCABEZADO(MSF, Campos$, ASINUMSE)
    
    Campos$ = "CI/A0;Código/A14;Descripción/A28;Cant./F8.1;C.Asig./F7.1;Lote/A16;NORITEM/A0;St.Disp/F0;A Transf/F0"
    If INCLUYEITEMNOTRAZABLES% > 0 Then
       Campos$ = "CI/A0;Código/A14;Descripción/A28;Cant./F7.1;C.Asig./F7.1;Lote/A12;NORITEM/A0;St.Disp/F9.1;Pend.Disp/F8.1;A Transf./F8.1"
       'DECLARO TOOLTIPTEXT A VISUALIZAR SEGUN COLUMNA EN LA GRILLA DE FRM GRILLAGRAL (VectorToolTipText ES UNA VARIABLE PUBLICA)
       Erase VectorToolTipText 'EL INDICE ES LA COLUMNA DONDE SE APLICA EL TOOLTIPTEXT
       VectorToolTipText(4) = "Cantidad Pedida Original"
       VectorToolTipText(5) = "Cantidades Asignada y Transferida "
       VectorToolTipText(8) = "Stock Disponible en Depósito Central (Pañol)"
       VectorToolTipText(9) = "Cantidades Disponible Pendiente a Transferir (si la cantidad pendiente supera la del stock cantidad pendiente = stock)"
       VectorToolTipText(10) = "Cantidades Sugerida a Transferir (Editable Con doble click Sobre la Fila)"
       
    End If
    
    Call CARGA_ENCABEZADO(MSF2, Campos$, ASINUMSE)

    If Control("PEDIDO", "ASILOTPE") = "SI" Then
        Me.Height = 8685
    End If
End Sub
Sub CARGALOTE_PEDIDO(NUPED&, REG&, ORD%)
    '
    DEPONU% = XVALO(Control$("", "DPREPPAÑ")) ' DEPOSITO 44 DE MEDITEA CONFIGURADO EN SETUP DE VENTAS  'Pañol/Factur'
    DepOrig% = XVALO(Control$("", "DEPOCEN")) ' DEPOSITO 50 PAÑOL DE MEDITEA CONFIGURADO EN SETUP DE FABRICACION/VALEMAT 'deP/Pañol'
    If DEPONU% < 1 Then Call COMUNI("Fala Configurar Depósito Utilizado como Depósito de Reposición de Facturación\El mismo se Configura Desde el Setup de Facturación (Pañol/Factur) "): Exit Sub
    
    If DepOrig% < 1 Then Call COMUNI("Fala Configurar Depósito Pañol\El mismo se Configura Desde el Setup de Fabricación (Dep.Pañol) "): Exit Sub
    
    'CARGA LA INFORMACION PARA ASIGNAR LOTES A LOS ITEMS DE UN PEDIDO.
    LOTESPED.List2.Clear
    'CAMPOS$ = "CI/A0;Código/A20;Descripción/A30;Cant./F8.1;Número de Serie/A16;Lote/A16;NORITEM/A0"
    'For U& = 1 To ASINUMSE.MSF2.Rows - 1
      CAN = ASINUMSE.MSF2.TextMatrix(REG&, 4)
      If CAN > 0.05 Then
        '
        CIXI% = XVALO(ASINUMSE.MSF2.TextMatrix(REG&, 1))
        If CIXI% > 0 Then
            NUORIT% = XVALO(ASINUMSE.MSF2.TextMatrix(REG&, 7))
            LOTEX$ = ASINUMSE.MSF2.TextMatrix(REG&, 6)
            '
            Call REPLA(TX$, CODEXT$(CIXI%), 1, 16) 'OJO-LONGITUD DE CODIGO EXTERNO
            Call REPLA(TX$, DESCRIT$(CIXI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN, "F11.1"), 50, 11)
            Call REPLA(TX$, LOTEX$, 62, 16)
            Call REPLA(TX$, SAFETEXT$(ORD%), 132, 3)
            LOTESPED.List2.AddItem TX$
         End If
        End If
   ' Next U&
    '
   'Presentar formulario para asignar lotes
   If LOTESPED.List2.ListCount > 0 Then
     LOTESPED.Label16 = NUPED&
     LOTESPED.Label7 = ""
30   LOTESPED.Show 1
     If LOTESPED.List2.ListCount < 1 Or LOTESPED.Label7 = "FALTAN_LOTES" Then
          LOTESPED.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         'PRIMERO VALIDO QUE TODOS TENGAN LOTE
         For i& = 1 To LOTESPED.List2.ListCount
           LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(i& - 1), 62, 16))
           Ci% = CODINT%(Mid$(LOTESPED.List2.List(i& - 1), 1, 16))
           If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote Para el Código: " & CODEXT(Ci%))
              GoTo 30 ' PRESENTA NUEVAMENTE EL LOTESPED
           End If
         Next i&

         J& = 0
        '
        If OK% >= 0 Then
          OPX% = COMALTER%("Grabación de Asignación de Lotes y Transferencia\\Cancelar\Asignar y Transferir")
          If OPX% <= 1 Then Exit Sub
        Else
          Exit Sub
        End If

        'RECORRE LOS LOTES TRAZABLES Y LOS GRABA A CONTINUACION EN CARDETPE.
        i& = 0
        While i& < ASINUMSE.MSF2.Rows - 1
          i& = i& + 1
         If ORD% = XVALO(ASINUMSE.MSF2.TextMatrix(i&, 7)) Then
             Call cmdEliminaItem(ASINUMSE.MSF2, i&)
          End If
        Wend
        
        J& = ASINUMSE.MSF2.Rows - 1
        For i& = 1 To LOTESPED.List2.ListCount
             'CODIGO DE LIST DE LOTESPED
              Ci% = CODINT%(TRIM$(Left$(LOTESPED.List2.List(i& - 1), 16)))
              CODI$ = TRIM$(Left$(LOTESPED.List2.List(i& - 1), 16))
              DESCRI$ = TRIM$(Mid$(LOTESPED.List2.List(i& - 1), 17, 32))
              CANTI = XVALO(Mid$(LOTESPED.List2.List(i& - 1), 50, 11))
              LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(i& - 1), 62, 16))
              ORDEX% = XVALO(Mid$(LOTESPED.List2.List(i& - 1), 132, 3))
              '
              J& = J& + 1
              ASINUMSE.MSF2.Rows = J& + 1
              ASINUMSE.MSF2.TextMatrix(J&, 1) = Ci%
              ASINUMSE.MSF2.TextMatrix(J&, 2) = CODI$
              ASINUMSE.MSF2.TextMatrix(J&, 3) = DESCRI$
              ASINUMSE.MSF2.TextMatrix(J&, 4) = CANTI
              ASINUMSE.MSF2.TextMatrix(J&, 5) = CANTI
              ASINUMSE.MSF2.TextMatrix(J&, 6) = LOTE$
              ASINUMSE.MSF2.TextMatrix(J&, 7) = ORD%
              OK% = 1
        Next i&
        '
     End If
   End If
   '
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans

'   If OK% > 0 Then
'      OPX% = COMALTER%("Grabación de Asignación de Lotes\\Cancelar\Asignar Lotes")
      If OPX% > 1 Then
         Call GrabarLotesPedido(NUPED&, ORD%, Ci%)
         Call TransfReservaPedidoConLote(NUPED&, ORD%, Ci%)
      Else
         LOTESPED.List2.Clear
         Label28 = ""
         Screen.MousePointer = 1
         Exit Sub
      End If
'   End If
   
   Screen.MousePointer = 1
   
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Screen.MousePointer = 1
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Screen.MousePointer = 1
      Call COMUNI("Asignación de Lotes/Transferencia Generada Correctamente")
      Call CARLISPED
    End If
End Sub
Sub TransfReservaPedidoConLote(NUPED&, ORD%, CI1%, Optional DESASIGNAR%, Optional SOLONOTRAZABLES%)
        'IF PARAMETRO OPCIONAL DESASIGNAR% > 0 REVIERTE EL MOVIMIENTO DE STOCK
        For J& = 1 To MSF2.Rows - 1
           If ORD% = XVALO(MSF2.TextMatrix(J&, 7)) Then
                Ci% = XVALO(MSF2.TextMatrix(J&, 1))
                CANTI = XVALO(MSF2.TextMatrix(J&, 4))
               
                If SOLONOTRAZABLES% > 0 Then
                   CANTI = XVALO(MSF2.TextMatrix(J&, 10))
                End If
                
                If DESASIGNAR% > 0 Then
                   CANTI = XVALO(MSF2.TextMatrix(J&, 5))
                End If
                If CANTI < 1 Then GoTo 70
                LOTE$ = MSF2.TextMatrix(J&, 6)
                ORDENITEM% = ORD%
                Ci% = CI1%
                DEPONU% = XVALO(Control$("", "DPREPPAÑ")) ' DEPOSITO 44 DE MEDITEA CONFIGURADO EN SETUP DE VENTAS  'Pañol/Factur'
                DepOrig% = XVALO(Control$("", "DEPOCEN")) ' DEPOSITO 50 PAÑOL DE MEDITEA CONFIGURADO EN SETUP DE FABRICACION/VALEMAT 'deP/Pañol'
                
                FECHH% = HOY(HO$)
                NUMCOM$ = ProNroComprobante("TR")
                NUMCOM$ = "TR." + TRIM$(NUMCOM$)
                REF$ = "Transf Deposito N°Ped-" & NUPED&
                '
                NPED& = NUPED&
                '
                If DESASIGNAR% > 0 Then
                   CANTI = CANTI * -1 'INVIERTO SIGNO DE CANTI PARA QUE REVIERTA.
                   REF$ = "Revierte Transf. N°Ped-" & NUPED&
                End If
                
                Call MOVISTOK(FECHH%, Ci%, LOTE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DepOrig%, CANTI * (-1) * 1, NPED&)
                Call MOVISTOK(FECHH%, Ci%, LOTE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DEPONU%, CANTI, NPED&)
                'ACTUALIZA LA CANTIDAD ASIGNADA EN EL PEDIDO
                CONDI$ = "NROPED = " & NUPED& & " AND NuOrItem = " & ORD%
                
                CANTI_A_ACTUALIZAR = XVALO(MSF2.TextMatrix(J&, 5)) + XVALO(MSF2.TextMatrix(J&, 10))
                If CANTI <= 0 Then CANTI_A_ACTUALIZAR = 0
                
                Call ACTUREGISTRO("PEDDETA", "CANASIG", CONDI$, CANTI_A_ACTUALIZAR, REGAF&, "NOMESS")

           End If
70      Next J&

End Sub

Private Sub Frame14_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label14_Change()
    Label13 = ""
    Ci% = CODINT%(Label14)
    If Ci% > 0 Then
       Label13 = DESCRIT0$(Ci%)
    End If
'    '
'    Call SETULTREG("!INUSEDIS", 0)
'    TTXX$ = REGBLAN$("INUSEDIS")
'    '
'    jj& = 0
'    If Ci% > 0 Then
'      '
'      If ESTRAZABLE(Ci%) = 2 Or ESTRAZABLE(Ci%) = -1 Then
'         SQLX$ = "SELECT NUMEROSERIE, CODIGO, DESCRIPCION FROM TANUMSE INNER JOIN MASTER WITH(NOLOCK) ON (TANUMSE.COD_INTE = MASTER.CODINT) "
'         SQLX$ = SQLX$ + " WHERE TANUMSE.Cod_Inte = " & Ci% & ""
'         SQLX$ = SQLX$ + " AND (TANUMSE.NumRemito = '' OR TANUMSE.NumRemito IS NULL)"
'         SQLX$ = SQLX$ + " AND DepoNumse >= 0" ' PARA QUE TRAIGA SOLO LOS TERMINADOS.
'         Set NUSERTMP = READSET(SQLX$)
'
'         With NUSERTMP
'           Do While Not .EOF
'              NUSERR$ = SAFETEXT$(!NUMEROSERIE)
'              NusePed$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NUSERR$ & "'")
'              If NusePed$ = "" Then
'                 Call REPLA(TTXX$, NUSERR$, 1, 16)
'                 COD$ = SAFETEXT$(!Codigo): DESC$ = SAFETEXT$(!Descripcion)
'                 Call REPLA(TTXX$, COD$ + "  " + DESC$, 17, 47)
'                 jj& = jj& + 1
'                 Call GRAREG("!INUSEDIS", TTXX$, jj&)
'              End If
'           .MoveNext
'           Loop
'         End With
'         NUSERTMP.Close
'         Set NUSERTMP = Nothing
'      End If
'    End If
'    Call SETULTREG("!INUSEDIS", jj&)
'    Call FRESHECHO("!INUSEDIS")
'    '
End Sub


Private Sub Label28_Change()
    Label8 = ""
    Label11 = ""
    Label10 = ""
End Sub

Private Sub Label4_Change()
    Label5.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Label4.Caption))))
End Sub

Private Sub Label8_Change()
    Label10 = ""
    Ci% = CODINT%(Label8)
    If Ci% > 0 Then
       Label10 = DESCRIT0$(Ci%)
    End If
    '
    Call SETULTREG("!INUSEDIS", 0)
    TTXX$ = REGBLAN$("INUSEDIS")
    '
    jj& = 0
    If Ci% > 0 Then
      '
      If ESTRAZABLE(Ci%) = 2 Or ESTRAZABLE(Ci%) = -1 Then

         SQLX$ = "SELECT NUMEROSERIE, CODIGO, DESCRIPCION FROM TANUMSE INNER JOIN MASTER WITH(NOLOCK) ON (TANUMSE.COD_INTE = MASTER.CODINT) "
         SQLX$ = SQLX$ + " WHERE TANUMSE.Cod_Inte = " & Ci% & ""
         SQLX$ = SQLX$ + " AND (TANUMSE.NumRemito = '' OR TANUMSE.NumRemito IS NULL)"
         SQLX$ = SQLX$ + " AND DepoNumse >= 0" ' PARA QUE TRAIGA SOLO LOS TERMINADOS.
         Set NUSERTMP = READSET(SQLX$)
         
         With NUSERTMP
           Do While Not .EOF
              NUSERR$ = SAFETEXT$(!NUMEROSERIE)
              NusePed$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NUSERR$ & "'")
              If NusePed$ = "" Then
                 Call REPLA(TTXX$, NUSERR$, 1, 16)
                 COD$ = SAFETEXT$(!Codigo): DESC$ = SAFETEXT$(!Descripcion)
                 Call REPLA(TTXX$, COD$ + "  " + DESC$, 17, 47)
                 jj& = jj& + 1
                 Call GRAREG("!INUSEDIS", TTXX$, jj&)
              End If
           .MoveNext
           Loop
         End With
         NUSERTMP.Close
         Set NUSERTMP = Nothing
      End If
    End If
    Call SETULTREG("!INUSEDIS", jj&)
    Call FRESHECHO("!INUSEDIS")
    '
End Sub

Private Sub List1_Click()
    If TRIM$(Mid$(List1.TEXT, 48, 16)) <> "" Then
      Call MENSERR(24, "No Válido Des-Asignar.\Utilice Función Correspondiente.")
      Exit Sub
    End If
    Label8 = TRIM$(Left$(List1.TEXT, 16))
    Label11 = TRIM$(Mid$(List1.TEXT, 64, 6))
    Label1 = ""
    Text1 = TRIM$(Str$(DEPOEXPE%))
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Command4_Click
    End If
End Sub

Private Sub mnuIngresoManualSerie_Click()
    'INGRESO DE EQUIPOS MANUAL PARA EYM POR LA TRANSICION PARA QUE PUEDAN INGRESAR L OS QUE NO SE HICIERON
    'CON LA  NUEVA MODALIDAD
    Call REPLA$(VDEF$, MKI$(HOY(HO$)), 35, 2) ' PRESENTA POR DEFECTO LA FECHA ACTUAL
    '
10  OBX$ = OBJPLA$("ALTASERIE", VDEF$) ' OBJETO TANUMSE PARA INGRESAR LOS NUMEROS DE SERIE
    VDEF$ = OBX$
    If OBX$ = "" Then Exit Sub
    NUSERIEE$ = TRIM$(Mid$(VDEF$, 3, 32))
    Ci% = CVI(Mid$(VDEF$, 1, 2))
    fech% = CVI(Mid$(VDEF$, 35, 2))
    FECC = FECHATEX(fech%)
    
    CANTIRE& = CantRegistros&(" TANUMSE", "NumeroSerie = '" & NUSERIEE$ & "'") 'VERIFICA SI EL NUMERO DE SERIE EXISTE
    '
    If CANTIRE& > 0 Then Call MENSERR(24, "Número de Serie Existente"): GoTo 10 ' VALIDACIONES
    If Ci% < 1 Then Call MENSERR(24, "Código no Válido"): GoTo 10
    
    DETER% = XVALO(Control("", "DEPOCIE"))
    
    SQLX$ = "SELECT * FROM TANUMSE WHERE COD_INTE=0" ' SI PASO LAS VALIDACIONES AGREGA EL NUEVO NUMERO DE SERIE
    Dim TANUTMP As ADODB.Recordset
    Set TANUTMP = New ADODB.Recordset
    TANUTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      With TANUTMP
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0
        !NUMEROSERIE = NUSERIEE$
        !MESANIO = Right(FECC, 2)
        !NUMEORD = 1
        !cod_inte = Ci%
        !Cod_Exte = CODEXT(Ci%)
        Descripcion$ = DESCRIT(Ci%)
        !Descrip = Descripcion$
        !FechAlta = FECC
        !TEN_ENSA = 0
        !DepoNumSe = DETER%
        .Update
        '
      End With
    TANUTMP.Close
    Set TANUTMP = Nothing
    '
    Call CARGALISEL("INDINUSE", "TANUMSE", "NUMEROSERIE/A16;COD_EXTE/A18;DESCRIP/A24", "NUMEROSERIE <> '' ORDER BY FECHALTA ASC, NUMEORD ASC")
    '
    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
    '
    CUERPO$ = "Alta N/Serie en Forma Manual"
    Call ANOTANUSE(NUSERIEE$, CUERPO$)

End Sub

Private Sub MSF_Click()
    REG& = MSF.MouseRow
    COLUM% = MSF.MouseCol
    If REG& = 0 Or REG& > MSF.Rows - 1 Then Exit Sub
    
    If TRIM$(MSF.TextMatrix(REG&, 4)) <> "" Then
      Call MENSERR(24, "No Válido Des-Asignar.\Utilice Función Correspondiente.")
      Exit Sub
    End If
    Label8 = TRIM$(MSF.TextMatrix(REG&, 2))
    Label11 = TRIM$(MSF.TextMatrix(REG&, 6))
    Label1 = ""
    Text1 = TRIM$(Str$(DEPOEXPE%))

End Sub


Private Sub MSF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Command4_Click
    End If
End Sub

Private Sub msf2_Click()
    REG& = MSF2.MouseRow
    If REG& = 0 Or REG& > MSF2.Rows - 1 Then Exit Sub
    Label14 = TRIM$(MSF2.TextMatrix(REG&, 2))
    Label12 = TRIM$(MSF2.TextMatrix(REG&, 7))
    Label1 = ""


End Sub

Private Sub MSF2_DblClick()
   REG& = MSF2.MouseRow
   If REG& = 0 Or REG& > MSF2.Rows - 1 Then Exit Sub

    
   If TRIM$(MSF2.TextMatrix(REG&, 6)) <> "" Then
'     Call MENSERR(24, "No Válido Des-Asignar.\Utilice Función Correspondiente.")
      Exit Sub
   End If

   REG& = MSF2.MouseRow
   ORD% = MSF2.TextMatrix(REG&, 7)
   CI1% = XVALO(MSF2.TextMatrix(REG&, 1))
   If ESTRAZABLE%(CI1%) > 0 Then
      Call CARGALOTE_PEDIDO(XVALO(Label28), REG&, ORD%)
   Else
      CantidadOriginal$ = TRIM$(MSF2.TextMatrix(REG&, 4))
      CantidadInformadaActual$ = TRIM$(MSF2.TextMatrix(REG&, 5))
      CantPendDisponible$ = TRIM$(MSF2.TextMatrix(REG&, 9))
      CANTIDADINFORMADA$ = TRIM$(InputBox$("Cantidad Pendiente Disponible: " & CantPendDisponible$, "Ingrese Cantidad A Transferir", MSF2.TextMatrix(REG&, 10)))
      'SI CANCELA
      If CANTIDADINFORMADA$ = "" Then Exit Sub
      '
      If XVALO(CantidadInformadaActual$) + XVALO(CANTIDADINFORMADA$) > CantidadOriginal$ Then
         Call COMUNI("Imposible Transferir!!!\La Cantidad a Transferir Agregada más la yá Transferida Supera la cantidad Total del Pedido")
         Exit Sub
      End If
      If XVALO(CANTIDADINFORMADA$) > XVALO(CantPendDisponible$) Then
         Call COMUNI("Imposible Transferir!!!\La Cantidad a Transferir Agregada Supera la Disponibilidad de Stock")
         Exit Sub
      End If
      
      MSF2.TextMatrix(REG&, 10) = FORMATNUM$(XVALO(CANTIDADINFORMADA$), "f8.1")
   End If
   
End Sub


Private Sub MSF2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If MSF2.Rows <= 1 Then Exit Sub
  With MSF2
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

Private Sub Text1_Change()
    Label6.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text1))))
End Sub

Sub CARLISPED() 'Carga la lista con el pedido seleccionado
    '
    NPSEL& = XVALO(ASINUMSE.Label28)
    Label1 = "": Label11 = ""
    Label4 = ""
    Text1 = ""
    MSF.Rows = 1
    MSF2.Rows = 1
    'ASINUMSE.List1.Clear
    '
    'Call CARGALISTA(ASINUMSE.List1, "PEDDETA", "CodiInt/C16+MASTER/A29; NUMEROSERIE/A16; NuOrItem/F5", "NroPed = " & NPSEL& & " AND Status < 8  AND ((CanPed - CantEnt) >= 0.05) ORDER BY NuOrItem ASC")
    'ITEMS TRAZABLES = 2
    REG& = 0
    SQLX$ = "SELECT P.CODIINT, P.CodiExt,P.Descrip,P.NUMEROSERIE,P.IdenLote,P.NuOrItem FROM PEDDETA P WITH(NOLOCK) "
    SQLX$ = SQLX$ + " INNER JOIN MASTER M ON P.CODIINT = M.CODINT WHERE (M.TRAZABLE = 2 OR M.TRAZABLE = -1) "
    SQLX$ = SQLX$ + " AND P.NroPed = " & NPSEL& & " AND P.Status < 8  AND ((P.CanPed - P.CantEnt) >= 0.05) ORDER BY P.NuOrItem ASC"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
        MSF.TextMatrix(REG&, 1) = XVALO(!CODIINT)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!CODIEXT)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 4) = SAFETEXT$(!NUMEROSERIE)
        MSF.TextMatrix(REG&, 5) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 6) = XVALO(!NuOrItem)
        
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    '
    'ITEMS TRAZABLES = 1
    REG& = 0
    DEPOPAÑOL1% = XVALO(Control$("", "DEPOCEN"))
    INCLUYEITEMNOTRAZABLES% = XVALO(Control("ASINUMSE", "INCNOTRA"))
    SQLX$ = "SELECT P.CODIINT, P.CodiExt,P.Descrip,P.NUMEROSERIE,P.IdenLote,P.NuOrItem,P.CANPED, P.CANASIG FROM PEDDETA P WITH(NOLOCK) "
    SQLX$ = SQLX$ + " INNER JOIN MASTER M ON P.CODIINT = M.CODINT "
    '
    If INCLUYEITEMNOTRAZABLES% > 0 Then 'CONFIGURACION MEDITEA DE INCLUIR TODOS LOS ITEMS TRAZABLE 0 Y 1
       SQLX$ = SQLX$ + " WHERE (M.TRAZABLE = 1  OR M.TRAZABLE = 0 OR M.TRAZABLE IS NULL)"
       Dim STOCK(32767):
    Else
       SQLX$ = SQLX$ + " WHERE M.TRAZABLE = 1  "
    End If
    SQLX$ = SQLX$ + " AND P.NroPed = " & NPSEL& & " AND P.Status < 8  AND ((P.CanPed - P.CantEnt) >= 0.05) ORDER BY P.NuOrItem ASC, M.TRAZABLE  DESC"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        REG& = REG& + 1
        MSF2.Rows = REG& + 1
        MSF2.TextMatrix(REG&, 1) = XVALO(!CODIINT)
        MSF2.TextMatrix(REG&, 2) = SAFETEXT$(!CODIEXT)
        MSF2.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
        MSF2.TextMatrix(REG&, 4) = XVALO(!CanPed)
        MSF2.TextMatrix(REG&, 5) = XVALO(!CanAsig)
        MSF2.TextMatrix(REG&, 6) = SAFETEXT$(!idenlote)
        MSF2.TextMatrix(REG&, 7) = XVALO(!NuOrItem)
        If INCLUYEITEMNOTRAZABLES% > 0 Then
          CI1% = XVALO(MSF2.TextMatrix(REG&, 1))
          If ESTRAZABLE(CI1%) < 1 Then
            ST = STODEPOS(XVALO(!CODIINT), DEPOPAÑOL1)
            If ST < 0 Then ST = 0
            MSF2.TextMatrix(REG&, 8) = ST
          End If
          
        End If
        
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    If INCLUYEITEMNOTRAZABLES% > 0 Then
        For REG& = 1 To MSF2.Rows - 1
              CI1% = XVALO(MSF2.TextMatrix(REG&, 1))
              If ESTRAZABLE(CI1%) < 1 Then
                CANPE = XVALO(MSF2.TextMatrix(REG&, 4))
                CANASI = XVALO(MSF2.TextMatrix(REG&, 5))
                NECESIDAD = CANPE - CANASI
                If STOCK(CI1%) = 0 Then
                   STOCK(CI1%) = MSF2.TextMatrix(REG&, 8) 'LE ASIGNO EL STOCK AL VECTOR
                End If
                If NECESIDAD < 0 Then NECESIDAD = 0
                If NECESIDAD > 0 Then
                   If NECESIDAD > XVALO(STOCK(CI1%)) Then
                      NECESIDAD = STOCK(CI1%)
                   End If
                End If
                MSF2.TextMatrix(REG&, 9) = FORMATNUM$(NECESIDAD, "F8.1")
                
                MSF2.TextMatrix(REG&, 10) = FORMATNUM$(NECESIDAD, "F8.1") 'por defecto presenta la cantidad pendinete disponible
                STOCK(CI1%) = XVALO(STOCK(CI1%)) - XVALO(MSF2.TextMatrix(REG&, 9))
                Call COLOREAR_GRILLA_SOLAFILA(MSF2, &HCDDADA, REG&)
              End If
        Next REG&
    End If
    
End Sub



Sub PRINTASINUMSE()
    '
    FORIPRE$ = TRIM$(Control$("", "FOASNUSE"))
    If FORIPRE$ = "" Then Exit Sub
    '
    NUMEDOC$ = Label1.Caption
    TIPODOC$ = "Asignacion Numero de Serie"
    '
    'CODIGO DE BARRAS DE NUMEROS DE SERIE
    NUMSE$ = TRIM$(Label1)
    'Convierte al formato del codigo de barras
    NUMSE1$ = CONVIERTOCODIGOBARRAS$(NUMSE$)
    '
    FECHH = HOY(HO$)
    CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = Label1.Caption     'Numero de serie
    CODPARAM$(2) = "COD-MAT": DOCPARAM$(2) = Label8.Caption      'codigo del articulo
    CODPARAM$(3) = "REF-MAT2": DOCPARAM$(3) = Text1.TEXT      'numero de Deposito de destino
    CODPARAM$(4) = "REF-MAT3": DOCPARAM$(4) = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text1.TEXT))))   'Nombre Deposito de destino
    CODPARAM$(5) = "FECH-EMI": DOCPARAM$(5) = FECHATEX(FECHH)
    CODPARAM$(6) = "N-PLANO": DOCPARAM$(6) = NUMSE1$    'Numero de serie (CODIGO DE BARRAS)
    CAPARDOC = 6
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub
Sub BLANFORMU()
    Label1 = ""
    Label11 = ""
    Label4 = ""
    Label8 = ""
    
    Text1 = ""
    Label28 = ""
'    List1.Clear
    MSF.Rows = 1

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Shift = 1 Then
     If Button = 2 Then
       TTX1$ = "En esta Grilla Se Cargan Tanto Items  Trazables =0 y 1" + vbCrLf
       TTX1$ = TTX1$ + "Los Items Pintados en Marrón Son Trazables = 0" + vbCrLf
       TTX1$ = TTX1$ + "Estos Utilizando el botón 'Transferir Items No Trazables'" + vbCrLf
       TTX1$ = TTX1$ + "Generan la Transferencia x la Cantidad de la columna 'Transf'" + vbCrLf
       TTX1$ = TTX1$ + "Del Dep. Pañol al Dep. Configurado en Setup Ventas: 'Pañol/Factur'" + vbCrLf
       TTX1$ = TTX1$ + "Es posible Realizar Transferencias Parciales (Solo en el caso de los No Trazables)" + vbCrLf
       TTX1$ = TTX1$ + "Para Esto Realizar Doble Clic Sobre la Fila a Modificar la Cantidad a Transferir" + vbCrLf
       TTX1$ = TTX1$ + "Tener En Cuenta que la Cantidad a Transferir por Defecto considera la Disponibilidad" + vbCrLf + vbCrLf
       TTX1$ = TTX1$ + "=============================================" + vbCrLf + vbCrLf
       TTX1$ = TTX1$ + "En el caso de los Trazables = 1 " + vbCrLf
       TTX1$ = TTX1$ + "Se Transfieren x Lote Realizando Doble Clic - Se Asigna El Lote y se Transfiere Inmediateamente" + vbCrLf + vbCrLf
       TTX1$ = TTX1$ + "=============IMPORTANTE=================" + vbCrLf
       TTX1$ = TTX1$ + "Para que en la Grilla Inferior También Muestre los Items no Trazables " + vbCrLf
       TTX1$ = TTX1$ + "Se Deberá Configurar desde el Setup de Esta Misma Pantalla" + vbCrLf
       TTX1$ = TTX1$ + "Marcar Check ('Incluir en Asignación de Items a Pedidos art. No Trazab....')"
       
       
       
       MsgBox TTX1$
       Exit Sub
     End If
  End If
  
  
  If Button = 2 Then
     Set MSFGLOBALFRM = MSF2
     GLOBALFRM_INDIX% = Index
     REG& = MSFGLOBALFRM.MouseRow
     COL% = MSFGLOBALFRM.MouseCol
     Call MENU_CELDA(MSFGLOBALFRM, REG&, COL%)
  End If

End Sub


Sub MENU_CELDA(MSF As msFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&

        PopupMenu MNUARCHIVO
        Call INVICTRL
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
  mnucopiar.Visible = Not (mnucopiar.Visible)
  mnuDesasignar.Visible = Not (mnuDesasignar.Visible)
  '
End Sub

Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuDesasignar_Click()
   REG& = GLOBALFRM_FILA&
   NUPED& = XVALO(Label28)
   If NUPED& < 1 Then
      Call COMUNI("No se Detectó un Pedido Activo")
      Exit Sub
   End If
   CI1% = XVALO(ASINUMSE.MSF2.TextMatrix(REG&, 1))
   ESTRAX% = ESTRAZABLE(CI1%)
   LOTE$ = TRIM$(ASINUMSE.MSF2.TextMatrix(REG&, 6))
   
   If ESTRAX% <= 0 Then
     CANTIASIG = XVALO(ASINUMSE.MSF2.TextMatrix(REG&, 5))
     If CANTIASIG <= 0 Then
        Call COMUNI("Registro Seleccionado No Tiene Cantidad Asignada")
        Exit Sub
     End If
     
   End If
   
   If ESTRAX% > 0 Then
     If LOTE$ = "" Then
      Call COMUNI("Registro Seleccionado No Posee Lote Asignado")
      Exit Sub
     End If
   End If
   
   ORD% = XVALO(ASINUMSE.MSF2.TextMatrix(REG&, 7))
   If ESTRAX% > 0 Then
     'SOLICITA CONFIRMACION
      OPX% = COMALTER("Atención !!! Al Desasignar Lote Para El Código Seleccionado\Tenga en Cuenta Que Se Des-Asginarán El Conjunto de Lotes\Para el Código Seleccionado\\Cancelar\Des-Asignar")
   Else
      
      OPX% = COMALTER("Atención !!! Confirma Revertir Material Transferido\\Cancelar\Revertir Transferencia")
   End If
   If OPX% < 2 Then Exit Sub
   '
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans

   Call DesasignaLotesPedido(NUPED&, ORD%, CI1%)
   Call TransfReservaPedidoConLote(NUPED&, ORD%, CI1%, 1) 'parametro 1 para revertir transferencia
   
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Screen.MousePointer = 1
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Screen.MousePointer = 1
      Call COMUNI("Des-Asignación Generada Correctamente")
      Call CARLISPED
    End If
    
    Screen.MousePointer = 1

End Sub

