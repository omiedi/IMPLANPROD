VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form DISPSTK07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disponibilidad General de Stocks"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1080
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   34
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "FORMATO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   1800
      TabIndex        =   31
      Top             =   3360
      Width           =   1905
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Disponibilidad"
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
         Left            =   210
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Cobertura"
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
         Left            =   210
         TabIndex        =   32
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   7080
      ScaleHeight     =   7545
      ScaleWidth      =   1665
      TabIndex        =   26
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   36
         Top             =   4080
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   105
         Picture         =   "DISPSTK07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   2640
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DISPSTK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1005
         Width           =   650
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   105
         Picture         =   "DISPSTK07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "DISPSTK07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1785
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   3585
      Begin VB.CheckBox CheckNoConsiderarCompras 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "No Considerar Compras Pendientes"
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
         Height          =   300
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   3180
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Todos los Tipos"
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
         Left            =   210
         TabIndex        =   18
         Top             =   285
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3150
         TabIndex        =   15
         Top             =   285
         Width           =   175
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   2730
         TabIndex        =   14
         Top             =   285
         Width           =   435
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Todos - Sin Clasificación"
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
         Left            =   210
         TabIndex        =   17
         Top             =   705
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
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
         Left            =   2010
         TabIndex        =   16
         Top             =   360
         Width           =   630
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   9135
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROLES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3810
      TabIndex        =   9
      Top             =   3360
      Width           =   3165
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Verificar Movimientos y Saldos de Stock"
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
         Height          =   660
         Left            =   180
         TabIndex        =   25
         Top             =   240
         Width           =   2580
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "RANGO DE CÓDIGOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6840
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
         Left            =   6480
         TabIndex        =   3
         Top             =   460
         Width           =   175
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
         Left            =   3180
         TabIndex        =   5
         Top             =   460
         Width           =   175
      End
      Begin VB.TextBox Text1 
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
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   2320
      End
      Begin VB.TextBox Text2 
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
         Left            =   4200
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   495
         Left            =   3465
         TabIndex        =   12
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   495
         Left            =   105
         TabIndex        =   11
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   570
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1590
      Begin VB.OptionButton Option12 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Impresión"
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
         TabIndex        =   22
         Top             =   630
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option11 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Pantalla"
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
         TabIndex        =   21
         Top             =   380
         Width           =   1170
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "FILTRO"
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
      Left            =   3810
      TabIndex        =   0
      Top             =   1800
      Width           =   3165
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Productos Terminados"
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
         Left            =   420
         TabIndex        =   24
         Top             =   1125
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Todos los Artículos"
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
         Left            =   420
         TabIndex        =   23
         Top             =   855
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Artículos con Movimiento"
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
         Left            =   420
         TabIndex        =   20
         Top             =   585
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Solo Artículos Faltantes"
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
         Left            =   420
         TabIndex        =   19
         Top             =   315
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "DISPSTK07.frx":0BA0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "DISPSTK07.frx":0DD4
      TabIndex        =   35
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "DISPSTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IAT&
Dim DETALLEOCOM$(32767) 'VECTOR PARA EL DETALLE DE ORDENES DE COMPRA EN TABLA DISPONIB

'

Sub CARGATABLADISP()
    '
    'BORRO LA TABLA DISPONIB
    CONDI$ = "COD_INTE > 0"
    Call BORRAREGISTROS("DISPONIB", CONDI$, REGAF&, "NOMESS")
    '
    FIN& = ULTREG&("STODISP")
    For J& = 1 To FIN&
    
       TX$ = REGLEIDO$("STODISP", J&)
       Call TRACE(20, J&, FIN&)
       CI% = CVI(Mid$(TX$, 1, 2)) 'CODIGO
       TIMA% = Asc(Mid$(TX$, 3, 1)) ' TIPO DE MATERIAL
       STACTU1 = CVS(Mid(TX$, 5, 4)) 'STOCK ACTUAL
       QC1 = CVS(Mid$(TX$, 9, 4)) 'CANTIDAD COMPRAS
       QOF1 = CVS(Mid$(TX$, 13, 4)) 'CANTIDAD O/F
       STMIN1 = CVS(Mid$(TX$, 17, 4)) 'STOCK MINIMO
       PPENT1 = CVS(Mid$(TX$, 21, 4)) 'PEDIDOS PENDIENTES
       RESE = CVS(Mid$(TX$, 25, 4))   'RESERVA
       DISPNETO = CVS(Mid$(TX$, 29, 4)) 'DISPONIBILIDAD NETA
       FALTA = CVS(Mid$(TX$, 33, 4))    'FALTANTE
       LOTECO = CVS(Mid$(TX$, 37, 4))    'LOTE ECONOMICO
       '
       SQLX$ = "SELECT COD_INTE,TIPOMAT,STOCK,COMPRAS,FABRICACION,MINIMO,PEDIDOS,DISPONIBLE,LOTEECON,CodiEmpr,reservas, "
       SQLX$ = SQLX$ + " FALTANTE,DATOCOM FROM DISPONIB"
       SQLX$ = SQLX$ + " WHERE COD_INTE < 1 "
       Dim rs As ADODB.Recordset
       Set rs = New ADODB.Recordset
25     On Error Resume Next
       rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With rs
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
         !cod_inte = CI%
         !TIPOMAT = TIMA%
         !STOCK = STACTU1
         !compras = QC1
         !FABRICACION = QOF1
         !Minimo = STMIN1
         !Pedidos = PPENT1
         !RESERVAS = RESE
         !DISPONIBLE = DISPNETO
         
         !FALTANTE = FALTA
         !LOTEECON = LOTECO
         !DATOCOM = DETALLEOCOM$(CI%)
         .Update
      End With
      
         
'       'ACTUALIZA TABLA DISPONIB
'       Call CreaRegistro("DISPONIB", "COD_INTE", CI%, "NOMESS")
'       CONDI$ = "COD_INTE = " & CI%
''       Call ACTUREGISTRO("DISPONIB", "COD_INTE", CONDI$, CI%, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "TIPOMAT", CONDI$, TIMA%, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "STOCK", CONDI$, STACTU1, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "COMPRAS", CONDI$, QC1, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "FABRICACION", CONDI$, QOF1, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "MINIMO", CONDI$, STMIN1, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "PEDIDOS", CONDI$, PPENT1, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "RESERVAS", CONDI$, RESE, REGAF&, "NOMESS")
''       Call ACTUREGISTRO("DISPONIB", "DISPONIBLE", CONDI$, DISPNETO, REGAF&, "NOMESS")
'       Call ACTUREGISTRO("DISPONIB", "FALTANTE", CONDI$, FALTA, REGAF&, "NOMESS")
'       Call ACTUREGISTRO("DISPONIB", "LOTEECON", CONDI$, LOTECO, REGAF&, "NOMESS")
'            'DETALLEOCOM$(CI%) VECTOR A NIVEL DE FRM QUE SE CARGA EN RUTINA STODISP
'       Call ACTUREGISTRO("DISPONIB", "DATOCOM", CONDI$, DETALLEOCOM$(CI%), REGAF&, "NOMESS")
    Next J&
    
   
End Sub

Sub CREARTABLADISPNIB()
    '
    Call ABRESTOCKS
    '
    Call CREACAMPO("STOCKS", "DISPONIB", "COD_INTE", 3, 0)
    Call CREACAMPO("STOCKS", "DISPONIB", "TIPOMAT", 3, 0)
    Call CREACAMPO("STOCKS", "DISPONIB", "STOCK", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "COMPRAS", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "FABRICACION", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "MINIMO", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "PEDIDOS", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "RESERVAS", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "DISPONIBLE", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "FALTANTE", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "LOTEECON", 7, 8)
    Call CREACAMPO("STOCKS", "DISPONIB", "DATOCOM", 12, 0)
  
End Sub

Private Sub Check1_Click()
   Text3.SetFocus
End Sub

Private Sub Check1_LostFocus()
   If Check1.Value = 1 Then
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Command1_Click()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Unload Me
End Sub

Sub Command12_Click()
   Command12.Enabled = False
   Call STODISP
   Command12.Enabled = True
End Sub
Private Sub Command2_Click()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   If (USNBR% Mod 100) = 1 Then
      On Error Resume Next
      OPDISTK07.Show 1
      On Error GoTo 0
   End If
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Sub ANACOBER()
  '\\\OT-06-0102-WAR-22/03/06///
   Screen.MousePointer = 11
   'ESCRIBO EL ARCHIVO COBEVEST
   JJ& = 0
   Y$ = REGBLAN$("COBEVEST")
   FIN& = ULTREG("STODISP")
   For J& = 1 To FIN&
     Call TRACE(20, J&, FIN&)
     X$ = REGLEIDO$("STODISP", J&)
     CIXJ% = CVI(Mid$(X$, 1, 2)) 'CODIGO
     STACTU1 = CVS(Mid(X$, 5, 4)) 'STOCK ACTUAL
     QC1 = CVS(Mid$(X$, 9, 4)) 'CANTIDAD COMPRAS
     QOF1 = CVS(Mid$(X$, 13, 4)) 'CANTIDAD O/F
     QAI1 = QC1 + QOF1 'CANTIDAD A INGRESAR
     STMIN1 = CVS(Mid$(X$, 17, 4)) 'STOCK MINIMO
     PPENT1 = CVS(Mid$(X$, 21, 4))
     TOTA1 = STACTU1 - PPENT1 + QAI1 'ST. ACT. + CANTIDAD A INGRESAR
     COB1 = 0
     COB2 = 0
     '
     If STMIN1 > 0 Then
        COB1 = (STACTU1 - PPENT1) / STMIN1
        COB2 = TOTA1 / STMIN1
     End If
     '
     Call REPLA(Y$, MKI$(CIXJ%), 1, 2)
     Call REPLA(Y$, MKS$(STMIN1), 3, 4)
     Call REPLA(Y$, MKS$(STACTU1), 7, 4)
     Call REPLA(Y$, MKS$(PPENT1), 11, 4)
     Call REPLA(Y$, MKS$(STACTU1 - PPENT1), 15, 4)
     Call REPLA(Y$, MKS$(COB1), 19, 4)
     Call REPLA(Y$, MKS$(QAI1), 23, 4)
     Call REPLA(Y$, MKS$(TOTA1), 27, 4)
     Call REPLA(Y$, MKS$(COB2), 31, 4)
     JJ& = JJ& + 1
     Call GRAREG("COBEVEST", Y$, JJ&)
     '
   Next J&
   '
   Call SETULTREG("COBEVEST", JJ&)
   Call FILESORT("COBEVEST", "COBEVEST", 7, 7, "ASC")
   Call CIERRARCH("COBEVEST")
   '
   Call HEADERS("COBEVEST", "")
   Call FINAL("*COBEVEST")
   '
99 Screen.MousePointer = 1
  ''\\\OT-06-0102-WAR-FIN///
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command6_Click()
End Sub


Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    INI& = 1
    '\\\OT-08-0990-OD-30/10/08///

     LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
     If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15

    '\\\OT-08-0990-OD-FIN///
    

5   XX$ = Left$(REGLEIDO$("ECOMAST", INI&), LONCODI%)
    Text1.TEXT = XX$
    CI1% = Abs(CODINT%(XX$))
    If CI1% < 1 Then
       If INI& < ULTREG&("ECOMAST") Then
          INI& = INI& + 1: GoTo 5
       End If
    End If
    Label1.Caption = DESCRIT$(CI1%)
    '
    AAA& = ULTREG&("ECOMAST")
10  XX$ = Left$(REGLEIDO$("ECOMAST", AAA&), LONCODI%)
    If TRIM$(XX$) = "" Then
      If AAA& > 1 Then
        AAA& = AAA& - 1
        GoTo 10
      End If
    End If
    Text2.TEXT = XX$
    CI2% = Abs(CODINT%(XX$))
    Label2.Caption = DESCRIT$(CI2%)
    '
    'CREA TABLA DE DISPONIBILIDAD
    Call CREARTABLADISPNIB
    'Call VERNOVTRA
    '
    '\\\OT-08-0990-OD-30/10/08///
    Text1.MaxLength = LONCODI%
    Text2.MaxLength = LONCODI%
    '\\\OT-08-0990-OD-fin///

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub


Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    COD1$ = TRIM$(VALACT1$("MASTER"))
    If COD1$ <> "" Then
       Text1.TEXT = COD1$
    End If
End Sub

Private Sub Text2_Change()
    CI2% = CODINT%(Text2.TEXT)
    If CI2% > 0 Then
         Label2.Caption = DESCRIT$(CI2%)
       Else
         Label2.Caption = ""
    End If
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("MASTER")
    COD2$ = TRIM$(VALACT1$("MASTER"))
    If COD2$ <> "" Then
       Text2.TEXT = COD2$
    End If
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("TATIMAT")
    VDEV$ = VALACT1$("TATIMAT")
    If VDEV$ <> "" Then
       Text3.TEXT = VDEV$
       If TRIM$(Text3.TEXT) = "" Then
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
          Else
            Label4.Caption = VALACT2$("TATIMAT")
            Check1.Value = 0
       End If
    End If
End Sub

Private Sub Text3_GotFocus()
   If TRIM$(Text3.TEXT) = "" Then
      Check1.Value = 1
      Label4.Caption = "Todos - Sin Clasificación"
   End If
End Sub

Private Sub Text3_LostFocus()
    If Text3.TEXT <> "" Then
      If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
        End If
      End If
      Text3.TEXT = ""
      Label4.Caption = "Todos - Sin Clasificación"
      Check1.Value = 1
    End If
End Sub


Private Sub Text1_Change()
    CI1% = CODINT%(Text1.TEXT)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
        VENDX$ = ECOARCH$("TATIMAT", Chr$(VENDE%))
        If VENDX <> "" Then
            Label4.Caption = VENDX$
            Check1.Value = 0
            Exit Sub
          Else
            Label4.Caption = "Todos - Sin Clasificación"
            Check1.Value = 1
        End If
    End If
End Sub

Private Sub Text8_Change()
End Sub

Sub STODISP()
    '
    HOII% = HOY(HOS$)
    '
    Call CARDEPOS
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    TMT% = XVALO(Text3.TEXT)
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
    End If
    '
    CI1% = CODINT%(Text1.TEXT)
    COD1$ = CODEXT$(CI1%)
    If TRIM$(COD1$) = "" Then Exit Sub
    CI2% = CODINT%(Text2.TEXT)
    COD2$ = CODEXT$(CI2%)
    If TRIM$(COD2$) = "" Then Exit Sub
    '
    NIVEFIL% = 1
    If Option2.Value = True Then NIVEFIL% = 2
    If Option1.Value = True Then NIVEFIL% = 3
    If Option4.Value = True Then NIVEFIL% = 4
    '
    Screen.MousePointer = 11
    Dim EXIS(16384, 6)
    Dim PUNOF$(16384)
    Dim PROTER%(16384)      ' MARCA PRODUCTO TERMINADO
    '
    UREMAS& = ULTREG&("MASTER")
    User% = USNBR% Mod 100
    If User% < 1 Then User% = 1
    '
    FIN& = NUMAS%
    For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       CIIX% = i&
       If Check2.Value = 1 Then
          Call VESARRAS(CIIX%)
       End If
       EXIS(i&, 1) = STOCKACT(CIIX%)
       AA = EXIS(i&, 1)
'       Debug.Print AA
       EXIS(i&, 4) = STOMIN(CIIX%)
    Next i&
    '
    'LIMPIAR EL VECTOR
    For R% = 1 To NUMAS%
      DETALLEOCOM$(R%) = ""
    Next R%
    '
    If CheckNoConsiderarCompras.Value > 0 Then GoTo 40
    
    Call ABRECOMPRAS
    SQLX$ = " SELECT N_ORSERV, Coef_Unids, cant_ocom, Cant_Rec, COD_INTE, Fentre_Sol, Nume_Ocom  FROM OCOMDETA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
    SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
    SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
    '
    If InStr(EPAC$, "GABAL") > 0 Then
        Dim OcoTemp As ADODB.Recordset
        Set OcoTemp = New ADODB.Recordset
        OcoTemp.CursorType = adOpenDynamic
        OcoTemp.LockType = adLockReadOnly
        OcoTemp.Open SQLX$, FLEXCONN, , , adCmdText
       Else
         Set OcoTemp = READSET(SQLX$)
    End If
    With OcoTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NUORSE& = !N_ORSERV
          If NUORSE& < 1 Then
            COEFCAN = !Coef_Unids
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = XVALO(!cant_ocom) * COEFCAN
            CAREC = XVALO(!Cant_Rec) * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              CIIX% = XVALO(!cod_inte)
              EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
              '
              'DATOS PARA INFORMAR EN VECTOR DETALLEOCOM$ LAS O.C. COMPROMETIDAS
              FE% = CVINT(!Fentre_Sol)
              NUMCOM& = XVALO(!nume_ocom)
              TX$ = ""
              Call REPLA(TX$, FECHATEX$(FE%), 1, 10)
              Call REPLA(TX$, " " + FORMATNUM$(CAPEN, "F14." & CNTDC$(CIIX%)), 11, 15)
              Call REPLA(TX$, " " + Str$(NUMCOM&), 26, 9)
              '
              DETALLEOCOM$(CIIX%) = DETALLEOCOM$(CIIX%) + TX$
              '
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    Call ABRECOMPRAS
    SQLX$ = " SELECT   Cant_Aten, Acuen_Aten, COD_INTE FROM AUTOENT WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Stat_Aten < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
    SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
    SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
    Dim AtenTemp As ADODB.Recordset
    Set AtenTemp = New ADODB.Recordset
    AtenTemp.CursorType = adOpenDynamic
    AtenTemp.LockType = adLockReadOnly
    If InStr(EPAC$, "GABAL") > 0 Then
        AtenTemp.Open SQLX$, FLEXCONN, , , adCmdText
      Else
        Set AtenTemp = READSET(SQLX$)
    End If
    With AtenTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NUORSE& = 0 ' !N_OrServ
          If NUORSE& < 1 Then
            COEFCAN = 1 ' !Coef_Unids
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = XVALO(!Cant_Aten) * COEFCAN
            CAREC = XVALO(!Acuen_Aten) * COEFCAN
            CAPEN = CACOM - CAREC
            If CAPEN > 0 Then
              CIIX% = XVALO(!cod_inte)
              EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
            End If
          End If
        .MoveNext
        Loop
      End If
    End With
    '
40
    Call ABRESTOCKS
    SQLX$ = " SELECT CANTNOM, CanToApro, COD_INTE FROM BOLRECEP WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    Dim COBRECEPTMP As ADODB.Recordset
    Set COBRECEPTMP = New ADODB.Recordset
    COBRECEPTMP.CursorType = adOpenDynamic
    COBRECEPTMP.LockType = adLockReadOnly
    If InStr(EPAC$, "GABAL") > 0 Then
         COBRECEPTMP.Open SQLX$, FLEXCONN, , , adCmdText
      Else
         Set COBRECEPTMP = READSET(SQLX$)
    End If
    With COBRECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While (.EOF = False)
          CACOM = XVALO(!CantNom)
          CAREC = XVALO(!CanToApro)
          CAPEN = CACOM - CAREC
          If CAPEN > 0 Then
            CIIX% = XVALO(!cod_inte)
            EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    '
    
    Call ABREORFAB
    SQLX$ = " SELECT CanProy, CanApro, COD_INTE FROM ORDEFABR WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE StatuOrf < 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
'    Dim OfaTemp As ADODB.Recordset
'    Set OfaTemp = New ADODB.Recordset
'    OfaTemp.CursorType = adOpenDynamic
'    OfaTemp.LockType = adLockReadOnly
'    OfaTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    Set OfaTemp = READSET(SQLX$)
    With OfaTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          CACOM = XVALO(!CANPROY)
          CAREC = XVALO(!CanApro)
          CAPEN = CACOM - CAREC
          If CAPEN > 0 Then
            CIIX% = XVALO(!cod_inte)
            EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    '
    Call ABRESERVA
    SQLX$ = " SELECT CantRes, CANTCONS, COD_INTE FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
'    Dim ReserTemp As ADODB.Recordset
'    Set ReserTemp = New ADODB.Recordset
'    ReserTemp.CursorType = adOpenDynamic
'    ReserTemp.LockType = adLockReadOnly
'    ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    '
    Set ReserTemp = READSET(SQLX$)
    With ReserTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          TORE# = XVALO(!CANTRES)
          CACO# = XVALO(!CANTCONS)
          '
          If TRIM$(UCase$(Unimed$(CI%))) = "GK" Then
             TORE# = TORE# / 1000
             CACO# = CACO# / 1000
          End If
          '
          SIGNO% = Sgn(CVS(Mid$(X$, 75, 4)))
          SARE = TORE# - CACO#
          If SARE * SIGNO% < 0 Then
            SARE = 0
          End If
          CAPEN = SARE
          If Abs(CAPEN) > 0 Then
            CIIX% = XVALO(!cod_inte)
            EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    Call ABREPEDCLI
    SQLX$ = " SELECT NROPED, NroClie, CanPed , CantEnt, CodiInt FROM PEDDETA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Status < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
    SQLX$ = SQLX$ + "AND VenRep = 'V' "
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.CursorType = adOpenDynamic
    PEDTEMP.LockType = adLockReadOnly
    If InStr(EPAC$, "GABAL") > 0 Then
         PEDTEMP.Open SQLX$, FLEXCONN, , , adCmdText
       Else
         Set PEDTEMP = READSET(SQLX$)
    End If
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NROX& = XVALO(!NROPED)
          NCI% = XVALO(!NROCLIE)
          CAPEN = XVALO(!CanPed) - XVALO(!CANTENT)
          '
          If CAPEN > 0 Then
            CIIX% = XVALO(!CODIINT)
            EXIS(CIIX%, 5) = EXIS(CIIX%, 5) + CAPEN
          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    If NIVEFIL% = 4 Then GoTo 100 ' situacion productos terminados
    '
    '
50  J& = 0
    NVII& = NUMAS%
    For i& = 1 To NVII&
      Call TRACE(20, i&, NVII&)
      CIIX% = i&
      If STATUIT%(CIIX%) > 0 Then
        CODE$ = CODEXT$(CIIX%)
        If CODE$ >= COD1$ Then
          If CODE$ <= COD2$ Then
            If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
              X$ = REGBLAN$("STODISP")
              Call REPLA(X$, MKI$(CIIX%), 1, 2)
              Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
              SUMA = 0: HAYMO% = 0
              For k% = 1 To 6
                SIG% = 1: If k% >= 4 Then SIG% = (-1)
                SUMA = SUMA + SIG% * EXIS(CIIX%, k%)
                If Abs(EXIS(CIIX%, k%)) >= 0.05 Then HAYMO% = 1
                Call REPLA(X$, MKS$(EXIS(CIIX%, k%)), 4 * k% + 1, 4)
              Next k%
              If SUMA > 0.05 Then
                  Call REPLA(X$, MKS$(SUMA), 29, 4)
                ElseIf SUMA < (-0.05) Then
                  Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                  'Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
              End If
              '
              If (NIVEFIL% = 1) Or (NIVEFIL% = 2 And HAYMO% = 1) Or (NIVEFIL% = 3 And SUMA <= (-0.05)) Then
                Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                J& = J& + 1
                Call GRAREG("STODISP", X$, J&)
              End If
            End If
          End If
        End If
      End If
    Next i&
    '
60  Call SETULTREG("STODISP", J&)
    Call CIERRARCH("STODISP")
    Call FILESORT("STODISP", "", 1, 1, "ASC")
    '
    Call CIERRARCH("*.*")
    OcoTemp.Close
    Set OcoTemp = Nothing
    AtenTemp.Close
    Set AtenTemp = Nothing
    COBRECEPTMP.Close
    Set COBRECEPTMP = Nothing
    OfaTemp.Close
    Set OfaTemp = Nothing
    ReserTemp.Close
    Set ReserTemp = Nothing
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
    Call GENEPLACO
    '
    If Option5.Value = True Then ' COBERTURA
      Call ANACOBER
      Exit Sub
    End If
    '
    Call CARGATABLADISP 'CARGA LA TABLA SQL DISPONIB CON LA INFORMACION DEL LISTADO
    
    If DISPSTK07.Visible = True Then '  VALIDACION PARA SI SE MUESTRA EL GEN.REP. PARA IMPRIMIR
        If Control("SITUASTO", "DISPONIB") = "SI" Then
           OPX% = COMALTER%("Opciones de  Listado\\ Listado Disponibilidad General\Disponibilidad con Detalle de O.Compra")
           If OPX% = 2 Then
              Call FINAL("?DISPONIB")
              Exit Sub
           End If
        End If
        '
        Call HEADERS("STODISP", "")
        If Option11.Value = True Then
             Call CONECRUN("*STODISC", "")
           ElseIf Option12.Value = True Then
             If InStr(1, EMPREAC$, "ENVAPLAST") > 0 Then
               Call FINAL("?DISPONI2") 'SE PUEDE PONER A MODO GENERAL, PERO VA A VER QUE VERIFICAR QUE ESTE EL REPORTE
             Else
               Call CONECRUN("*STODISL", "")
             End If
             
        End If
    End If
    Exit Sub
    
    '
100 Call ABRESTRUC
    SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) ORDER BY CODICONJ"
    Set FormTemp = READSET(SQLX$)
    With FormTemp
      Do While Not .EOF
          CIIX% = XVALO(!CODICONJ)
          If CIIX% > 0 Then
            PROTER%(CIIX%) = 1
          End If
          CIIX% = XVALO(!codielem)
          If CIIX% > 0 Then
            PROTER%(CIIX%) = 0
          End If
      .MoveNext
      Loop
   End With
   FormTemp.Close
   Set FormTemp = Nothing

'    With Formulas
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'          CIIX% = !CODICONJ
'          If CIIX% > 0 Then
'            PROTER%(CIIX%) = 1
'          End If
'        .MoveNext
'        Loop
'      End If
'      '
'      .MoveFirst
'      If Err < 1 Then
'        Do While Not .EOF
'          CIIX% = !CODIELEM
'          If CIIX% > 0 Then
'            PROTER%(CIIX%) = 0
'          End If
'        .MoveNext
'        Loop
'      End If
'      '
'    End With
    '
    J& = 0
    NVII& = NUMAS%
    For i& = 1 To NVII&
      Call TRACE(20, i&, NVII&)
      CIIX% = i&
      If STATUIT%(CIIX%) > 0 Then
        If PROTER%(CIIX%) = 1 Then
          CODE$ = CODEXT$(CIIX%)
          If CODE$ >= COD1$ Then
            If CODE$ <= COD2$ Then
              If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
                X$ = REGBLAN$("STODISP")
                Call REPLA(X$, MKI$(CIIX%), 1, 2)
                Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
                SUMA = 0: HAYMO% = 0
                For k% = 1 To 6
                  SIG% = 1: If k% >= 4 Then SIG% = (-1)
                  SUMA = SUMA + SIG% * EXIS(CIIX%, k%)
                  If Abs(EXIS(CIIX%, k%)) >= 0.05 Then HAYMO% = 1
                  Call REPLA(X$, MKS$(EXIS(CIIX%, k%)), 4 * k% + 1, 4)
                Next k%
                If SUMA > 0.05 Then
                    Call REPLA(X$, MKS$(SUMA), 29, 4)
                  ElseIf SUMA < (-0.05) Then
                    Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                    'Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
                End If
                Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                J& = J& + 1
                Call GRAREG("STODISP", X$, J&)
                '
              End If
            End If
          End If
        End If
      End If
    Next i&
    GoTo 60
    '

End Sub

'
Sub GENEPLACO()
    '
    JC& = 0: JF& = 0
    FIN& = ULTREG&("stodisp")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("STODISP", U&)
       CANTX = CVS(Mid$(XX$, 33, 4))
       If CANTX >= 0.05 Then
          CI% = CVI(Left$(XX$, 2))
          PHB% = PROVHABI%(CI%)
          FF% = HOY(HOS$)
          COSUN = COSPESOS(CI%)
          Call AREXPLO(CI%, 1, 1)
          If CAIT% <= 1 Then
               Z$ = REGBLAN$("PLACOMP")
               Call REPLA(Z$, MKI$(CI%), 1, 2)
               Call REPLA(Z$, MKI$(FF%), 3, 2)
               Call REPLA(Z$, PERIO$, 5, 16)
               Call REPLA(Z$, MKS$(CANTX), 21, 4)
               Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
               Call REPLA(Z$, MKS$(0), 29, 4) ' filler
               Call REPLA(Z$, MKI$(PHB%), 33, 2)
               Call REPLA(Z$, MKS$(COSUN), 35, 4)
               Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
               '
               JC& = JC& + 1
               Call GRAREG("PLACOMP", Z$, JC&)
               '
               FF1$ = REGBLAN$("OCOMEMI")
               Call REPLA(FF1$, MKI$(CI%), 1, 2)
               Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
               Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
               FEX = FF%
               Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
               Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
               Call REPLA(FF1$, MKI$(FF%), 57, 2)
               Call REPLA(FF1$, MKI$(CI%), 59, 2)
               Call GRAREG("OCOMEMI", FF1$, JC&)
               '
            Else
               '
               Z$ = REGBLAN$("PLAFABR")
               Call REPLA(Z$, MKI$(CI%), 1, 2)
               Call REPLA(Z$, MKI$(FF%), 3, 2)
               Call REPLA(Z$, PERIO$, 5, 16)
               Call REPLA(Z$, MKS$(CANTX), 21, 4)
               Call REPLA(Z$, MKS$(0), 29, 4)  ' FILLER
               JF& = JF& + 1
               Call GRAREG("PLAFABR", Z$, JF&)
               '
               FF1$ = REGBLAN$("ORFAEMI")
               Call REPLA(FF1$, MKI$(CI%), 1, 2)
               Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
               Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
               FEX = FF%
               Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
               Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
               Call REPLA(FF1$, MKI$(FF%), 57, 2)
               Call REPLA(FF1$, MKI$(CI%), 59, 2)
               Call GRAREG("ORFAEMI", FF1$, JF&)
               '
          End If
       End If
    Next U&
    '
    Call SETULTREG("PLACOMP", JC&)
    Call SETULTREG("OCOMEMI", JC&)
    Call SETULTREG("PLAFABR", JC&)
    Call SETULTREG("ORFAEMI", JC&)
    '
    Call CARPROSEL
    '
End Sub

Sub CARPROSEL()
   '
   Call COPYSTRU("PROVED1", "PROVEPLA")
   ASELPRO$ = "PROVEPLA"
   Screen.MousePointer = 11
   PPXXZ$ = String$(32767, 0)
   For U& = 1 To ULTREG("PLACOMP")
     XX$ = REGLEIDO$("PLACOMP", U&)
     NPRX% = CVI(Mid$(XX$, 33, 2))
     If NPRX% > 0 Then
       Mid$(PPXXZ$, NPRX%, 1) = Chr$(1)
     End If
   Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("ACREDOR")
     XX$ = REGLEIDO$("ACREDOR", U&)
     NPRX% = CVI(Left$(XX$, 2))
     If NPRX% > 0 Then
       If Asc(Mid$(PPXXZ$, NPRX%, 1)) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("PROVEPLA", XX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("PROVEPLA", JJ&)
   Call CIERRARCH("PROVEPLA")
   '
   Screen.MousePointer = 1
   '
End Sub



