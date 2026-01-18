VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form GENORPRO11 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Destinos de  Fabricación"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   210
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   3375
      Left            =   6120
      ScaleHeight     =   3315
      ScaleWidth      =   1320
      TabIndex        =   30
      Top             =   3440
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   31
         Top             =   780
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "GENORPRO11.frx":0000
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   1050
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Produccion Tercerista "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   165
      TabIndex        =   21
      Top             =   3360
      Width           =   5775
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
         Left            =   5355
         TabIndex        =   22
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text5 
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
         Left            =   945
         MaxLength       =   32
         TabIndex        =   27
         ToolTipText     =   "Este Campo es OPCIONAL"
         Top             =   1080
         Width           =   2115
      End
      Begin VB.TextBox Text6 
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
         Left            =   4215
         MaxLength       =   16
         TabIndex        =   26
         ToolTipText     =   "Este Campo es OBLIGATORIO, solo si se ha ingresado el CLIENTE"
         Top             =   1095
         Width           =   1320
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   4545
         TabIndex        =   23
         ToolTipText     =   "Este Campo es OBLIGATORIO"
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   1110
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/Compra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3210
         TabIndex        =   28
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label Label2 
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
         Left            =   960
         TabIndex        =   25
         Top             =   660
         Width           =   4605
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2985
         TabIndex        =   24
         Top             =   315
         Width           =   1545
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   690
      Left            =   6120
      Picture         =   "GENORPRO11.frx":1F22
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Graba - Aprueba la Transacción"
      Top             =   2625
      Width           =   780
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos de Produccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   165
      TabIndex        =   11
      Top             =   1680
      Width           =   5775
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   35
         ToolTipText     =   "Este Campo es OBLIGATORIO"
         Top             =   1200
         Width           =   2280
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
         Height          =   285
         Left            =   5370
         TabIndex        =   19
         Top             =   1200
         Width           =   175
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   3
         ToolTipText     =   "Este Campo es OBLIGATORIO"
         Top             =   810
         Width           =   1440
      End
      Begin VB.TextBox Text2 
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
         Left            =   975
         MaxLength       =   32
         TabIndex        =   1
         ToolTipText     =   "Este Campo es OPCIONAL"
         Top             =   345
         Width           =   2100
      End
      Begin VB.TextBox Text1 
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
         Left            =   4245
         MaxLength       =   16
         TabIndex        =   2
         ToolTipText     =   "Este Campo es OPCIONAL"
         Top             =   330
         Width           =   1320
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Presentacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   50
         TabIndex        =   36
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   4245
         TabIndex        =   20
         ToolTipText     =   "Este Campo es OPCIONAL"
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Entrega:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   3600
         TabIndex        =   18
         Top             =   900
         Width           =   2370
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº.OP.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -480
         TabIndex        =   14
         Top             =   390
         Width           =   1400
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   300
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Partida:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   390
         Width           =   1545
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Código Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   165
      TabIndex        =   7
      Top             =   465
      Width           =   5805
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
         Height          =   285
         Left            =   5370
         TabIndex        =   4
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   3240
         TabIndex        =   0
         ToolTipText     =   "Este Campo es OBLIGATORIO"
         Top             =   360
         Width           =   2160
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   5445
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   1680
         TabIndex        =   9
         Top             =   430
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2535
      Left            =   6120
      ScaleHeight     =   2475
      ScaleWidth      =   1320
      TabIndex        =   6
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton AGenRep 
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
         Left            =   85
         Picture         =   "GENORPRO11.frx":222C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   840
         Width           =   690
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
         Left            =   85
         Picture         =   "GENORPRO11.frx":2536
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   690
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "GENORPRO11.frx":2680
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "GENORPRO11.frx":28B4
      TabIndex        =   34
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label7 
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
      Left            =   4635
      TabIndex        =   17
      Top             =   180
      Width           =   1305
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero S/F:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   16
      Top             =   195
      Width           =   1485
   End
End
Attribute VB_Name = "GENORPRO11"
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
    '
    Call CARGALISEL("!ECOMAST", "MASTER INNER JOIN FORMULAS ON MASTER.CODINT = FORMULAS.CodIConj", "CODIGO/A16;DESCRIPCION/A46", "CODINT > 0 AND MASTER.STATUS > 0", "DISTINCT")
    Call SELECHO("!ECOMAST")
    If VALACT1$("!ECOMAST") <> "" Then
       Text7 = VALACT1$("!ECOMAST")
    End If
    '
End Sub

Private Sub Command11_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Command14_Click()
    Command14.Enabled = False
    If DERACCE%("STPOFAB", "Configuracion de Funcionamiento de O/Fabricacion") = 2 Then
        OPORFA09.Show 1
    End If
    Command14.Enabled = True
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   Call ABREORFAB

   'VALIDACIONES
   If ValiDato%(Text7, "ESCODI", "Código no Válido") < 1 Then GoTo 99
   '
   CI% = XVALO(CODINT(Text7))
   If TRIM$(Text2) <> "" Then
     CONDI$ = "Destino ='" & TRIM$(Text2) & "'"
     If CantRegistros("ORDEFABR", CONDI$) > 0 Then
       CIXXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", CONDI$, "NOMESS"))
       OPT1% = COMALTER%("Nro de O.P " & TRIM$(Text2) & " Existente para Código " & CODEXT$(CIXXI%) & ".\¿Desea Continuar?\\No, Cancelar\Si, Continuar")
       If OPT1% <> 2 Then
          Text2.SetFocus
          GoTo 99
       End If
     End If
   End If
   '
   CONDI$ = "CodIConj =" & CODINT(Text7)
   If CantRegistros("FORMUPAT", CONDI$) <= 1 Then
      Call COMUNI("Codigo Debe Tener Formula Patron")
      Text7.SetFocus
      GoTo 99
   End If
   '
   'VALIDA QUE LA FORMULA PATRON ESTE APROBADA
   If FEREFOR%(CI%) <= FECHANUM("01/01/80") Then
      Call MENSERR(24, "Imposible Continuar.\Formula Patron No Aprobada")
      GoTo 99
   End If
   '
   'Cantidad
   If XVALO(Text3) < 0.005 Then
      Call COMUNI("Cantidad No Válida")
      Text3.SetFocus
      GoTo 99
   End If
   '
   ' NUMERO DE ORDEN DE PRODUCCION
   If TRIM$(Text2) = "" Then
      Call MENSERR(24, "Falta Numero de Orden de Produccion.")
      Text2.SetFocus
      GoTo 99
   End If
   If CantRegistros("ORDEFABR", "NUORPROD = '" & TRIM$(Text2) & "'") > 0 Then
      Call MENSERR(24, "Numero de Orden de Produccion Repetido.")
      Text2.SetFocus
      GoTo 99
   End If
   '
   ' NUMERO DE LOTE O PARTIDA
   If TRIM$(Text1) = "" Then
      Call MENSERR(24, "Falta Numero de Lote o Partida.")
      Text1.SetFocus
      GoTo 99
   End If
   If CantRegistros("ORDEFABR", "IDENLOTE = '" & TRIM$(Text1) & "'") > 0 Then
      Call MENSERR(24, "Numero de Lote o Partida Repetido.")
      Text1.SetFocus
      GoTo 99
   End If
   '
   'cliente
   If TRIM$(text4) <> "" Then  ' ACEPTA QUE NO HAYA CLIENTE PERO SI EXISTE LO VALIDA
      NC% = XVALO(text4)
      If NC% < 1 Or RASOCLI$(NC%) = "" Then
         Call COMUNI("Cliente no Válido")
         GoTo 99
      End If
   End If
   '
   If NC% > 0 Then
      If TRIM$(Text6) = "" Then
         Call COMUNI("Falta Ingresar O/Compra")
         GoTo 99
      End If
   End If
   '
   'VALIDA QUE LA CANTIDAD NO DIFIERA DE LA PRESENTADA ORIGINALMENTE
   CANTI = XVALO(Text3)
'   LOTESTAN (CI%)
'   If XVALO(Text3) <> CANTI Then
'      Call COMUNI("Lote de Producción Debe Coincidir con Lote Standard")
'      GoTo 99
'   End If
   '
   PRESENTA$ = TRIM$(Text8)
   If PRESENTA$ = "" Then
      Call MENSERR(24, "Falta Presentacion")
      GoTo 99
   End If
   '
   'Si el usuario no Ingresa Fecha de Entrega, se asigna la de hoy
   FECHENTRE = TRIM$(Label11)
   If FECHENTRE = "" Then
      HOII% = HOY(HOS$)
      FECHENTRE = CVDAT(HOII%)
      Label11 = HOS$
   End If
  
   'Numero de Remito
   NUMREMPRO$ = TRIM$(Text5)
   
   'Numero Orden de compra
   NUMOCO$ = TRIM(Text6)
   '
   CM$ = CONTROL("", "CODMORFA")
   CONDI$ = "IdSubOr LIKE '" & CM$ & "%'"
   NRO& = 1 + XVALO(VALOBADA("OrdeFabr", "MAX(NumeOrFa)", CONDI$, "NOMESS"))
   IDOF$ = TRIM$(Str$(NRO&))
   While Len(IDOF$) < 6
      IDOF$ = "0" + IDOF$
   Wend
   IDOF$ = CM$ + "-" + IDOF$ + "-1"
   '
   DESCRI$ = Left$(Label2, 64)
   HOII% = HOY(HOS$)
   FECH = CVDAT(HOII%)
   NUMOP& = XVALO(Text2)
   '
   SQLX$ = " SELECT * FROM OrdeFabr "
   Dim ORDEFTMP As ADODB.Recordset
   Set ORDEFTMP = New ADODB.Recordset
   ORDEFTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With ORDEFTMP
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
          !NUMEORFA = NRO&
             CAM2$ = "OP." & TRIM$(Text2) & " - L." & TRIM$(Text1) & " - " & FANTCLI$(NC%) & " - " & DESCRIT0$(CI%)
          !REFEREN = Left(CAM2$, 64)
          !IdSubOr = IDOF$
          !FechGen = FECH
          !Cod_Inte = CI%
          !NUORPROD = TRIM$(Text2)
          !DESTINO = TRIM$(Label2) ' CLIENTE
          !idenLote = TRIM$(Text1) ' LOTE
          !COD_EXTE = CODEXT$(CI%)
          !Descrip = DESCRIT0$(CI%)
          !Unidad = UNIMED$(CI%)
             If PRESENTA$ <> "" Then !Unidad = PRESENTA$
          !CanProy = XVALO(Text3)
          !STATUORF = 1
          !FechLan = FECH
          !CanApro = 0
          !CanRech = 0
          !IdenRetra = ""
          !FechProyEntre = CVDAT(0)
          !FechTempEntre = CVDAT(0)
          !FechRecal = CVDAT(0)
          !Observa = ""
          !Nume_Ocom = NUMOCO$
          !NumRem_Prov = NUMREMPRO$
          !Priorid = 0
          !Nume_Cli = NC%
          !FechEnSol = FECHENTRE
          .Update
   End With
   ORDEFTMP.Close
   Set ORDEFTMP = Nothing

   '   GENERACION DE ECORDEF
   Call GENLISORF
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Call COMUNI("Proceso Terminado")
   '
90 Call BLANFORMU
99 Screen.MousePointer = 1
   Command2.Enabled = True
   '
End Sub

Private Sub Command3_Click()
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   text4.TEXT = TRIM$(Str$(NCLIE%))
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label11 = VDEV$
   End If
   '
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
   Call ABRECONEXION
   Call MostrarNumOF
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
         Text3.SetFocus
    End If
    '
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    '
     If KeyAscii = 13 Then
          Text1.SetFocus
     End If
     '
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      text4.SetFocus
   End If
   '
End Sub

Private Sub Text3_LostFocus()
   Text3 = FORMATNUM((Text3), "F13.0")
End Sub

Private Sub Text4_Change()
   Label2 = ""
   
   NC% = XVALO(text4.TEXT)
   If NC% > 0 Then
     If NC% <= 999999 Then
       If STATCLI(NC%) > 0 Then
         Label2 = RASOCLI$(NC%)
       End If
     End If
   End If

End Sub

Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command3_Click
   End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then ' ENTER
      Text5.SetFocus
 End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then ' ENTER
      Text6.SetFocus
 End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ' ENTER
      Command2.SetFocus
 End If
End Sub

Private Sub Text7_Change()
    Label1 = ""
    Text3 = ""
    CI% = CODINT%(Text7.TEXT)
    If CI% > 0 Then
      Label1 = DESCRIT$(CI%)
      Text3 = FORMATNUM(LOTESTAN(CI%), "F13.0") ' LOTE STANDARD
      If XVALO(Text3) = 0 Then
         COMUNI ("Imposible Continuar, Debe Informar Cant. de Lote Standard para Codigo " & TRIM(Text7))
         Call BLANFORMU
      End If
      
    End If
End Sub
Sub BLANFORMU()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Label1 = ""
   Label2 = ""
   Label11 = ""
  Call MostrarNumOF
End Sub

Private Sub Text7_DblClick()
  Call Command1_Click
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 113 Then ' CORRESPONDE TECLA <F2>
     Call Command1_Click
   End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then ' ENTER
      Text2.SetFocus
 End If

End Sub
Sub MostrarNumOF()
   CM$ = CONTROL("", "CODMORFA")
   CONDI$ = "IdSubOr LIKE '" & CM$ & "%'"
   NRO& = 1 + XVALO(VALOBADA("OrdeFabr", "MAX(NumeOrFa)", CONDI$, "NOMESS"))
   
   IDOF$ = TRIM$(Str$(NRO&))
   While Len(IDOF$) < 6
      IDOF$ = "0" + IDOF$
   Wend
   
   IDOF$ = CM$ + "-" + IDOF$ + "-1"
   'IDOF$ = "OF-" + IDOF$ + "-1"
   Label7 = IDOF$
End Sub

Sub GENERANOTAPE(NORFA$, PROV$)  'ATENCION ESTA RUTINA TAMBIEN SE ENCUENTRA EN PROYECTO MORCOM09
   '
   FEMI% = HOY(HOS)
   '
   ' AQUI AGREGA UN UNICO REGISTRO DE ENCABEZADO EN LA TABLA OCOMENCA
   Call APERBASDAT("COMPRAS")
   NUMOCOMPRA& = 1 + XVALO(VALOBADA("OCOMENCA", "Max(Nume_Pedi)", "Nume_pedi > 0", "NOMESS"))
   '
   CONDI$ = "IDSUBOR ='" & NORFA$ & "'"
   NOCOM$ = VALOBADA("ordefabr", "NUME_OCOM", CONDI$)
   If TRIM$(NOCOM$) = "" Then
      COMUNI ("Imposible Continuar\Falta O/Compra Para " & NORFA$)
      GoTo 99
   End If
   '
   'Variables utilizadas en Alta Detalle de Pedido
   YAENCA% = 0
   APRODIR% = 0
   If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then
     APRODIR% = 1
   End If
   ORDIT% = 0
   '
   Dim OCOMDETA As ADODB.Recordset
   Set OCOMDETA = New ADODB.Recordset
   OCOMDETA.Open FILTSQL$("select * from OCOMDETA where raso_prov ='x'"), FLEXCONN, adOpenKeyset, adLockPessimistic
   '
   Dim BALTEMP As ADODB.Recordset
   Set BALTEMP = New ADODB.Recordset
   BALTEMP.Open FILTSQL$("SELECT * FROM BALAREQMO WHERE IDENREQUE='X'"), FLEXCONN, adOpenKeyset, adLockOptimistic
   '
   CONDI$ = "Destino = '" & NORFA$ & "'"
   SQLX$ = "SELECT IDENREQUE,Cod_Inte,Sum(Kilo_Reque) as Sureque,Sum(Kilo_Asig) as Suasig FROM BALAREQMO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " GROUP BY IDENREQUE,COD_INTE"
   '
   Dim BALAREQ As ADODB.Recordset
   Set BALAREQ = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With BALAREQ
      Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         If CRITREP%(CIXI%) = 1 Then
            If XVALO(!SUREQUE) > XVALO(!SUASIG) Then
              If YAENCA% < 1 Then
                 GoSub GENENCAOCOM
                 If YAENCA% < 1 Then Exit Sub
              End If
              '
              CODE$ = CODEXT$(CIXI%)
              '
              'Registro de Detalle Pedido
              '
              ORDIT% = ORDIT% + 1
              CI0% = XVALO(!Cod_Inte)
              
              OCOMDETA.AddNew   ' YA CONVERTIDO
                      On Error Resume Next
              OCOMDETA!CodiEmpr = CodiEmp%
                      On Error GoTo 0
              OCOMDETA!Nume_Pedi = NUMOCOMPRA&
              OCOMDETA!NuOrd_Item = ORDIT%
              OCOMDETA!Femi_Ocom = CVDAT(FEMI%)
              OCOMDETA!Nume_Cli = XVALO(PROV$)
              OCOMDETA!Raso_Prov = Left$(RASOCLI$((-1) * XVALO(PROV$)), 64) 'OT-06-0227-WAR-05/06/06///
              OCOMDETA!Medi_Ocom = PedInt$
              OCOMDETA!Cod_Inte = CI0%
              OCOMDETA!COD_EXTE = CODEXT$(CI0%)
              OCOMDETA!Descritem = DESCRIT0$(CI0%)
              OCOMDETA!Delibre = DESCRIT0$(CI0%)
              OCOMDETA!Valo_Unid = 0
              OCOMDETA!Uni_Mas = UNIMED$(CI0%) 'Mid$(ZZ$, 11, 4)
                CAUNIS# = XVALO(!SUREQUE) - XVALO(!SUASIG) ' CANT. EN UNIDADES DE STOCK
              OCOMDETA!Unids_Ocom = CAUNIS#
                'CANTU = XVALO(!KILO_SALDO)  ' CANT EN UNIDADES DE COMPRA
                CANTU = REDON_INDIVIS(CI0%, CAUNIS#)
              OCOMDETA!CANT_OCOM = CANTU
              OCOMDETA!Uni_Com = UNIMED$(CI0%) ' unidad de compra
              OCOMDETA!Coef_Unids = 1
              OCOMDETA!Pre_Unit = 0
              OCOMDETA!Desc_Item = 0
              OCOMDETA!PrUn_Neto = 0
              OCOMDETA!Inet_Item = 0
              OCOMDETA!Fentre_Sol = FECHATEX(HOY(HO$))
              OCOMDETA!CONSIMAT = ""
              OCOMDETA!Cant_Rec = 0
              OCOMDETA!Cod_Oper = 0
              OCOMDETA!Doc_Orig = ""
              OCOMDETA!N_OrServ = 0
              OCOMDETA!Mone_Cos = 1        ' antes pos 83,1
              OCOMDETA!Mone_Emis = 1        ' antes pos 84,1
              OCOMDETA!Tsum_Ocom = 0              ' materiales
              OCOMDETA!Stat_Ocom = APRODIR% - 1  ' status Pendiente de aprobacion
              OCOMDETA!IdenReque = SAFETEXT(!IdenReque)
              OCOMDETA!IContab = 0 ' Imputacion contable
              OCOMDETA!Sect_Ocom = ""
              OCOMDETA.Update
              'Fin Registro de Pedido
              '
              '
              'Se agrega Registro en Balareqmo Informando el Pedido Generado
              NROPED$ = "NP." & SAFETEXT(NUMOCOMPRA&)
              CONDI$ = "IDENREQUE = '" & SAFETEXT$(!IdenReque) & "' AND Kilo_Reque > 0"
              BALTEMP.AddNew   ' YA CONVERTIDO
                     On Error Resume Next
              BALTEMP!CodiEmpr = CodiEmp%
                     On Error GoTo 0
              BALTEMP!CODICOMPRO = SAFETEXT(VALOBADA("BALAREQMO", "CODICOMPRO", CONDI$, "NOMESS"))
              BALTEMP!NUMECOMPRO = XVALO(VALOBADA("BALAREQMO", "NUMECOMPRO", CONDI$, "NOMESS"))
              BALTEMP!CODSUCU = XVALO(VALOBADA("BALAREQMO", "CODSUCU", CONDI$, "NOMESS"))
              BALTEMP!NUORITEM = XVALO(VALOBADA("BALAREQMO", "NUORITEM", CONDI$, "NOMESS"))
              BALTEMP!IdenReque = SAFETEXT(VALOBADA("BALAREQMO", "IDENREQUE", CONDI$, "NOMESS"))
              BALTEMP!Cod_Inte = XVALO(VALOBADA("BALAREQMO", "Cod_Inte", CONDI$, "NOMESS"))
              BALTEMP!idenLote = NROPED$
              BALTEMP!Kilo_Asig = CANTU ' XVALO(!KILO_SALDO)
              BALTEMP!FeReMovi = FECHATEX$(HOY(HO))
              BALTEMP!fechMov = FECHATEX$(HOY(HO))
              BALTEMP!CREADOPOR = "FLEX"
              BALTEMP!DESTINO = SAFETEXT(VALOBADA("BALAREQMO", "Destino", CONDI$, "NOMESS"))
              BALTEMP.Update
              '
            End If
         End If
      .MoveNext
      Loop
    End With
    OCOMDETA.Close
    Set OCOMDETA = Nothing
    BALTEMP.Close
    Set BALTEMP = Nothing
    BALAREQ.Close
    Set BALAREQ = Nothing
    '
    'Ventana de Opcion Impresion o Email
    PORPRIN% = 1: PORMAIL% = 0
    NCIX% = XVALO(PROV$) * (-1)
    EMAI$ = TRIM$(EMAILCLI$(NCIX%))
    If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
       GoTo 50 'NO TIENE DIRECCION O DIRECCION NO VALIDA
    End If
    '
12  SALIPED07.Show 1
    Call FRESHALL
    If SALIPED07.Check1.Value < 1 Then
       If SALIPED07.Check2.Value < 1 Then
          Call MENSERR(24, "Debe Marcar al Menos\una Dirección de Salida.")
          GoTo 12
       End If
    End If
    PORPRIN% = 0: If SALIPED07.Check1.Value = 1 Then PORPRIN% = 1
    PORMAIL% = 0: If SALIPED07.Check2.Value = 1 Then PORMAIL% = 1
    '
50  If PORPRIN% > 0 Then      ' EMITE POR IMPRESORA
       Call PRICOMPRA(NUMOCOMPRA&)
    End If
    '
    If PORMAIL% > 0 Then      ' ENVIA POR MAIL
        SENDBYMAIL$ = EMAI$
        For Each PRX1 In Printers
           If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 25
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        '
25      Call PRICOMPRA(NUMOCOMPRA&)
        SENDBYMAIL$ = ""
    End If
    '
    If YAENCA% < 1 Then
         Call COMUNI("No se ha Generado Nota de Pedido Alguna.")
       Else
         Call COMUNI("Nota de Pedido Generada OK")
    End If
    '
Exit Sub
'
'
GENENCAOCOM:
   '
   Dim OCOMENCA As ADODB.Recordset
   Set OCOMENCA = New ADODB.Recordset
   SQL$ = "Select * from OCOMENCA WHERE NUME_PEDI = " & NUMOCOMPRA&
   OCOMENCA.Open FILTSQL$(SQL$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With OCOMENCA
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
       !Nume_Pedi = NUMOCOMPRA&
          REFE$ = FECHATEX$(HOY(HO$)) + " - " + RASOCLI$(XVALO(PROV$))
       !Refe_Ocom = Left$(REFE$, 64)
       !Femi_Ocom = CVDAT(FEMI%)
       !Nume_Cli = XVALO(PROV$) 'NPRO%
       '!RASO_PROV = Left$(RASOCLI$(XVALO(PROV$) * (-1)), 48) 'Left$(RASOCLI$((-1) * NPRO%), 48) 'OT-06-0227-WAR-05/06/06///
       !Raso_Prov = Left$(RASOCLI$(XVALO(PROV$)), 48)    ' SE QUITA EL "*(-1)" PORQUE LAS N/P SON EN BASE A CLIENTES
       !Tsum_Ocom = 0               ' materiales
          STATUOC% = (-1)           ' SIGNIFICA PENDIENTE DE APROBACIÓN
          If Left$(CONTROL$("OCOMPRA", "MODOAPRO"), 6) = "DIRECT" Then STATUOC% = 0
       !Stat_Ocom = STATUOC%        ' PASA A 0 (CERO) SI APROBACION DIRECTA
       '
       !Fgen_Ocom = Now
       !Usua_Gen = USNBR%
          MONECOI% = 1 ' XVALO(Text15)
       !Mone_Cos = MONECOI%        ' MONEDA PARA COSTEO
       !Mone_Emis = MONECOI%        ' MONEDA DE EMISIÓN
       !Tcam_Ocom = 0               ' TIPO DE CAMBIO / COTIZACION
       !Cpag_Ocom = 0
       !Embala_Ocom = 0
       !AtSr_Ocom = ""
          '
          PORDESCUENTO$ = 0
          DESC$ = 0
          TOTANE# = 0
          IDESCUEN# = 0
          TOTANEDES# = 0
          TOTIVA# = 0
          TOTPER# = 0
          TOTATO# = 0
          '
      !ToBru_Ocom = TOTANE#
      !Pdes_Ocom = PORDESCU
      !IDes_Ocom = IDESCUEN#
      !Iiva_Ocom = TOTIVA#
      !Iotr1_Ocom = TOTPER#
      !Total_Ocom = TOTATO#
      !Lentrega = ""
      !Obse_Ocom = ""
      !Anex_Ocom = ""
      !PorcIva = ""
      !Anex_Ocom = SAFETEXT(NORFA$)
      !ID_COMPROB = "NP." + SAFETEXT(NUMOCOMPRA&)
      .Update
   End With  'fin de OcomEnca
   OCOMENCA.Close
   Set OCOMENCA = Nothing
   YAENCA% = 1
Return
   '
99 End Sub

Sub PRICOMPRA(OCOM&)
    '
    PRIPLANOI% = 0
    PRIESPEC% = 0
    PRIHRINS% = 0
    '
    ' IMPRESION DE LA O/COMPRA
    '
    FORIPRE$ = CONTROL$("", "FORNOPE")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHEMII% = HOY(HO$) 'FECHANUM(Text9)
      FECHDOC$ = FETEXTO$(FECHEMII%)
      NUMEDOC$ = OCOM&
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "NUME-COM"
      CODPARAM$(2) = "FECH-EMI"
      CODPARAM$(3) = "DOCUORIG"
      CODPARAM$(4) = "DESTINO"
      CODPARAM$(5) = "NRO-LOTE"
      CODPARAM$(6) = "N-REMITO"
      CODPARAM$(7) = "N-OCOMPR"
      CODPARAM$(8) = "RASO-CLI"
      CODPARAM$(9) = "DOMI-CLI"
      CODPARAM$(10) = "CPOS-CLI"
      CODPARAM$(11) = "TRANSPOR"
      CODPARAM$(12) = "LOCA-CLI"
      
      CAPARDOC% = 12
      '
      'Encabezado
      '
      CONDIOCOM$ = "NUME_PEDI = " & OCOM&
      '
      NORFA$ = VALOBADA("OCOMENCA", "ANEX_OCOM", CONDIOCOM$)
      '
      CONDIFAB$ = "IDSUBOR ='" & NORFA$ & "'"
      '
      FECH$ = FECHATEX$(CVINT%(VALOBADA("OCOMENCA", "FGEN_OCOM", CONDIOCOM$)))
      '
      DESTI$ = VALOBADA("ORDEFABR", "DESTINO", CONDIFAB$)
      LOTE$ = VALOBADA("ORDEFABR", "IDENLOTE", CONDIFAB$)
      REMI$ = VALOBADA("ORDEFABR", "NumRem_Prov", CONDIFAB$)
      OCOMCLI$ = VALOBADA("ORDEFABR", "NUME_OCOM", CONDIFAB$)
      NPROV% = XVALO(VALOBADA("OCOMENCA", "NUME_CLI", CONDIOCOM$))
      
      NUTRA% = NUTRANCLI(NPROV%)
      DENOTRA$ = DENOTRANS$(NUTRA%)
     
     
      DOCPARAM$(1) = NUMEDOC$
      'DOCPARAM$(2) = FECH$
      DOCPARAM$(3) = NORFA$
      DOCPARAM$(4) = DESTI$                                    '  Text12
      DOCPARAM$(5) = LOTE$
      DOCPARAM$(6) = REMI$
      DOCPARAM$(7) = OCOMCLI$                                  ' Mid$(X1$, 147, 20)
      DOCPARAM$(8) = RASOCLI$(NPROV%)                          ' Mid$(X1$, 67, 40)
      DOCPARAM$(9) = DOMICLI$(NPROV%)                          ' Mid$(X1$, 167, 40)
      DOCPARAM$(10) = CPOSCLI$(NPROV%)                         ' Mid$(X1$, 107, 40)
      DOCPARAM$(11) = DENOTRA$
      DOCPARAM$(12) = LOCACLI$(NPROV%)
      
      
      HOII% = HOY(HOS$)
      FECHENT% = HOII%
'      DOCPARAM$(15) = Text14
'      If FECHENT% = HOII% Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (inmediato)"
'        ElseIf FECHENT% < HOII% + 4 Then
'          DOCPARAM$(15) = DOCPARAM$(15) + " (urgente!!)"
'      End If
'      DOCPARAM$(16) = TRIM$(Label11)
'      DOCPARAM$(17) = ""
'      If MONEMI% > 1 Then DOCPARAM$(17) = Str$(TICAMONE(MONEMI%))
'      DOCPARAM$(18) = Eco21(1)
'      DOCPARAM$(19) = Text10
'      DOCPARAM$(20) = Total(4)
'      '
'      ' esto es para poner el importe en letras
'      TTTT# = XVALO(Total(3))
'      TOTFAC$ = NUMTEX$(TTTT#)
'      If Label11 <> "" And Label11 <> "Pesos" Then
'         PPXX% = InStr(TOTFAC$, "PESOS")
'         If PPXX% > 0 Then
'           TOTFAC$ = Left$(TOTFAC$, PPXX% - 1) + UCase$(Label11) + Mid$(TOTFAC$, PPXX% + 5)
'         End If
'      End If
'      Call FRATEXTO(TOTFAC$, 40)
'      CODPARAM$(22) = "IMP-LET1": DOCPARAM$(22) = RETEX$(1)
'      CODPARAM$(23) = "IMP-LET2": DOCPARAM$(23) = RETEX$(2)
'      CODPARAM$(24) = "IMP-LET3": DOCPARAM$(24) = RETEX$(3)
'      '
'      CODTABU2$(1) = "REF-MAT3"
'      CACOLTAB2% = 1
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "ORD-ITEM"
      CODTABU1$(6) = "COD-CONJ"
      CACOLTAB1% = 6
      '
      DESTIDOC% = (-1) * XVALO(Text1)
      CAITAB1% = 0
      CAITCRU% = 0
      INCLUSPE% = 1
      REPLAORI% = 0
      RECOPROP% = 0
      REVICODX% = 0
      '
      
      '
      Dim OCOMDETA As ADODB.Recordset
      Set OCOMDETA = New ADODB.Recordset
      SQLX$ = "SELECT * FROM OCOMDETA WHERE NUME_PEDI = " & OCOM&
      Set OCOMDETA = FLEXCONN.Execute(FILTSQL$(SQLX$))
      If OCOMDETA.EOF And OCOMDETA.BOF Then GoTo 999
      '
      With OCOMDETA
       .MoveFirst
       CAITAB1% = 0
       Do While Not (.EOF)
       CAITAB1% = CAITAB1% + 1
       CODDX$ = SAFETEXT(!COD_EXTE)    ' TRIM$(Text7)
       TETABU1$(1, CAITAB1%) = CODDX$
       DECRIT$ = SAFETEXT(!Descritem)   'Text18           ' TRIM$(DESCRIT0$(CI0%))
       TETABU1$(2, CAITAB1%) = DECRIT$
       TETABU1$(3, CAITAB1%) = SAFETEXT(!Uni_Mas)          ' Mid$(X$, 59, 4)
       TETABU1$(4, CAITAB1%) = XVALO(!CANT_OCOM)
       TETABU1$(5, CAITAB1%) = XVALO(!NuOrd_Item)
       TETABU1$(6, CAITAB1%) = NUPLANO$(XVALO(!Cod_Inte))
       '
       .MoveNext
       Loop
      End With
      OCOMDETA.Close
      Set OCOMDETA = Nothing
      '
      Call PRINTDOC("FORNOPE")                  'Pedidos  ' Orden de Compra
      '
      End If
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub


Function PRONUOCO&()
   '
   PRONU& = 1
   Call APERBASDAT("COMPRAS")
   '
   SQLX$ = "SELECT Nume_Ocom FROM OcomEnca "
   SQLX$ = SQLX$ + "ORDER BY Nume_Ocom DESC;"
   'Set OEncaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim OEncaTemp As ADODB.Recordset
   Set OEncaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   With OEncaTemp
     .MoveFirst
     If Err < 1 Then
       PRONU& = 1 + !Nume_Ocom
     End If
   End With
   
   On Error GoTo 0
   Set OEncaTemp = Nothing
   '
   PRONUOCO& = PRONU&
   '
End Function

Sub ENVIMAILACOMPRA(NCLI%, SOLFA$)
   '
   MAILCOMEXT$ = TRIM$(CONTROL$("BOLRECEP", "COEXMAIL"))
   '
   mailCompras$ = TRIM$(CONTROL$("BOLRECEP", "COMPMAIL"))
   '
   mailCalidad$ = TRIM$(CONTROL$("BOLRECEP", "GECAMAIL"))
   '
   ASUNTO$ = "Solicitud de Fabricacion " & SOLFA$ & " Generada"
   '
   TEXTOMAI$ = "Se ha generado la Solicitud de Fabricaicon " & SOLFA$ & " correspondiente al cliente " & RASOCLI$(NCLI%) & CRLF$
   CONDI$ = "IdSubOr = '" & SOLFA$ & "'"
   NUMOPCL$ = VALOBADA("ORDEFABR", "DESTINO", CONDI$, "NOMESS")
   NULOTE$ = VALOBADA("ORDEFABR", "IDENLOTE", CONDI$, "NOMESS")
   CANTI = XVALO(VALOBADA("ORDEFABR", "CANPROY", CONDI$, "NOMESS"))
   CODI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$, "NOMESS"))
   TEXTOMAI$ = TEXTOMAI$ & "para el producto: " & TRIM$(CODEXT$(CODI%)) & " - " & TRIM$(DESCRIT$(CODI%)) + CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Numero de O/P: " & NUMOPCL$ + CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Lote: " & NULOTE$ & CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Cantidad Proyectada: " & CANTI & CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Verifique mediante el Re-Calculo de Necesidades si es necesario realizar una Nota de Pedido u Orden de Compra" & CRLF$
   '
   MAI_DESTI$ = TRIM$(mailCompras$)
   If PIVACLI(NCLI%) = 5 Then          'si el cliente es del exterior, segun posicion de iva, agrega el mail de Com. Ext.
        If MAILCOMEXT$ <> "" Then
           If MAI_DESTI$ <> "" Then MAI_DESTI$ = MAI_DESTI$ + "; "
           MAI_DESTI$ = MAI_DESTI$ + MAILCOMEXT$
        End If
        If mailCalidad$ <> "" Then
           If MAI_DESTI$ <> "" Then MAI_DESTI$ = MAI_DESTI$ + "; "
           MAI_DESTI$ = MAI_DESTI$ + mailCalidad$
        End If
   End If
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   '
   Call GENERAMAIL
   '
End Sub
