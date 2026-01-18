VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CCTEANEX14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente de Clientes (Pagos de Impuestos) CAJA ANEXA"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "Detalle de Cuenta Corriente (Pago de Impuestos) - CAJA ANEXA - "
      Top             =   2100
      Width           =   12015
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   12015
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   10695
         Begin VB.CommandButton Command1 
            Caption         =   "Todos"
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
            Left            =   2040
            TabIndex        =   20
            ToolTipText     =   "Seleccionar Rango de Fechas"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TXTHASTA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6675
            Locked          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Doble Click Para Limpiar Casillero"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   2
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Hasta: "
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TXTDESDE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6675
            Locked          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Doble Click Para Limpiar Casillero"
            Top             =   360
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Desde: "
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Sel. Fecha"
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
            Left            =   4800
            TabIndex        =   14
            ToolTipText     =   "Seleccionar Rango de Fechas"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   7680
            Picture         =   "CCTEANEX14.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Imprime Muestra del Formulario Activo"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   2880
         End
         Begin VB.TextBox TXTID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   3720
            Locked          =   -1  'True
            MaxLength       =   24
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "ID"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TXTNOMBRE 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   840
            Width           =   4695
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   960
            MaxLength       =   24
            TabIndex        =   0
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Text            =   "Cta."
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Text            =   "Nombre"
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton Command6 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   750
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   195
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "CCTEANEX14.frx":066A
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CCTEANEX14.frx":06C8
      Top             =   120
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   5235
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9234
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16501405
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   ""
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
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuAj 
      Caption         =   "Ajustes"
      Begin VB.Menu mnuajdev 
         Caption         =   "Ajustes por Devolución"
      End
   End
End
Attribute VB_Name = "CCTEANEX14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
 

Private Sub Command1_Click()
  Call CTACTE(0)
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
    OBX$ = ""
    If TRIM$(TXTDESDE) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTDESDE)), 1, 2)
    End If
    If TRIM$(TXTHASTA) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTHASTA)), 3, 2)
    End If
    VDEF$ = OBX$
10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        VDEF$ = OBX$
        GoTo 10
    End If
    '
    If DES% < 1 Or HAS% < 1 Then
       Call COMUNI("Fechas Inexistentes o No Válidas")
       GoTo 99
    End If
    
    
    DESDE = FECHATEX$(DES%)
    HASTA = FECHATEX$(HAS%)
    '
    TXTDESDE = "" ' PARA QUE PROVOQUE EL EVENTO CHANGE DE ACTUALIZAR
    TXTHASTA.TEXT = HASTA
    TXTDESDE.TEXT = DESDE
    '
99  Command3.Enabled = True

End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   NC% = XVALO(Text4)
   If MSF1.Rows <= 1 Then
     Call COMUNI("Falta Detalle de Listado")
     GoTo 99
   End If
   '
   Call GENERATABLAIMPRE
   CONDI$ = "CAMPO1 = '' OR CAMPO1 <> '' OR CAMPO1 IS NULL AND CodiEmpr = " & CodiEmp%
   Call BORRAREGISTROS("TABLA_IMPRE", CONDI$, REGAF&, "NOMESS")
   '
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
25 On Error Resume Next
   SQLX$ = "SELECT *  FROM TABLA_IMPRE  WHERE CAMPO1 =  ''  "
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   
'   Call CARGA_ENCABEZADO(CTACTE.MSF1, "Documento/A12;Fecha/D8;Concepto;Debe/F12.2;Haber/F12.2;Saldo/F12.2;", CTACTE14)
   For I& = 1 To MSF1.Rows - 1
      DOCUMENTO$ = MSF1.TextMatrix(I&, 1)
      FECHA$ = MSF1.TextMatrix(I&, 2)
      CONCEPTO$ = MSF1.TextMatrix(I&, 3)
      CONCEPTO2$ = MSF1.TextMatrix(I&, 4)
      DEBEX$ = MSF1.TextMatrix(I&, 5)
      HABEX$ = MSF1.TextMatrix(I&, 6)
      SALDOX$ = MSF1.TextMatrix(I&, 7)
      With RS1
        .AddNew
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        !CodiEmpr = CodiEmp%
        !CAMPO1 = TRIM$(DOCUMENTO$)
        !CAMPO2 = TRIM$(FECHA$)
        !CAMPO3 = Left$(TRIM$(CONCEPTO$), 48)
        !CAMPO7 = Left$(TRIM$(CONCEPTO2$), 48)
        !CAMPO4 = TRIM$(DEBEX$)
        !CAMPO5 = TRIM$(HABEX$)
        !CAMPO6 = TRIM$(SALDOX$)
        .Update
      End With
   Next I&
   '
'   SALDOANTERIOR$ = "00.00"
'   If CVI(FECHANUM(TXTDESDE)) > 0 And CVI(FECHANUM(TXTHASTA)) > 0 Then
'     SALDOANTERIOR$ = FORMATNUM$(SALDOAFECHA#(NC%, DES%, HAS%), "F12.2")
'   End If
   SALDOACTUAL$ = FORMATNUM$(XVALO(SALDOX$), "F12.2")
   If XVALO(SALDOACTUAL$) = 0 Then SALDOACTUAL$ = "0.00"
   TITULOXX$ = "Cuenta Corriente de Pagos de Impuestos"
   FEDES% = FECHANUM(TRIM$(TXTDESDE))
   FEHAS% = FECHANUM(TRIM$(TXTHASTA))
   FILTX$ = TRIM$(Str$(NC%)) & ";" & TRIM$(SALDOACTUAL$)
    
   Call STDFORM("CTIM-GSM", FILTX$)

99 Command4.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text4.TEXT = TRIM$(Str$(NCLIE%))
    End If
    Command6.Enabled = True

End Sub

Private Sub Form_Load()
'Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
   
    Call APLICACIONSKINS(Me)
    '
    Call CARGA_ENCABEZADO(CCTEANEX14.MSF1, "Documento/A14;Fecha/D8;Concepto 1/A18;Concepto 2/A18;Debe/F12.2;Haber/F12.2;Saldo/F12.2;", CCTEANEX14)
    '
    Call ABRECABAN
End Sub

Private Sub mnuajdev_Click()
    NC% = XVALO(Text4)
    If NC% < 1 Then
      Call COMUNI("Debe Seleccionar un Cliente")
      Exit Sub
    End If
    '
    If DERACCE%("DEVODINE", "DEVOLUCION DE DINERO DE IMPUESTOS") < 2 Then Exit Sub
    '
    AJUSTE14.Label1 = "Retira:"
    AJUSTE14.Caption = "Devoluciòn de Dinero"
    SALDOX# = SALDOAFECHAIMPUESTO#(NC%, , , 1, 0)
    AJUSTE14.Text28 = ""
    AJUSTE14.Text2 = ""
    AJUSTE14.Text1 = FORMATNUM$(SALDOX#, "F12.2")
20  AJUSTE14.Show 1
    '
    'SI SE CERRO EL FRM DE AJUSTES
    If Abs(XVALO(AJUSTE14.Text1)) + Abs(XVALO(AJUSTE14.Text2)) = 0 Then Exit Sub
    '
    SAJUSTADO# = XVALO(AJUSTE14.SALDO_AJUSTADO#)
    If SAJUSTADO# < 0.005 Then
      Call COMUNI("Imposible Realizar Devolución con Importe Negativo")
      Exit Sub
    End If
    '
    Observa$ = AJUSTE14.Text28
    CCONT% = CUECOINT%(Control$("RECISTA", "CUEREDON"))
    If CCONT% < 1 Then
       Call COMUNI("Falta Cuenta Contable de Ajustes")
       Exit Sub
    End If
    If SAJUSTADO# > SALDOX# Then
       Call COMUNI("Imposible Realizar Devolución Mayor al Saldo")
       GoTo 20
    End If
    '
    FEX = HOY(HO$)
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "DVI-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    CONDI$ = " TipoMovim = 'AJUD' AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    ORMA% = XVALO(Mid$(VALOBADA("CAJABANC", "max(CodiCom_Cor)", CONDI$, "NOMESS"), 11, 2))
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    CUECONTABLECLI% = CUECOINT(Control$("AMACLI", "ECUECON"))
    CUECONTABLE% = CUECONTABLECLI%
    Call ABRECABAN
    Dim RS1 As ADODB.Recordset
    Set RS1 = New ADODB.Recordset
25  On Error Resume Next
   
    SQLX$ = "SELECT * FROM CAJABANC  WHERE nuclien < 1  "
    RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
    With RS1
          '
40        .AddNew
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODICOM_COR = Left$(NRECIB$, 12)
          !CODICOM_LAR = Left$(NRECIB$, 24)
          !RefeCom = Left$("Dev Imp- " + RASOCLI$(NC%), 48)
          !DESCRI_IMPU = Left$("Dev Imp.- " + RASOCLI$(NC%), 64)
          FEX1% = FEX
          !FECHCONTA = CVDAT(FEX1%)
          !fechacred = CVDAT(FEX1%)
          !CUECONTAI = CUECONTABLE%
          !NUMEIMPU = 0
          !ItemConc = "" 'nuevo p/ concepto
          !NUCLIEN = NC%
          !NuProve = 0
          !OpciMovim = "CCORR"
          !VARIMOVIM = ""
          !DESCRIMOV = Left$(Observa$, 48)
          !FECHEMISI = CVDAT(FEX1%)
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = ORMA%
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = Abs(IMPOX#)
          '
          If SALIDA% = 0 Then
              !IDEBECOMP = SAJUSTADO#
              !IHABECOMP = 0
              !TipoMovim = "DVIM"
              !ID = 0
          Else
            !IDEBECOMP = 0
            !IHABECOMP = Abs(SAJUSTADO#)
            !TipoMovim = "DVDIFIM"
            !ID = 0
            !Status = 2 ' PARA DIFERENCIAR QUE ES UN AJUSTE X VUELTO DE UN IMPUESTO PAGO
          End If
          !Impo_Dola = 0
          !BASEIMPO = 0
          !NROCAI = 0
          On Error Resume Next
          !IncSubIva = ""
          On Error GoTo 0
         !FEVTOCAI = CVDAT(FEX1%)
         !ObserGen = ""
         !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
         !IdenBaja = String$(6, 0)
         !FeHoReal = Now()
         !NuAsiDiar = ""
         !Niv_Clasi = 0
         !CAE = ""
         !IDCAE = ""
         !IDCAENUM = 0
         !CODCARPIMP = ""
         !NOTAS = ""
         !INCOTERMS = ""
         !CondPag = 0
         !PEDIDOS = ""
         !OCompra = ""
         !Cli_Real = NC%
         !SimMone = "$" ' GRABA EL SIMBOLO DE LA MONEDA
         !AtraCance = 0
         !AtraCheque = 0
         !MARBORRA = 0
         !REFERENCIA_LISTA = ""
         !NOTAS = ""
         !TEXTOCAE = Observa$
         '
         'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
         !Descu1 = ""
         !Descu2 = ""
         !Impodesc1 = 0
         !Impodesc2 = 0
           NUSER1% = USNBR% Mod 100
         !NUMEUSER = NUSER1%
         !NUMLISTA = 0 ' NUMERO DE LISTA
         .Update
        If SALIDA% = 0 Then
          'CARGO INFORMACION PARA QUE GRABE EL MOVIMIENTO POR EL TOTAL
           SALIDA% = 1
           CUECONTABLE% = CCONT%
           GoTo 40
        End If
        '
    End With
    RS1.Close
    Set RS1 = Nothing
    Call COMUNI("Ajuste Realizado")
    CTA = XVALO(Text4)
    Text4 = ""
    Text4 = CTA
End Sub

Private Sub Text4_Change()
    TXTID = Text4
End Sub

Private Sub TXTDESDE_Change()
   MSF1.Rows = 1
   If TRIM$(TXTDESDE) <> "" Then
     Call TXTID_Change
   ElseIf TRIM$(TXTDESDE) = "" And TRIM$(TXTHASTA) <> "" Then
      Call TXTID_Change
   End If
End Sub

Private Sub TXTDESDE_DblClick()
   TXTDESDE = ""
   
End Sub


Private Sub TXTHASTA_DblClick()
    TXTHASTA = ""
End Sub


Private Sub TXTID_Change()
   Call BLANFORMU
   MSF1.Rows = 1
   NC% = XVALO(TXTID)
   If NC% < 1 Then Exit Sub
   '
   Call CTACTE(NC%)
End Sub
Sub CTACTE(Optional NC%)
   TXTNOMBRE = RASOCLI$(NC%)
'   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
'   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
   DES% = FECHANUM(TXTDESDE)
   HAS% = FECHANUM(TXTHASTA)
   MSF1.Rows = 1
      
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE Status >= 2 "
'      SQLX$ = SQLX$ + " or Status = 2) "
      SQLX$ = SQLX$ + " and (tipomovim = 'CHENTR' or TIPOMOVIM = 'CHRCIM' OR TIPOMOVIM = 'DVDIFIM')"
      If NC% > 0 Then SQLX$ = SQLX$ + " And NUCLIEN = " & NC%

      'FILTRO DE FECHAS
      SALFE# = 0

      If DES% > 0 And HAS% > 0 Then
         SQLX$ = SQLX$ + " and FECHEMISI >= '" & FETEXCOR1(DES%) & "' AND FECHEMISI <= '" & FETEXCOR1$(HAS%) & "'"
         SALFE# = CCTEANEX14.SALDOAFECHAIMPUESTO(NC%, DES%, HAS%, 1, 0)
         I& = MSF1.Rows
        
         MSF1.Rows = MSF1.Rows + 1
            MSF1.TextMatrix(I&, 1) = "Transporte"
            MSF1.TextMatrix(I&, 2) = TXTDESDE
            MSF1.TextMatrix(I&, 3) = "Saldo Acumulado"
            HABEX1$ = "0"
            DEBEX1$ = "0"
            If SALFE# > 0 Then
               DEBEX1$ = FORMATNUM$(SALFE#, "F12.2")
            Else
               HABEX1$ = FORMATNUM$(SALFE#, "F12.2")
            End If
            MSF1.TextMatrix(I&, 5) = DEBEX1$
            MSF1.TextMatrix(I&, 6) = HABEX1$
            ACUMUMUSALDO# = Abs(XVALO(DEBEX1$) - XVALO(HABEX1$))
            MSF1.TextMatrix(I&, 7) = FORMATNUM$(ACUMUMUSALDO#, "F12.2")
      End If
      NCPARAM% = NC% 'GUARDO EL NUMERO DE CLIENTE PARA VALIDAR EN EL DO WHILE

      If NC% > 0 Then
         SQLX$ = SQLX$ + " ORDER BY FECHEMISI"
      Else
         SQLX$ = SQLX$ + " ORDER BY NUCLIEN"
      End If
      Set RS = READSET(SQLX$)
      VUELTA% = 0
      With RS
        If Not .EOF Then
          Do While Not .EOF
            STAT% = XVALO(!Status)
            TMOV$ = SAFETEXT$(!TipoMovim)
            If TMOV$ = "CHRCIM" And STAT% = 0 Then GoTo 75 ' SALTEA LOS IMPUESTOS COBRADOS QUE NO SE GENERO O.PAGO
            '
            NC% = XVALO(!NUCLIEN)
            '
            If NCPARAM% > 0 Or NC% <> NC_ANTERIOR% Or VUELTA% = 0 Then 'ESTE IF VALIDA QUE SI SELECCIONO CLIENTE INGRESA SIEMPRE
               I& = MSF1.Rows                           'SI NO TIENE NUMERO DE CLIENTE SOLO VA A ENTRAR CUANDO
               MSF1.Rows = I& + 1                       'CAMBIE DE NUMERO DE CLIENTE
               ACUMUMUDEBEX# = 0
               ACUMUMUHABEX# = 0
               If NCPARAM% < 1 Then ACUMUMUSALDO# = 0
            End If
            '
            If NCPARAM% > 0 Then 'SI TIENE NC IMPRIME LOS DATOS
              NURECI$ = SAFETEXT$(!CODICOM_COR)
              CONCEPTO$ = LISTAP14.DESCRIIMPU$(XVALO(!NUMEIMPU))
              DESCRIMOV$ = SAFETEXT$(!DESCRIMOV)
              FE% = CVINT(!FECHEMISI)
            End If
            DEBEX1$ = FORMATNUM$(XVALO(!IDEBECOMP), "F12.2")
            HABEX1$ = FORMATNUM$(XVALO(!IHABECOMP), "F12.2")
            SALDO1# = XVALO(!IDEBECOMP) - XVALO(!IHABECOMP)
            ACUMUMUDEBEX# = ACUMUMUDEBEX# + XVALO(DEBEX1$)
            ACUMUMUHABEX# = ACUMUMUHABEX# + XVALO(HABEX1$)
            ACUMUMUSALDO# = ACUMUMUSALDO# + SALDO1#
            NC_ANTERIOR% = NC%
            '
            MSF1.TextMatrix(I&, 1) = NURECI$
            MSF1.TextMatrix(I&, 2) = FECHATEX$(FE%)
            MSF1.TextMatrix(I&, 3) = CONCEPTO$
            MSF1.TextMatrix(I&, 4) = DESCRIMOV$
            If NCPARAM% > 0 Then 'SI TIENE NC IMPRIME LOS DATOS
              MSF1.TextMatrix(I&, 5) = DEBEX1$
              MSF1.TextMatrix(I&, 6) = HABEX1$
            Else
              MSF1.TextMatrix(I&, 1) = NC%
              MSF1.TextMatrix(I&, 2) = RASOCLI$(NC%)
              MSF1.TextMatrix(I&, 5) = ACUMUMUDEBEX#
              MSF1.TextMatrix(I&, 6) = ACUMUMUHABEX#
            End If
            MSF1.TextMatrix(I&, 7) = FORMATNUM$(ACUMUMUSALDO#, "F12.2")
            VUELTA% = 1
75          .MoveNext
          Loop
      End If
      End With
      RS.Close
      Set RS = Nothing
      '
      MSF1.COL = 2
      MSF1.Sort = 1
      Call COLOREAR_GRILLASQL(MSF1, &HE0E0E0)
   End Sub

Function SALDOAFECHA#(NC%, Optional DES%, Optional HAS%)
   'CONSULTA EL SALDO DEL CLIENTE
   'SI TIENE RANGO DE FECHAS FILTRA LA SUMA DEL HABER Y DEL DEBE SEGUN FECHA
   'SI NO TIENE FECHA NO CONSIDERA RANGO DE FECHAS
   
''''   CONDI$ = "NUCLIEN = " & NC% & " AND TIPOMOVIM = 'FHON' AND STATUS < 9 "
''''   If DES% > 0 And HAS% > 0 Then
''''      DESDE = FETEXCOR1$(DES%): HASTA = FETEXCOR1$(HAS%)
''''      CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDE & "' AND FECHEMISI <= '" & HASTA & "'"
''''   End If
''''   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
''''   CONDI$ = "NUCLIEN = " & NC% & " AND TIPOMOVIM = 'RHON' AND STATUS < 9 "
''''   If DES% > 0 And HAS% > 0 Then
''''      CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DESDE & "' AND FECHEMISI <= '" & HASTA & "'"
''''   End If
''''   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
''''
''''   SALDOX# = IDEBEX# - IHABEX#
''''   SALDOAFECHA# = SALDOX#
   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'FHON' OR TIPOMOVIM = 'AJUD' OR TIPOMOVIM = 'RHON') AND STATUS < 9"
   If DES% > 0 And HAS% > 0 Then
      DESDE = FETEXCOR1$(DES%): HASTA = FETEXCOR1$(HAS%)
      CONDI$ = CONDI$ + " AND FECHEMISI < '" & DESDE & "' " 'AND FECHEMISI <= '" & HASTA & "'"
   End If
   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
   SALDOX# = IDEBEX# - IHABEX#
   SALDOAFECHA# = SALDOX#
   
End Function

Function SALDOAFECHAIMPUESTO#(NC%, Optional DES%, Optional HAS%, Optional SOLO_IMP_PAGOS%, Optional INCLUIR_AJUSTE_CTACTE_GENERAL%)
   'CONSULTA EL SALDO DEL CLIENTE
   'SI TIENE RANGO DE FECHAS FILTRA LA SUMA DEL HABER Y DEL DEBE SEGUN FECHA
   'SI NO TIENE FECHA NO CONSIDERA RANGO DE FECHAS
   'SI SOLO_IMP_PAGOS% > 0 TRAE SOLO LOS IMPUESTOS QUE SE LE HAYAN REALIZADO LA ORDEN DE PAGO.
   '
   CONDI$ = "(TIPOMOVIM = 'CHENTR' OR TIPOMOVIM = 'DVDIFIM' OR TIPOMOVIM = 'CHRCIM'"
   If INCLUIR_AJUSTE_CTACTE_GENERAL% > 0 Then
   CONDI$ = CONDI$ + " OR TIPOMOVIM = 'AJIM' "
   End If
   CONDI$ = CONDI$ + ")"
   If NC% > 0 Then CONDI$ = CONDI$ + " AND NUCLIEN = " & NC%
   CONDI$ = CONDI$ + " AND STATUS < 9"
   
   If DES% > 0 And HAS% > 0 Then
      DESDE = FETEXCOR1$(DES%): HASTA = FETEXCOR1$(HAS%)
      CONDI$ = CONDI$ + " AND FECHEMISI < '" & DESDE & "' " 'AND FECHEMISI <= '" & HASTA & "'"
   End If
   COBROS_IMPU_NO_PAGADOS# = 0
   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
   SALDOX# = IDEBEX# - IHABEX#
   SALDOAFECHAIMPUESTO# = SALDOX#
   '
   If SOLO_IMP_PAGOS% > 0 Then
      CONDI2$ = "TIPOMOVIM = 'CHRCIM'  AND STATUS < 1  AND NUCLIEN = " & NC%
      COBROS_IMPU_NO_PAGADOS# = XVALO(VALOBADA("CAJABANC", "Sum(IDEBECOMP)", CONDI2$, "NOMESS"))
      SALDOAFECHAIMPUESTO# = SALDOAFECHAIMPUESTO# - COBROS_IMPU_NO_PAGADOS#
   End If
   '
End Function

Sub BLANFORMU()
   CCTEANEX14.TXTNOMBRE = ""
End Sub

