VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CARHON14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga de Honorarios"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5080
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Saldo:"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TXTSALDO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Doble Click Para Visualizar Cta.Cte."
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Carga de Honorarios"
      Top             =   120
      Width           =   4920
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nuevo Honorario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   7695
      Begin VB.TextBox TXTCONCEPTO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1680
         Width           =   7185
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
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Usuario"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox TXTUSUARIO 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2880
         Width           =   3495
      End
      Begin VB.TextBox Text2 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Id.Honorario"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TXTNURECIBO 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
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
         Left            =   4920
         TabIndex        =   17
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
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
         Left            =   6000
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   195
      End
      Begin VB.TextBox TXTFECHA 
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
         Left            =   6200
         TabIndex        =   15
         Top             =   840
         Width           =   1240
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
         Left            =   3630
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   195
      End
      Begin VB.TextBox TEXT17 
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
         Index           =   3
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   840
         Width           =   1065
      End
      Begin VB.TextBox Text10 
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Importe"
         Top             =   1320
         Width           =   1065
      End
      Begin VB.TextBox TXTIMPORTE 
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
         Left            =   5966
         TabIndex        =   1
         Top             =   1320
         Width           =   1475
      End
      Begin VB.TextBox Text8 
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Concepto"
         Top             =   1320
         Width           =   1095
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Nombre"
         Top             =   840
         Width           =   1095
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Cta.:"
         Top             =   360
         Width           =   615
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
         Left            =   3840
         MaxLength       =   24
         TabIndex        =   0
         Top             =   360
         Width           =   975
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   3495
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "ID"
         Top             =   360
         Width           =   1095
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   24
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "CARHON14.frx":0000
         TabIndex        =   14
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
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "CARHON14.frx":005E
      Top             =   120
   End
   Begin VB.Menu MNUPH 
      Caption         =   "PAGOS DE HONORARIOS"
      Begin VB.Menu mnuregph 
         Caption         =   "Registrar pago de honorarios"
      End
   End
   Begin VB.Menu mnuaj 
      Caption         =   "Ajustes"
      Begin VB.Menu mnuajctaclte 
         Caption         =   "Realizar Ajustes de Cuentas de Clientes"
      End
   End
   Begin VB.Menu mnuccte 
      Caption         =   "Cta. Corriente"
      Begin VB.Menu mnucctecliente 
         Caption         =   "Cta.Corriente x Cliente"
      End
   End
   Begin VB.Menu mnr 
      Caption         =   "Reportes"
      Begin VB.Menu mnusp 
         Caption         =   "Saldos Pendientes de Cobranza"
      End
   End
End
Attribute VB_Name = "CARHON14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   NC% = XVALO(TXTID)
   '
   If XVALO(TXTID) < 1 Then
      Call COMUNI("Identificador No Válido o Inexistente")
      GoTo 99
   End If
   '
   If TRIM$(TXTNOMBRE) = "" Then
      Call COMUNI("Cliente No Válido o Inexistente")
      GoTo 99
   End If
   '
   If TRIM$(TXTCONCEPTO) = "" Then
      Call COMUNI("Falta Concepto ")
      GoTo 99
   End If
  
   If XVALO(TXTIMPORTE) < 1 Then
      Call COMUNI("Importe No Válido o Inexistente")
      GoTo 99
   End If
   '
   FECH1% = FECHANUM(TXTFECHA)
   If FECH1% < 1 Then
     Call COMUNI("Fecha Inválida o Inexistente")
     GoTo 99
   End If
   '
   CUECONTABLECLI% = CUECOINT(Control$("AMACLI", "ECUECON"))
   If CUECONTABLECLI% < 1 Then
      Call COMUNI("Falta Cuenta Madre de Clientes")
      GoTo 99
   End If
   '
   CUECONTABLECAJAHONO% = CUECOINT(XVALO(Control$("SVD451S", "CUEVE-0")))  'CUENTA DE HONORARIOS
   If CUECONTABLECAJAHONO% < 1 Then
      Call COMUNI("Falta Cuenta Contable de Caja")
      GoTo 99
   End If
   '
   If XVALO(TXTNURECIBO) <> PRONUHON& Then
      Call COMUNI("Se va a Re-Numerar Identificador de Recibo\El Mismo Esta Ocupado o No es Válido")
      TXTNURECIBO = PRONUHON&
      GoTo 99
   End If
   '
   COCOR$ = "00000000" & TRIM$(TXTNURECIBO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "FH-" & COCOR$
   COLAR$ = COCOR$

   FECH = FECHATEX$(FECH1%)
   'ASIGNA PRIMERO LA CUENTA CONTABLE DE CAJA LUEGO CUANDO GRABA EL MOVIMIENTO POR EL IMPORTE TOTAL SE CAMBIA A LA DE D.X.VENTA
   CUECONTABLE% = CUECONTABLECLI%
   '
   Call ABRECABAN

   CONDI$ = "Nume_Cli = " & NC%
   If COMALTER%("Atención !! Se va a Dar de Alta el Honorario Activo\\Continuar\Cancelar") = 1 Then
        On Error GoTo ERROR_ALTA
        Dim RS1 As ADODB.Recordset
        Set RS1 = New ADODB.Recordset
25      On Error Resume Next
        SQLX$ = "SELECT * FROM CAJABANC  WHERE NUCLIEN =  1 "
        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
        FLEXCONN.BeginTrans
        With RS1
   
40        .AddNew
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !CODICOM_COR = Left$(COCOR$, 12)
          !CODICOM_LAR = Left$(COLAR$, 24)
          !RefeCom = Left$("Honorarios de : " & RASOCLI$(NC%), 64)
          !DESCRI_IMPU = ""
          !FECHCONTA = SafeDate(FECH)
          !fechacred = SafeDate(FECH)
          !CUECONTAI = CUECONTABLE%
          !NUMEIMPU = 0
          !ItemConc = "" 'nuevo p/ concepto
          !NUCLIEN = NC%
          !NuProve = 0
          !OpciMovim = ""
          !VARIMOVIM = ""
          !TEXTOCAE = TXTCONCEPTO
          !FECHEMISI = SafeDate(FECH)
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = XVALO(TXTNURECIBO)
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = XVALO(TXTIMPORTE)
          '
          If VUELTA% = 0 Then
            !IDEBECOMP = XVALO(TXTIMPORTE)
            !IHABECOMP = 0
            !TIPOMOVIM = "FHON"
            !ID = 0
          Else
            !IDEBECOMP = 0
            !IHABECOMP = XVALO(TXTIMPORTE)
            !TIPOMOVIM = "HONO"
            !ID = 0
          End If
          !Impo_Dola = 0
          !BASEIMPO = 0
          !NROCAI = 0
          On Error Resume Next
          !IncSubIva = ""
          On Error GoTo 0
         !FEVTOCAI = SafeDate(FECH)
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
         !Pedidos = ""
         !OCompra = ""
         !Cli_Real = NC%
         !SimMone = "$" ' GRABA EL SIMBOLO DE LA MONEDA
         !AtraCance = 0
         !AtraCheque = 0
         !MARBORRA = 0
         !REFERENCIA_LISTA = ""
         !NOTAS = ""
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
         If VUELTA% = 0 Then
          'CARGO INFORMACION PARA QUE GRABE EL MOVIMIENTO POR EL TOTAL
           VUELTA% = 1
           CUECONTABLE% = CUECONTABLECAJAHONO%
           GoTo 40
         End If
         '
        End With
ERROR_ALTA:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL DAR DE ALTA,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación  de Generación de Honorario: " & TXTNURECIBO & " Realizada Exitosamente")
      CARHON14.TXTCONCEPTO = ""
      CARHON14.TXTFECHA = ""
      CARHON14.TXTID = ""
      CARHON14.TXTIMPORTE = ""
      CARHON14.TXTNOMBRE = ""
      CARHON14.TXTNURECIBO = ""
      CARHON14.Text4 = ""
    End If
        RS1.Close
        Set RS1 = Nothing

   End If
   TXTNURECIBO = PRONUHON&
99 Command1.Enabled = True

End Sub

Private Sub Command3_Click()
   Command6.Enabled = False
   Call SELECHO("SELFECHA")
   TXTFECHA.TEXT = VALACT1$("SELFECHA")
99 Command6.Enabled = True

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
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
''    Call ABRECONEXION
''    'VALIDACION PARA AMBEST POR SI NO PAGA
''    'If InStr(EPAC$, "AMBEST") > 0 Then
''         FF% = FECHANUM("20/06/14")
''         If FF% <= HOY(HO$) Then
''           Call Form_Load
''         End If
''    'End If

    UTILIZA_SKIN% = 1
    TXTFECHA = FECHATEX$(HOY(HO$))
    TXTNURECIBO = PRONUHON&
    TXTUSUARIO = USERNAME$

    
    Call APLICACIONSKINS(Me)
    '
    Call GENERATABLAIMPRE

End Sub
Function PRONUHON&()
    CONDI$ = "TipoMovim = 'FHON' "
    CONDI$ = CONDI$ + " AND CodiCom_Cor LIKE 'FH%' "
    PRONUHON& = 1 + XVALO(VALOBADA("CAJABANC", "MAX(NUMECOMP)", CONDI$, "NOMESS"))
End Function

Private Sub TXTIVA_Change()
   Text2 = ECOARCH$("CAPOSIV", Chr(XVALO(TXTIVA)))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Close
   End
End Sub

Private Sub mnuajctaclte_Click()

  
    Call SELECHO("DEUDOR1")
    NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NC% < 1 Then Exit Sub

''    SQLX$ = "SELECT SUM(IDEBECOMP) AS DEBEX , SUM(IHABECOMP) AS HABEX  FROM CAJABANC WITH(NOLOCK)"
''    SQLX$ = SQLX$ + " where NuClien = " & NC%
''    SQLX$ = SQLX$ + " AND (TipoMovim = 'FHON' "
''    SQLX$ = SQLX$ + " or TipoMovim = 'RHON' "
''    SQLX$ = SQLX$ + " or TipoMovim = 'AJUD') "
''    SQLX$ = SQLX$ + " GROUP BY NUCLIEN "
''    Set RS = READSET(SQLX$)
''    With RS
''     SALDOX# = XVALO(!DEBEX) - XVALO(!HABEX)
''    End With
    '
    SALDOX# = CTACTE14.SALDOAFECHA#(NC%)
    AJUSTE14.Text28 = ""
    AJUSTE14.Text2 = ""
    AJUSTE14.Text1 = FORMATNUM$(SALDOX#, "F12.2")
    AJUSTE14.Show 1
    '
    'SI SE CERRO EL FRM DE AJUSTES
    If Abs(XVALO(AJUSTE14.Text1)) + Abs(XVALO(AJUSTE14.Text2)) = 0 Then Exit Sub
    '
    SAJUSTADO# = XVALO(AJUSTE14.SALDO_AJUSTADO#)
    Observa$ = AJUSTE14.Text28
    CCONT% = CUECOINT%(Control$("RECISTA", "CUEREDON"))
    If CCONT% < 1 Then
       Call COMUNI("Falta Cuenta Contable de Ajustes")
       Exit Sub
    End If
    DIFERENCIA# = SAJUSTADO# - SALDOX#
    If DIFERENCIA# = 0 Then
       Call COMUNI("Saldo Ajustado Igual a Saldo ")
       Exit Sub
    End If
    '
    FEX = HOY(HO$)
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
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
          !RefeCom = "Aj.Saldo - " + RASOCLI$(NC%)
          !DESCRI_IMPU = Left$("Aj.Saldo - " + RASOCLI$(NC%), 64)
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
          !DESCRIMOV = Left$("Aj.Saldo - " + RASOCLI$(NC%), 48)
          !FECHEMISI = CVDAT(FEX1%)
          !CodiEmpr = CodiEmp%
          !CodiSucu = 0
          !NumeTalo = 0
          !NUMECOMP = XVALO(Mid$(NRECIB$, 4, 6))
          !Status = 0
          On Error Resume Next
          !Mone_Emis = 1
          !Tipo_Cam = 1
          !ValAbsComp = Abs(IMPOX#)
          '
          If SALIDA% = 0 Then
            If DIFERENCIA# > 0 Then
              !IDEBECOMP = Abs(DIFERENCIA#)
              !IHABECOMP = 0
            Else
              !IHABECOMP = Abs(DIFERENCIA#)
              !IDEBECOMP = 0
            End If
            !TIPOMOVIM = "AJUD"
            !ID = 0

          Else
           If DIFERENCIA# > 0 Then
            !IDEBECOMP = 0
            !IHABECOMP = Abs(DIFERENCIA#)
           Else
            !IDEBECOMP = Abs(DIFERENCIA#)
            !IHABECOMP = 0
           End If
            !TIPOMOVIM = "AJUSALD"
            !ID = 0
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
         !Pedidos = ""
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

End Sub

Private Sub mnucctecliente_Click()
   CTACTE14.Show 1
End Sub

Private Sub mnuregph_Click()
     CAPAHO14.Show 1
End Sub

Private Sub mnusp_Click()

   '
   Dim IMPORDEBE()
   Dim IMPORHABE()
   Dim NUMECLIENTE()
   '
   Call GENERATABLAIMPRE
   CONDI$ = "CAMPO1 = '' OR CAMPO1 <> '' OR CAMPO1 IS NULL AND CodiEmpr = " & CodiEmp%
   Call BORRAREGISTROS("TABLA_IMPRE", CONDI$, REGAF&, "NOMESS")
   '
   '
   CONDI$ = "  (Status = 0 OR STATUS IS NULL) AND (TIPOMOVIM = 'FHON' OR TIPOMOVIM = 'AJUD' OR TIPOMOVIM = 'RHON') "
   '
   IJ& = 0
   SQLX$ = "SELECT SUM(IDEBECOMP)AS IMDEB , SUM(IHABECOMP)AS IMHAB ,NUCLIEN FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " AND NUCLIEN > 0 "
   SQLX$ = SQLX$ + " GROUP BY NUCLIEN"
   SQLX$ = SQLX$ + " ORDER BY NUCLIEN"
   Set RS = READSET(SQLX$)
   With RS
      Do While Not .EOF
        IJ& = IJ& + 1
        ReDim Preserve IMPORDEBE(IJ&), IMPORHABE(IJ&), NUMECLIENTE(IJ&)
        IMPORDEBE(IJ&) = XVALO(!IMDEB)
        IMPORHABE(IJ&) = XVALO(!IMHAB)
        NUMECLIENTE(IJ&) = XVALO(!NUCLIEN)
        'RSOCIAL$(IJ&) = RASOCLI$(NUMECLIENTE(IJ&))
        .MoveNext
      Loop
   End With
   RS.Close
   Set RS = Nothing
   '
   Dim RS1 As ADODB.Recordset
   Set RS1 = New ADODB.Recordset
   SQLX$ = "SELECT *  FROM TABLA_IMPRE  WHERE CAMPO1 =  ''  "
   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   '
   ACUMUSALDO = 0
   If IJ& < 1 Then Exit Sub
   For I& = 0 To UBound(NUMECLIENTE)
      With RS1
        SALDO = IMPORDEBE(I&) - IMPORHABE(I&)
        If SALDO > 0.005 Then
            ACUMUSALDO = ACUMUSALDO + SALDO
            .AddNew
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            !CodiEmpr = CodiEmp%
            !CAMPO1 = NUMECLIENTE(I&)
            !CAMPO2 = RASOCLI$(NUMECLIENTE(I&))
            !CAMPO3 = FORMATNUM$(SALDO, "F12.2")
            .Update
        End If
      End With
   Next I&
   '
   FILTX$ = TRIM$(Str$(ACUMUSALDO))
    
   Call STDFORM("SDEU-GSM", FILTX$)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub TXTFECHA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TXTID_Change()
   Call BLANFORMU
   NC% = XVALO(TXTID)
   If NC% < 1 Then Exit Sub
   '
   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
   
      SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE NUME_CLI = " & NC% & " And Stat_Cli >= 0"
      Set RS = READSET(SQLX$)
      With RS
        If Not .EOF Then
            TXTNOMBRE = SAFETEXT$(!Raso_Cli)
        End If
      End With
   End If
   Call VERISALDO(NC%, CARHON14.TXTSALDO)
   
End Sub

Private Sub Text4_Change()
   TXTID = Text4
   
End Sub
Sub BLANFORMU()
        TXTNOMBRE = ""
        TXTDOMICILIO = "" '64
        TXTLOCALIDAD = "" '48
        TXTTELEFONO1 = ""  '48
        TXTTELEFONO2 = ""  '32
        TXTEMAIL1 = ""     '64
        TXTEMAIL2 = ""     '128
        TXTCODIGOPOSTAL = "" '16
        TXTCONTACTO = ""    '48
        TXTCUIT = ""   '24
        TXTIVA = ""
End Sub

Private Sub TXTLOCALIDAD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTIMPORTE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTIMPORTE_LostFocus()
   TXTIMPORTE = FORMATNUM$(XVALO(TXTIMPORTE), "F12.2")

End Sub

Private Sub TXTNOMBRE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub TXTSALDO_DblClick()
  CTACTE14.TXTID = XVALO(TXTID)
  CTACTE14.Show 1

End Sub
