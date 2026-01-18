VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CAPAHO14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Honorarios"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "Pago de Honorarios"
      Top             =   240
      Width           =   7680
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
      Top             =   600
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
         Height          =   780
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   7185
      End
      Begin VB.TextBox TXTSALDO 
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
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Visualizar Cta.Cte."
         Top             =   1560
         Width           =   1420
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
         Index           =   2
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Saldo"
         Top             =   1560
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
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Usuario"
         Top             =   1150
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
         Top             =   1150
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
         Text            =   "N° Recibo"
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
         Caption         =   "Aceptar"
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
         Left            =   4920
         TabIndex        =   17
         Top             =   2800
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "!"
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
         Left            =   6000
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   760
         Width           =   195
      End
      Begin VB.TextBox TXTFECHA 
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
         Left            =   6200
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   760
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
         Top             =   760
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
         Top             =   1150
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
         Top             =   1150
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
         Top             =   1560
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
         Top             =   760
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
         Text            =   "Cta."
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
         Top             =   760
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
         OleObjectBlob   =   "CAPAHO14.frx":0000
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
      OleObjectBlob   =   "CAPAHO14.frx":005E
      Top             =   120
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu mnureci 
         Caption         =   "Ver Recibos de Honorios"
      End
   End
   Begin VB.Menu mnurecibo 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnumantreci 
         Caption         =   "Mantenimiento de Recibos"
      End
   End
   Begin VB.Menu mnuctacte 
      Caption         =   "Cuenta Corriente"
      Begin VB.Menu mnuctacorriente 
         Caption         =   "Cta.Corriente"
      End
   End
End
Attribute VB_Name = "CAPAHO14"
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
   CUECONTABLECAJA% = CUECOINT(XVALO(Control$("", "CUENCAJA")))
   If CUECONTABLECAJA% < 1 Then
      Call COMUNI("Falta Cuenta Contable de Caja")
      GoTo 99
   End If
   '
   If XVALO(TXTNURECIBO) <> PRONURECIHON& Then
      Call COMUNI("Se va a Re-Numerar Identificador de Recibo\El Mismo Esta Ocupado o No es Válido")
      TXTNURECIBO = PRONURECIHON&
      GoTo 99
   End If
   '
   COCOR$ = "00000000" & TRIM$(TXTNURECIBO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "RH-" & COCOR$
   COLAR$ = COCOR$

   FECH = FECHATEX$(FECH1%)
   'ASIGNA PRIMERO LA CUENTA CONTABLE DE CAJA LUEGO CUANDO GRABA EL MOVIMIENTO POR EL IMPORTE TOTAL SE CAMBIA A LA DE D.X.VENTA
   CUECONTABLE% = CUECONTABLECLI%
   '
   Call ABRECABAN

   On Error GoTo ERROR_ALTA
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
          !RefeCom = "Carga de Honorarios de : " & RASOCLI$(NC%)
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
          !TEXTOCAE = TXTCONCEPTO 'CAMBIE POR QUE NECESITO UN CAMPO TIPO MEMO
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
            !IDEBECOMP = 0
            !IHABECOMP = XVALO(TXTIMPORTE)
            !TIPOMOVIM = "RHON"
            !ID = 0
          Else
            !IDEBECOMP = XVALO(TXTIMPORTE)
            !IHABECOMP = 0
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
           CUECONTABLE% = CUECONTABLECAJA%
           GoTo 40
         End If
         '
         .Update
        End With
ERROR_ALTA:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL DAR DE ALTA,  NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación de Recibo de Honorario: " & TXTNURECIBO & " Realizada Exitosamente")
    End If
        RS1.Close
        Set RS1 = Nothing
   End If
   '
   USUARIONOMBRE$ = USERNAM$(NUSER1%)
   FILTX$ = COCOR$ & ";" & TRIM$(Str$(NC%)) & ";" & USUARIONOMBRE$ & ";" & TXTCONCEPTO
   '
   Call STDFORM("RCOH-GSM", FILTX$)
   TXTCONCEPTO = ""
   TXTFECHA = ""
   TXTID = ""
   TXTIMPORTE = ""
   TXTNOMBRE = ""
   TXTNURECIBO = ""
   Text4 = ""
   TXTSALDO = ""
   '
   TXTNURECIBO = PRONURECIHON&
   '
   
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
    UTILIZA_SKIN% = 1
    TXTFECHA = FECHATEX$(HOY(HO$))
    TXTNURECIBO = PRONURECIHON&
    TXTUSUARIO = USERNAME$

    
    Call APLICACIONSKINS(Me)
    '
End Sub
Function PRONURECIHON&()
    CONDI$ = "TipoMovim = 'RHON' "
    CONDI$ = CONDI$ + " AND CodiCom_Cor LIKE 'RH%' "
    PRONURECIHON& = 1 + XVALO(VALOBADA("CAJABANC", "MAX(NUMECOMP)", CONDI$, "NOMESS"))
End Function

Private Sub TXTIVA_Change()
   Text2 = ECOARCH$("CAPOSIV", Chr(XVALO(TXTIVA)))
End Sub

Private Sub mnuajctaclte_Click()
    
End Sub

Private Sub mnuctacorriente_Click()
   CTACTE14.Show 1
End Sub

Private Sub mnumantreci_Click()
   MANREC14.Show 1
End Sub

Private Sub mnuReci_Click()
    CONDI$ = "TIPOMOVIM= 'RHON' AND STATUS = 0 ORDER BY CODICOM_COR"
    Call CARGALISEL("LISRECIB", "CAJABANC", "CODICOM_COR/A12;NUCLIEN/I5+DEUDOR12/A32;FECHEMISI/D8;IHABECOMP/F12.2", CONDI$, "NOMESS")
    Call SELECHO("LISRECIB")
    NURECIBO$ = TRIM$(VALACT1$("LISRECIB"))
    If NURECIBO$ = "" Then Exit Sub
    NC% = XVALO(TRIM$(Left$(VALACT2$("LISRECIB"), 5)))
   '
    CONDI$ = "CODICOM_COR = '" & NURECIBO$ & "' AND TIPOMOVIM = 'RHON' AND NUCLIEN = " & NC%
    CONCEPTO$ = VALOBADA("CAJABANC", "TEXTOCAE", CONDI$, "NOMESS")
    NUSER1% = XVALO(VALOBADA("CAJABANC", "NUMEUSER", CONDI$))
    USUARIONOMBRE$ = USERNAM$(NUSER1%)
    FILTX$ = NURECIBO$ & ";" & TRIM$(Str$(NC%)) & ";" & USUARIONOMBRE$ & ";" & CONCEPTO$
    Call STDFORM("RCOH-GSM", FILTX$)
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
      RS.Close
      Set RS = Nothing
      '
      Call VERISALDO(NC%, CAPAHO14.TXTSALDO)
      
   End If
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

Private Sub TXTIMPORTE_GotFocus()
   TXTIMPORTE = FORMATNUM$(XVALO(TXTIMPORTE), "F12.2")
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
CTACTE14.Text4 = XVALO(Text4)
CTACTE14.Show 1
End Sub


