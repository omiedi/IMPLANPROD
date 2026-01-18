VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form MANREC14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Recibos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
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
      Text            =   "Mantenimiento de Recibos"
      Top             =   120
      Width           =   5400
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
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   5415
      Begin VB.TextBox txtusuariorealizo 
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Us.Realizó"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Eliminar"
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
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox TXTHABER 
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
         Left            =   1290
         TabIndex        =   20
         Top             =   2640
         Width           =   1230
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Haber"
         Top             =   2640
         Width           =   1065
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
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Us. Actual"
         Top             =   4080
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4080
         Width           =   3720
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
         Left            =   3100
         TabIndex        =   16
         Top             =   3480
         Width           =   1935
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   1320
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
         Left            =   3480
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   2040
         Width           =   1095
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Debe"
         Top             =   2040
         Width           =   1065
      End
      Begin VB.TextBox TXTDEBE 
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
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2040
         Width           =   1230
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
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TXTCONCEPTO 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   1440
         Width           =   3735
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Recibo"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TXTNURECIBO 
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
         Left            =   3720
         MaxLength       =   24
         TabIndex        =   0
         Top             =   360
         Width           =   1335
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
         Width           =   3735
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
         Text            =   "Cta Cliente"
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
         OleObjectBlob   =   "MANREC14.frx":0000
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
      OleObjectBlob   =   "MANREC14.frx":005E
      Top             =   120
   End
   Begin VB.Menu mnurev 
      Caption         =   "Revertir"
      Begin VB.Menu mnurevertir_rec 
         Caption         =   "Revertir Recibos Anulados"
      End
   End
End
Attribute VB_Name = "MANREC14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   NC% = XVALO(TXTID)
   '
   NURECI$ = TRIM$(TXTNURECIBO)
   If NURECI$ = "" Then
      Call COMUNI("Falta Seleccionar Recibo")
      GoTo 99
   End If

   If NC% < 1 Then
      Call COMUNI("Número de Cliente No Válido o Inexistente")
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
  
   If XVALO(TXTDEBE) + XVALO(TXTHABER) < 0.005 Then
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
   TX$ = "" & NURECI$ & vbCrLf
   TX$ = TX$ + "Concepto: " & TXTCONCEPTO
   If XVALO(TXTDEBE) > 0 Then TX$ = TX$ + "Debe: " & FORMATNUM$(XVALO(TXTDEBE), "F12.2")
   If XVALO(TXTHABER) > 0 Then TX$ = TX$ + "Haber: " & FORMATNUM$(XVALO(TXTHABER), "F12.2")
   
   If COMALTER%("Confirmación de Modificación\\Cancelar\Confirma Modificar Recibo: " & TX$ & "") = 2 Then
        'ACTUALIZO EL CONCEPTO EN LOS DOS MOVIMIENTOS DEL RECIBO
        CONDI$ = "CODICOM_LAR = '" & NURECI$ & "' AND (TIPOMOVIM= 'RHON' OR TIPOMOVIM= 'HON0'"
        Call ACTUREGISTRO("CAJABANC", "DESCRIMOV", CONDI$, TXTCONCEPTO, REGAF&, "NOMESS")
        'ACTUALIZO EL IMPORTE DE HABER
        CONDI$ = "CODICOM_LAR = '" & NURECI$ & "' AND TIPOMOVIM= 'RHON' "
        Call ACTUREGISTRO("CAJABANC", "IHABECOMP", CONDI$, FORMATNUM$(XVALO(TXTHABER), "F16.5"), REGAF&, "NOMESS")
        'ACTUALIZO EL IMPORTE DEL DEBE
        CONDI$ = "CODICOM_LAR = '" & NURECI$ & "' AND TIPOMOVIM= 'HON0' "
        Call ACTUREGISTRO("CAJABANC", "IDEBECOMP", CONDI$, FORMATNUM$(XVALO(TXTDEBE), "F16.5"), REGAF&, "NOMESS")
   End If
   '
   Call BLANFORMU
   
99 Command2.Enabled = True

End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    '
    NURECI$ = TRIM$(TXTNURECIBO)
    TX$ = "" & NURECI$
'    TX$ = TX$ + "Concepto: " & TXTCONCEPTO
'    If XVALO(TXTDEBE) > 0 Then TX$ = TX$ + "Debe: " & FORMATNUM$(XVALO(TXTDEBE), "F12.2")
'    If XVALO(TXTHABER) > 0 Then TX$ = TX$ + "Haber: " & FORMATNUM$(XVALO(TXTHABER), "F12.2")
   
    If COMALTER%("Confirmación de Anulación\\Cancelar\   Confirma Anular : " & TX$ & "") = 2 Then
        CONDI$ = "CODICOM_LAR = '" & NURECI$ & "' AND (TIPOMOVIM= 'RHON' OR TIPOMOVIM= 'HONO')"
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
        Call COMUNI("Anulación Completada")
    End If
    Call BLANFORMU
    
99  Command2.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    CONDI$ = "TIPOMOVIM= 'RHON' AND STATUS < 9 ORDER BY CODICOM_COR"
    Call CARGALISEL("LISRECIB", "CAJABANC", "CODICOM_COR/A12;NUCLIEN/I5+DEUDOR12/A32;FECHEMISI/D8;IHABECOMP/F12.2", CONDI$, "NOMESS")
    Call SELECHO("LISRECIB")
    NURECIBO$ = TRIM$(VALACT1$("LISRECIB"))
    If NURECIBO$ = "" Then GoTo 99
    '
    TXTNURECIBO = NURECIBO$
    
99  Command6.Enabled = True

End Sub


Private Sub Form_Load()
'Call APLICAR_SKIN(Me, LUDAT$ & "\SKINS\dogmax.skn")
    UTILIZA_SKIN% = 1
    TXTFECHA = FECHATEX$(HOY(HO$))
    TXTUSUARIO = USERNAME$

    Call APLICACIONSKINS(Me)
    '
End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub









Private Sub mnurevertir_rec_Click()
    CONDI$ = "TIPOMOVIM= 'RHON' AND STATUS = 9 ORDER BY CODICOM_COR"
    Call CARGALISEL("RECIBANU", "CAJABANC", "CODICOM_COR/A12;NUCLIEN/I5+DEUDOR12/A32;FECHEMISI/D8;IHABECOMP/F12.2", CONDI$, "NOMESS")
    Call SELECHO("RECIBANU")
    NURECIBO$ = TRIM$(VALACT1$("RECIBANU"))
    If NURECIBO$ = "" Then Exit Sub
    '
   
    If COMALTER%("Confirmación de Revertir  Anulación\\Cancelar\Confirma Revertir Anulación de  Recibo: " & NURECIBO$ & "") = 2 Then
        CONDI$ = "CODICOM_LAR = '" & NURECIBO$ & "' AND (TIPOMOVIM= 'RHON' OR TIPOMOVIM= 'HONO')"
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
        Call COMUNI("Anulación Revertida Completada")
    End If
    '
    Call BLANFORMU

End Sub

Private Sub TXTCONCEPTO_KeyPress(KeyAscii As Integer)
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
'   Call BLANFORMU
'   NC% = XVALO(TXTID)
'   If NC% < 1 Then Exit Sub
'   '
'   CONDI$ = "Nume_Cli = " & NC% & " And Stat_Cli >= 0"
'   If CantRegistros("DEUDOR12", CONDI$, "NOMESS") > 0 Then
'
'      SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE NUME_CLI = " & NC% & " And Stat_Cli >= 0"
'      Set RS = READSET(SQLX$)
'      With RS
'        If Not .EOF Then
'            TXTNOMBRE = SAFETEXT$(!Raso_Cli)
'        End If
'      End With
'   End If
End Sub

Private Sub Text4_Change()
   TXTID = Text4
   
End Sub
Sub BLANFORMU()
        MANREC14.TXTNOMBRE = ""
        MANREC14.TXTDEBE = ""
        MANREC14.TXTCONCEPTO = ""
        MANREC14.TXTFECHA = ""
        MANREC14.TXTHABER = ""
        MANREC14.TXTID = ""
        MANREC14.TXTNURECIBO = ""
        MANREC14.TXTUSUARIO = ""
        txtusuariorealizo = ""
End Sub


Private Sub TXTIMPORTE_Change()

End Sub

Private Sub TXTIMPORTE_GotFocus()
   TXTIMPORTE = FORMATNUM$(XVALO(TXTIMPORTE), "F12.2")
End Sub

Private Sub TXTIMPORTE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TXTNOMBRE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub TXTNURECIBO_Change()
   NURECI$ = TRIM$(TXTNURECIBO)
   If NURECI$ = "" Then Exit Sub
   '
   If DERACCE%("REVERREC", "Revertir Eliminación de Recibo") < 2 Then Exit Sub
   
   
   CONDI$ = "CODICOM_COR = '" & NURECI$ & "'"
   MANREC14.TXTCONCEPTO = VALOBADA("CAJABANC", "DESCRIMOVI", CONDI$, "NOMESS")
   MANREC14.TXTDEBE = FORMATNUM$(XVALO(VALOBADA("CAJABANC", "IDEBECOMOP", CONDI$, "NOMESS")), "F12.2")
   MANREC14.TXTHABER = FORMATNUM$(XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS")), "F12.2")
   MANREC14.TXTFECHA = CVDAT(CVINT(VALOBADA("CAJABANC", "FECHEMISI", CONDI$, "NOMESS")))
   MANREC14.TXTID = XVALO(VALOBADA("CAJABANC", "NUCLIEN", CONDI$, "NOMESS"))
   MANREC14.TXTNOMBRE = RASOCLI$(MANREC14.TXTID)
   MANREC14.txtusuariorealizo = XVALO(VALOBADA("CAJABANC", "NUMEUSER", CONDI$, "NOMESS"))
   '
End Sub
