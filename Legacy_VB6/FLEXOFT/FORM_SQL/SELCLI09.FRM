VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELCLI09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Clientes con Deuda"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9045
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ListBox LISCLISE 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4365
      ItemData        =   "SELCLI09.frx":0000
      Left            =   50
      List            =   "SELCLI09.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde para seleccionar"
      Top             =   410
      Width           =   8025
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   8145
      ScaleHeight     =   5745
      ScaleWidth      =   1485
      TabIndex        =   3
      Top             =   -120
      Width           =   1515
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "SELCLI09.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Búsqueda de Cliente"
         Top             =   970
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "SELCLI09.frx":136A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   285
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Left            =   120
         Picture         =   "SELCLI09.frx":14B4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1670
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   120
         Picture         =   "SELCLI09.frx":15FE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Envia Mail a Clientes Seleccionados"
         Top             =   2360
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "SELCLI09.frx":1908
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   6360
      OleObjectBlob   =   "SELCLI09.frx":1B3C
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4995
      TabIndex        =   14
      Top             =   165
      Width           =   840
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   12
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Label77 
      BackStyle       =   0  'Transparent
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1155
      TabIndex        =   11
      Top             =   165
      Width           =   1455
   End
   Begin VB.Label Label81 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8715
      TabIndex        =   10
      Top             =   1125
      Width           =   840
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5500
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   5
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   50
      TabIndex        =   13
      Top             =   120
      Width           =   8025
   End
End
Attribute VB_Name = "SELCLI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'

Private Sub BOTREC_Click()
  For U& = 1 To LISCLISE.ListCount
       LISCLISE.ListIndex = U& - 1
         If LISCLISE.Selected(U& - 1) = True Then
           TX$ = LISCLISE.TEXT
           MAI$ = TRIM$(Mid$(TX$, 40, 32))
           NC% = XVALO(Mid$(TX$, 1, 6))
             If MAI$ <> "" And InStr(MAI$, "@") > 0 Then
               Call mailvenci(NC%)
             End If
         End If
   Next U&
   Call COMUNI("Envio de Deuda por Mail Finalizado")
End Sub

Private Sub Command1_Click()
   command1.Enabled = False
   TPI = LISCLISE.TopIndex
   For U& = 1 To LISCLISE.ListCount
     If LISCLISE.Selected(U& - 1) = True Then
         LISCLISE.Selected(U& - 1) = False
       Else
         LISCLISE.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LISCLISE.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   command1.Enabled = True

End Sub

Private Sub Command2_Click()
   LISCLISE.Clear
   Unload Me
End Sub

Sub mailvenci(NC%)
   ENCA$ = ""
   Call REPLA(ENCA$, "Nro Factura", 1, 18)
   Call REPLA(ENCA$, "   Imp.Bruto", 20, 12)
   Call REPLA(ENCA$, " F.Emis.", 34, 8)
   Call REPLA(ENCA$, "F.Vencim.", 44, 8)
   Call REPLA(ENCA$, "       Saldo", 54, 12)
   ENCA$ = ENCA$ + vbCrLf + String$(66, "=") + vbCrLf
   
   FEX = HOY(HO$)
   FEXU% = DIASPOST(FEX, 5)
   FEVE = FETEXCOR1$(FEXU%)
   Dim SADTEMP As ADODB.Recordset
   Set SADTEMP = New ADODB.Recordset
   SQLX$ = "SELECT * FROM sadeupen with (NoLock)"
   SQLX$ = SQLX$ + " WHERE fevencim <= '" & FEVE & " '"
   SQLX$ = SQLX$ + " AND Nuclien = " & NC%
   SQLX$ = SQLX$ + " AND  salddebe > 0.005"
   SQLX$ = SQLX$ + " ORDER BY codicom_lar "
   
   Set SADTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With SADTEMP
      Do While (Not .EOF)
         TX$ = ""
         NFC$ = SAFETEXT$(!codicom_lar)
         IB# = XVALO(!IBruComp)
         FEM% = CVINT(!FECHEMISI)
         FEVEN% = CVINT(!FEVENCIM)
         SALDO# = XVALO(!SaldDebe) - XVALO(!SaldAcre)
         Call REPLA(TX$, NFC$, 1, 18)
         Call REPLA(TX$, FORMATNUM$(IB#, "f12.2"), 20, 12)
         Call REPLA(TX$, FECHATEX$(FEM%), 34, 8)
         Call REPLA(TX$, FECHATEX$(FEVEN%), 44, 8)
         Call REPLA(TX$, FORMATNUM$(SALDO#, "f12.2"), 54, 12)
         TXTX$ = TXTX$ + TX$ + vbCrLf
         
         .MoveNext
      Loop
     End With
     
     TXTX$ = TXTX$ + vbCrLf + String$(66, "=")
     ENCA2$ = "Sres." + vbCrLf + rasocli(NC%) + vbCrLf + vbCrLf
     ENCA2$ = ENCA2$ + "Nos dirigimos a Uds a fin de recordarles " + vbCrLf
     ENCA2$ = ENCA2$ + "el proximo vencimiento de la(s) siguiente(s) facturas(s)" + vbCrLf
     textox$ = ENCA2$ + vbCrLf + ENCA$
'     TEXTOX$ = TEXTOX$ + vbCrLf + TXTX$
     textox$ = textox$ + vbCrLf + TXTX$ + vbCrLf + OBSERVAMAILDEUDA$

     '
     MAI_DESTI$ = rasocli$(NC%)
     MAI_DIREL$ = EMAILCLI(NC%)
     MAI_TEXTO$ = textox$
     MAI_ASUNT$ = "Informe de Saldo"
     Call GENERAMAIL
      
End Sub
Function OBSERVAMAILDEUDA$()
      '
      'DEVUELVE LAS OBSERVACIONES DE MAIL DE DEUDA QUE SE CARGAN EN SETUP DE CUENTAS POR COBRAR
      '
      OBSERVAMAILDEUDA$ = ""
      '
      Call COPYSTRU("OBSERVOF", "OBSMAIDE")
      For U& = 1 To ULTREG&("OBSMAIDE")
           TX$ = REGLEIDO$("OBSMAIDE", U&)
           TEXTOBS$ = TEXTOBS$ & TRIM$(TX$) & vbCrLf
      Next U&
      '
      If TRIM$(TEXTOBS$) <> "" Then
         OBSERVAMAILDEUDA$ = TEXTOBS$
      End If
      '
      Call CIERRARCH("OBSMAIDE")
      Call CIERRARCH("OBSERVOF")

End Function
Private Sub Command3_Click()
    Call SELECHO("DEUDOR1")
    valor = TRIM$(VALACT1$("DEUDOR1"))
    If valor <> "" Then
      For U& = 1 To LISCLISE.ListCount
         VALOR_LIST = TRIM$(Mid(LISCLISE.List(U&), 1, 6))
         If valor = VALOR_LIST Then
            LISCLISE.ListIndex = U&
            GoTo 10
         End If
      Next U&
     End If
     Call COMUNI("Cliente No Encontrado")
     
10: Exit Sub
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture2)
End Sub

