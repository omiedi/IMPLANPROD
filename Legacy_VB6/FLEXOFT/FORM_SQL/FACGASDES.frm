VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FACGASDES 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Seleccionar Gasto a Re-Aplicar"
   ClientHeight    =   4680
   ClientLeft      =   -3135
   ClientTop       =   555
   ClientWidth     =   10365
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton Command3 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   850
         TabIndex        =   7
         ToolTipText     =   "Selección de Proveedor"
         Top             =   630
         Width           =   240
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "Proveedor:"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   1090
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   9125
      End
      Begin VB.TextBox TXTNUMEPROVE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         TabIndex        =   4
         Top             =   630
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2955
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Importe es en la moneda original "
         Top             =   1200
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   5212
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   16501405
         BackColorSel    =   12632064
         SelectionMode   =   1
         Appearance      =   0
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
         Caption         =   "TEPRUEBA"
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   8415
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "FACGASDES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODO$
Private Sub Command5_Click()
  Command5.Enabled = False
         Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         TXTNUMEPROVE = NPRO%
   Command5.Enabled = True
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  
  Command1.Enabled = True
End Sub

Private Sub Form_Load()

    Call CARGA_ENCABEZADO(MSF1, "Fecha/D8;Factura/A20;Importe/F12.2;Se Aplicó A Fc./A24;Aplicado/F12.2", SELFACTUMAT)
    
End Sub


Private Sub MSF1_DblClick()
   If (MSF1.MouseRow = 0) Then Exit Sub
   '
   REG& = MSF1.MouseRow
   SELFACTUMAT.Text5(2).Visible = True
   SELFACTUMAT.lblpendiente.Visible = True
   SELFACTUMAT.TXTNFACTURA = MSF1.TextMatrix(REG&, 2)
   SELFACTUMAT.Label8 = TRIM$(FORMATNUM$(XVALO(MSF1.TextMatrix(REG&, 3)), "F15.2"))  'POR SI ES NOTA DE CREDITO LO PONE EN NEGATIVO
   NPROV = XVALO(Me.TXTNUMEPROVE)
   If FACGASDES.Caption = "Aplicación de Facturas de Gastos" Then
     CONDI$ = "NUME_CLI = " & NPROV & " AND FACTURAPROV = '" & MSF1.TextMatrix(REG&, 2) & "' and TIPO_FACTURA = 'M'"
     FACTURAMATERIALES% = CantRegistros&("FACTURAYGASTOS", CONDI$, "NOMESS")
     If FACTURAMATERIALES% > 0 Then
        Call COMUNI("Imposible Aplicar Como Gasto  - Factura de Materiales")
        Exit Sub
     End If
     SELFACTUMAT.lblpendiente = FORMATNUM$(XVALO(MSF1.TextMatrix(REG&, 3)) - GastosAplicados2(MSF1.TextMatrix(REG&, 2), NPROV, 1), "F12.2")
   Else
     SELFACTUMAT.lblpendiente = FORMATNUM$(XVALO(MSF1.TextMatrix(REG&, 3)) - XVALO(MSF1.TextMatrix(REG&, 5)), "F12.2")
   End If
   SELFACTUMAT.TXTPROVGASTOS = NPROV
   SELFACTUMAT.Show 1
End Sub

Private Sub TXTFECHA_Change()
   
End Sub
Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("ACREDOR")
    TXTNUMEPROVE = XVALO(TRIM$(VALACT1$("ACREDOR")))
    Command3.Enabled = True
End Sub


Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then
        If (MSF1.MouseRow = 0) Then Exit Sub
        REG& = MSF1.MouseRow
        MSF1.Row = REG&
        NUFACTU$ = MSF1.TextMatrix(REG&, 2)
        MSF1.SelectionMode = 1
        OPX% = COMALTER%("Desea Desaplicar Gasto de la Factura: " & NUFACTU$ & " \\No, Cancelar\Sí Desaplicar")
        If OPX% < 2 Then Exit Sub
        NPROV = Me.TXTNUMEPROVE
        CONDI$ = "NUME_CLI = " & NPROV
        CONDI$ = CONDI$ + " AND FACTURAPROV = '" & NUFACTU$ & "'"
        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI$, 0, REGAF&, "NOMESS")
        Call TXTNUMEPROVE_Change
        '
     End If

End Sub

Private Sub TXTNUMEPROVE_Change()
    NPROV = XVALO(TXTNUMEPROVE)
    MSF1.Rows = 1
    If NPROV > 0 Then
      Text1 = ECOARCH$("PROVED1", MKI$(NPROV))
      Call CARGA_FACTURAS_PROVEEDOR(NPROV, MSF1)
    Else
      Text1 = ""
    End If
    
End Sub
Sub CARGA_FACTURAS_PROVEEDOR(NPROV, MSF As MSFlexGrid)
    '
    If NPROV < 1 Then Exit Sub
    MSF.Rows = 1
    '
    MSF.Rows = 1
    '
    If FACGASDES.Caption = "Aplicación de Facturas de Gastos" Then
       GoTo 50
       Exit Sub
    End If
    Call CARGA_ENCABEZADO(MSF1, "Fecha/D8;Factura/A20;Importe/F12.2;Se Aplicó A Fc./A24;Aplicado/F12.2", SELFACTUMAT)

    'PRESENTO LAS FACTURAS DE GASTOS.
    SQLX1$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    SQLX1$ = SQLX1$ + " WHERE Nume_Cli = " & NPROV & " "
    SQLX1$ = SQLX1$ + " AND TIPO_FACTURA = 'G'"
    SQLX1$ = SQLX1$ + " AND IMPO_GASTO > 0 "
    SQLX1$ = SQLX1$ + " ORDER BY FECHEMISI"
    Set RS1 = READSET(SQLX1$)
    J& = 0
    With RS1
       Do While Not .EOF
         '
         FACTURAPROV$ = TRIM$(SAFETEXT$(!FACTURAPROV))
         IMPORTE_FACT# = XVALO(!IMPORTE_FACTURA)
         IMPORTE_APL# = XVALO(!IMPO_GASTO)
         FACTURADONDESEAPLICO$ = TRIM$(SAFETEXT$(!FACTURAPROV_ASOC))
         TICAMB# = XVALO(!TIPOCAM)
         MONEX% = XVALO(!MONEVAL)
         COEFCON = TICAMB# / TICAMONE(1)
         If MONEX% > 1 Then
           IMPORTE_FACT# = COEFCON * IMPORTE_FACT#
           IMPORTE_APL# = COEFCON * IMPORTE_APL#
         End If
             '
         J& = J& + 1
         MSF.Rows = J& + 1
         MSF.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHEMISI))
         '
         MSF.TextMatrix(J&, 2) = FACTURAPROV$
         MSF.TextMatrix(J&, 3) = FORMATNUM$(IMPORTE_FACT#, "F12.2")
         MSF.TextMatrix(J&, 4) = FACTURADONDESEAPLICO$
         MSF.TextMatrix(J&, 5) = FORMATNUM$(IMPORTE_APL#, "F12.2")
         
         .MoveNext
       Loop
    End With
    '
    If J& > 0 Then
      On Error Resume Next
      FACGASDES.Show 1
      On Error GoTo 0
    Else
      Call COMUNI("No se Encontraron Facturas de Gastos")
    End If
    Exit Sub
 '''''''''''''''''''''
50  Call CARGA_ENCABEZADO(MSF1, "Fecha/D8;Factura/A24;Importe/F18.2", SELFACTUMAT)
  
    SQLX$ = "SELECT CODICOM_LAR, Tipo_Cam, Mone_Emis, FECHEMISI, SUM(ValAbsComp) AS IMPORTE FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'COMPRA') "
    SQLX$ = SQLX$ + " AND nuprove = " & NPROV
'    SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & "' "
'    SQLX$ = SQLX$ + "AND FechConta > '" & FECHALU & "' "
    SQLX$ = SQLX$ + "GROUP by CODICOM_LAR, Tipo_Cam, Mone_Emis, FECHEMISI "
    SQLX$ = SQLX$ + "Order by FECHEMISI ASC "

    '
    Set RS = READSET(SQLX$)
    With RS
        Do While Not .EOF
          MONEX% = XVALO(!Mone_Emis)
          TICAMB# = XVALO(!Tipo_Cam)
          IMPO# = XVALO(!IMPORTE)
          If IMPO# >= 0.005 Then
               If MONEX% > 1 And TICAMB# > 0.00005 Then
                  IMPO# = IMPO# / TICAMB#
               End If
               FACTURAPROV$ = SAFETEXT$(!CODICOM_LAR)

               IMPORTE_FACT1$ = FORMATNUM$(IMPO#, "F12.2")
    
               J& = J& + 1
               MSF.Rows = J& + 1
               MSF.TextMatrix(J&, 1) = CVDAT(CVINT(!FECHEMISI))
               '
               MSF.TextMatrix(J&, 2) = FACTURAPROV$
               MSF.TextMatrix(J&, 3) = IMPORTE_FACT1$
          End If
70      .MoveNext
        Loop
    End With
    RS.Close
    Set RS = Nothing
    









End Sub
