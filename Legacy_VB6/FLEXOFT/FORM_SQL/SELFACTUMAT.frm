VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SELFACTUMAT 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5895
   ClientLeft      =   -3135
   ClientTop       =   555
   ClientWidth     =   10350
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10350
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      Caption         =   "NPROV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Mostrar Solo Facturas Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6960
         TabIndex        =   28
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.TextBox Text7 
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Buscar a partir de:"
         Top             =   360
         Width           =   1680
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
         Left            =   6150
         Picture         =   "SELFACTUMAT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Actualizar"
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox Text8 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Max.Aplic."
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox txtmaxaplic 
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
         Left            =   5160
         TabIndex        =   25
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox TXTPROVGASTOS 
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
         Left            =   6480
         TabIndex        =   24
         Top             =   630
         Width           =   1200
      End
      Begin VB.TextBox Text2 
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Prov. Gastos"
         Top             =   360
         Width           =   1200
      End
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
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   240
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Pend. de Aplicar"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1680
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Proveedor:"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox Text3 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Factura de Gastos N°"
         Top             =   360
         Width           =   2000
      End
      Begin VB.TextBox TXTNFACTURA 
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
         Left            =   7680
         TabIndex        =   18
         Top             =   630
         Width           =   2000
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Tot. Aplicado:"
         Top             =   4680
         Width           =   1680
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Total Fact.Gastos"
         Top             =   3240
         Width           =   1680
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4920
         TabIndex        =   12
         Top             =   630
         Width           =   240
      End
      Begin VB.TextBox TXTFECHA 
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
         Left            =   3480
         TabIndex        =   11
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   600
         Width           =   1680
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TXTNUMEPROVE 
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
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TXTMOVIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   5880
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Costo segun el proveedor y la via de entrega"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   5325
         Width           =   1680
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "Facturas Seleccionadas para Aplicar Gastos"
         Top             =   3240
         Width           =   8280
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   1875
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   3307
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   16501405
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
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2235
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   3942
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   8421504
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
      Begin VB.Label lblpendiente 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   21
         Top             =   4230
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   8520
         TabIndex        =   14
         Top             =   4950
         Width           =   1680
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   13
         Top             =   3515
         Width           =   1680
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3240
      Top             =   5040
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "SELFACTUMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command5_Click()
  Command5.Enabled = False
         Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         TXTNUMEPROVE = NPRO%
   Command5.Enabled = True
End Sub

Private Sub Check1_Click()

   If XVALO(TXTNUMEPROVE) > 0 Then
     Call CARGA_FACTURAS_PROVEEDOR(XVALO(TXTNUMEPROVE), MSF1)
   End If
   
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If lblpendiente.Visible = True Then
      IMPORTE_FACT_GASTOS# = XVALO(lblpendiente)
   Else
      IMPORTE_FACT_GASTOS# = XVALO(Label8)
   End If
   NPROV = XVALO(TXTNUMEPROVE)
   If NPROV < 1 Then
     Call COMUNI("Falta Ingresar Proveedor.")
     GoTo 99
   End If
   If XVALO(Label5) > IMPORTE_FACT_GASTOS# Then
      Call COMUNI("Total Aplicado es Superior al Total de la Factura o Supera el Saldo Pendiente")
      GoTo 99
   End If
   '
   SIG% = 1
   If Left(TXTNFACTURA, 3) = "NC-" Then SIG% = -1
   For i& = 1 To SELFACTUMAT.MSF2.Rows - 1
20      RNDX$ = RANDSTRING$(-10)
        CONDI1$ = "FACTURAPROV ='" & RNDX$ & "'"
        If CantRegistros("FACTURAYGASTOS", CONDI1$, "NOMESS") > 0 Then GoTo 20
        Call CreaRegistro("FACTURAYGASTOS", "FACTURAPROV", RNDX$, "NOMESS")
        '
        Call ACTUREGISTRO("FACTURAYGASTOS", "VAL_COMPRA", CONDI1$, 0, RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", "FECHEMISI", CONDI1$, FETEXCOR1$(HOY(HO$)), RAFE&, "NOMESS")
        MONEX% = MONEMI%
        IMPOFAC = XVALO(Label8)
        IMPOGASTO = XVALO(MSF2.TextMatrix(i&, 5))
        NPROV = XVALO(TXTNUMEPROVE)
        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI1$, FORMATNUM$((SIG% * IMPOFAC), "F16.5"), RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", "IMPO_GASTO", CONDI1$, FORMATNUM$((SIG% * IMPOGASTO), "F16.5"), RAFE&, "NOMESS")
            FACTURA_ASOCIADA$ = SELFACTUMAT.MSF2.TextMatrix(i&, 2)
        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV_ASOC", CONDI1$, FACTURA_ASOCIADA$, RAFE&, "NOMESS")
        
        Call ACTUREGISTRO("FACTURAYGASTOS", "TIPO_FACTURA", CONDI1$, "G", RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", " Nume_Cli", CONDI1$, XVALO(Me.TXTPROVGASTOS), RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", "MoneVal", CONDI1$, MONEMI%, RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", "TipoCam", CONDI1$, FORMATNUM$(TICAMONE(MONEMI%), "F16.5"), RAFE&, "NOMESS")
        Call ACTUREGISTRO("FACTURAYGASTOS", "Nume_Cli_Asoc", CONDI1$, XVALO(TXTNUMEPROVE), RAFE&, "NOMESS")
        '
        Call ACTUREGISTRO("FACTURAYGASTOS", "FACTURAPROV", CONDI1$, TXTNFACTURA, RAFE&, "NOMESS")
    Next i&
    Unload Me

99 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        TXTFECHA.TEXT = VDEVV$
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("ACREDOR")
    
    TXTNUMEPROVE = XVALO(TRIM$(VALACT1$("ACREDOR")))
    
    Command3.Enabled = True
End Sub

Private Sub Command43_Click()
   NPROV = XVALO(Me.TXTNUMEPROVE)
   Me.TXTNUMEPROVE = ""
   Me.TXTNUMEPROVE = NPROV
End Sub

Private Sub Form_Load()

    Call CARGA_ENCABEZADO(MSF1, "Fecha/D8;Factura/A20;Importe/F12.2;Yá Aplicado/F12.2;Nueva Aplic. /F12.2", SELFACTUMAT)
    Call CARGA_ENCABEZADO(MSF2, "Fecha/D8;Factura/A20;Importe/F12.2;Yá Aplicado/F12.2;Nueva Aplic. /F12.2", SELFACTUMAT)
    TXTFECHA = FECHATEX(DIANTE(HOY(HO$), 90))
    Me.txtmaxaplic = 2

    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    
End Sub

Private Sub MSF1_DblClick()
   If (MSF1.MouseRow = 0) Then Exit Sub
   REG& = MSF1.Row
   
   MSF2.Rows = MSF2.Rows + 1
   MSF2.Cols = MSF1.Cols
   MSF2.TextMatrix(MSF2.Rows - 1, 1) = MSF1.TextMatrix(REG&, 1)
   MSF2.TextMatrix(MSF2.Rows - 1, 2) = MSF1.TextMatrix(REG&, 2)
   MSF2.TextMatrix(MSF2.Rows - 1, 3) = MSF1.TextMatrix(REG&, 3)
   MSF2.TextMatrix(MSF2.Rows - 1, 4) = MSF1.TextMatrix(REG&, 4)
   MSF2.TextMatrix(MSF2.Rows - 1, 5) = 0
   '
   'ACTIVA LA CAJA DE TEXTO DE EDICION
   If MSF2.Rows > 1 Then
    MSF2.COL = 5
    MSF2.Row = MSF2.Rows - 1
    With TXTMOVIL
      .ZOrder 0
      .Width = MSF2.CellWidth
      .Left = MSF2.CellLeft + MSF2.Left
      .Top = MSF2.CellTop + MSF2.Top
      .TEXT = MSF2.TEXT
      .Visible = True
      .ToolTipText = MSF2.Row
      .SetFocus
    End With
   End If

   
'   Label5 = FORMATNUM$(XVALO(Label5) - XVALO(MSF1.TextMatrix(REG&, 3)), "F12.2") 'TOTAL FACTURA DE GASTOS
'   Label8 = FORMATNUM$(XVALO(Label8) + XVALO(MSF1.TextMatrix(REG&, 3)), "F12.2")  'TOTAL SELECCIONADO
   '
   Call borrafila(MSF1.Row, Me, Me.MSF1)
   MSF2.Sort = 1
   Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
End Sub

Private Sub msf2_Click()

          COLUM = MSF2.MouseCol
          If COLUM <> 5 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          MSF2.COL = 5
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF2.CellWidth
            .Left = MSF2.CellLeft + MSF2.Left
            .Top = MSF2.CellTop + MSF2.Top
            .TEXT = MSF2.TEXT
            .Visible = True
            .ToolTipText = MSF2.Row
            .SetFocus
          End With

End Sub

Private Sub MSF2_DblClick()
   If (MSF2.MouseRow = 0) Then Exit Sub
   
   REG& = MSF2.Row
   MSF1.Rows = MSF1.Rows + 1
   MSF1.Cols = MSF2.Cols
   MSF1.TextMatrix(MSF1.Rows - 1, 1) = MSF2.TextMatrix(REG&, 1)
   MSF1.TextMatrix(MSF1.Rows - 1, 2) = MSF2.TextMatrix(REG&, 2)
   MSF1.TextMatrix(MSF1.Rows - 1, 3) = MSF2.TextMatrix(REG&, 3)
   MSF1.TextMatrix(MSF1.Rows - 1, 4) = MSF2.TextMatrix(REG&, 4)
   MSF1.TextMatrix(MSF1.Rows - 1, 5) = 0
   
   Call borrafila(Me.MSF2.Row, Me, Me.MSF2)
   Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
End Sub

Private Sub Text4_Change()

End Sub

Private Sub txtmaxaplic_DblClick()
   If txtmaxaplic.Locked = True Then
     txtmaxaplic.Locked = False
     txtmaxaplic.BackColor = vbWhite
     
   Else
     txtmaxaplic.Locked = True
     txtmaxaplic.BackColor = &HE0E0E0
   End If
   
   
  
End Sub

Private Sub Text4_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub TXTFECHA_Change()
   
  Call CARGA_FACTURAS_PROVEEDOR(NPROV, MSF1)
End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub


Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     
      TXTMOVIL.ToolTipText = REG&
      MSF2.TEXT = FORMATNUM$(XVALO(TXTMOVIL), "F12.2")
      TXTMOVIL.Visible = False
      Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
   End If
End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
       REG& = XVALO(TXTMOVIL.ToolTipText)
       MSF2.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL), "F12.2")
       TXTMOVIL.Visible = False
     End If
     Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
End Sub

Private Sub TXTNUMEPROVE_Change()
    NPROV = XVALO(TXTNUMEPROVE)
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
    MSF2.Rows = 1
    MONEX% = MONEMI%
    IMPOFAC# = XVALO(Label8)
    If MONEX% < 2 Then
        TICAMB# = TICAMONE(2)
        COEFCON = TICAMB# / TICAMONE(1)
        IMPOFAC# = IMPOFAC# / COEFCON
    End If
    Screen.MousePointer = 11
    'Label8 = FORMATNUM$(IMPOFAC#, "F12.2") ' EL TOTAL DE LA FACTURA
    '
    
    SQLX1$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    SQLX1$ = SQLX1$ + " WHERE Nume_Cli = " & NPROV & " AND TIPO_FACTURA = 'M'"
    If TRIM$(TXTFECHA.TEXT) <> "" Then SQLX1$ = SQLX1$ + " AND FECHEMISI >= '" & FETEXCOR1$(FECHANUM(TXTFECHA.TEXT)) & "'" 'HABRIA QUE VALIDAR BUSCANDO LA FECHA DE LA FACTURA ORIGINAL
    Set RS1 = READSET(SQLX1$)
    MSF.Redraw = False
    With RS1
       Do While Not .EOF
         FACTURAPRO$ = SAFETEXT$(!FACTURAPROV)
         'VALIDO QUE ESTE CANCELADA (NO DEBE EXISTIR EN SACREPEN)
         If Check1.Value = 1 Then
            CONDI$ = "NUPROVE = " & NPROV
            CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & FACTURAPRO$ & "'"
            If CantRegistros("SACREPEN", CONDI$, "NOMESS") = 0 Then GoTo 50
         End If
         '
         'VALIDO QUE NO TENGA MAS DE 2 APLICACIONES.
         CONDI$ = " NUME_CLI_ASOC = " & NPROV
         CONDI$ = CONDI$ + " AND FACTURAPROV_ASOC = '" & FACTURAPRO$ & "' AND IMPO_GASTO > 0"
         APLIMAX% = XVALO(txtmaxaplic)
         If APLIMAX% < 1 Then APLIMAX% = 1000
         
         If CantRegistros("FACTURAYGASTOS", CONDI$, "NOMESS") >= APLIMAX% Then GoTo 50
         '
         'ESCRIBO SIEMPRE LAS 3 CAMPOS
         j& = j& + 1
         MSF.Rows = j& + 1
         
         MSF.TextMatrix(j&, 1) = CVDAT(CVINT(!FECHEMISI))
           FACTURAPROV$ = TRIM$(SAFETEXT$(!FACTURAPROV))
         MSF.TextMatrix(j&, 2) = FACTURAPROV$
         IMPOFAC# = XVALO(!IMPORTE_FACTURA)
         MONEX% = XVALO(!MONEVAL)
         If MONEX% < 2 Then
            TICAMB# = XVALO(!TIPOCAM)
            COEFCON = TICAMB# / TICAMONE(1)
            IMPOFAC# = IMPOFAC# / COEFCON
         End If

         MSF.TextMatrix(j&, 3) = FORMATNUM$(IMPOFAC#, "F12.2")
           CONDI$ = "TIPO_FACTURA = 'G' AND FACTURAPROV_ASOC = '" & FACTURAPROV$ & "' AND Nume_Cli_Asoc = " & XVALO(TXTNUMEPROVE)
         APLICADO# = XVALO(VALOBADA("FACTURAYGASTOS", "SUM(IMPO_GASTO)", CONDI$, "NOMESS"))
         MSF.TextMatrix(j&, 4) = TRIM$(FORMATNUM$(APLICADO#, "F12.2"))
         MSF.TextMatrix(j&, 5) = 0
         
50       .MoveNext
       Loop
    End With
    MSF.Redraw = True
    RS1.Close
    Set RS1 = Nothing
    Screen.MousePointer = 1

End Sub
Function ACUMULA_SELEC()
   ACUMULA = 0
   For i& = 1 To MSF1.Rows - 1
     ACUMULA = ACUMULA + XVALO(MSF1.TextMatrix(i&, 3))
   Next i&
   ACUMULA_SELEC = ACUMULA
End Function
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    MSF2.COL = 5
    Select Case KeyCode
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF2.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F12.2")
            'MSF2.SetFocus
            If MSF2.Row > 1 Then MSF2.Row = MSF2.Row - 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF2.CellWidth
            .Left = MSF2.CellLeft + MSF2.Left
            .Top = MSF2.CellTop + MSF2.Top
            
             .TEXT = FORMATNUM$(MSF2.TEXT, "F12.2")
            .Visible = True
            .ToolTipText = MSF2.Row
            .SetFocus
          End With
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            'MSF2.SetFocus
            MSF2.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F12.2")
            If MSF2.Row < MSF2.Rows Then MSF2.Row = MSF2.Row + 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF2.CellWidth
            .Left = MSF2.CellLeft + MSF2.Left
            .Top = MSF2.CellTop + MSF2.Top
             .TEXT = FORMATNUM$(MSF2.TEXT, "F12.2")
            .Visible = True
            .ToolTipText = MSF2.Row
            .SetFocus
          End With
            On Error GoTo 0
    End Select
    Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
    '
End Sub

