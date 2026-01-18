VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ANUFACTGAS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplicaciones de Factura Seleccionada"
   ClientHeight    =   4350
   ClientLeft      =   -3135
   ClientTop       =   555
   ClientWidth     =   10365
   DrawWidth       =   10
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
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
      Begin VB.TextBox TXTIIMPORTE 
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
         Left            =   8640
         TabIndex        =   13
         Top             =   630
         Width           =   1560
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Importe "
         Top             =   360
         Width           =   1560
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
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Proveedor:"
         Top             =   360
         Width           =   4335
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Tipo de Factura"
         Top             =   360
         Width           =   1680
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Factura Seleccionada:"
         Top             =   360
         Width           =   2160
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
         Left            =   4560
         TabIndex        =   9
         Top             =   630
         Width           =   2160
      End
      Begin VB.TextBox TXTTIPOFACTURA 
         Alignment       =   2  'Center
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
         Left            =   6840
         TabIndex        =   7
         Top             =   630
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   3615
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
         TabIndex        =   5
         Top             =   630
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anular - Factura Seleccionada y Aplicaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5040
      End
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   2475
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   4366
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
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         Height          =   195
         Left            =   0
         TabIndex        =   8
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
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "ANUFACTGAS"
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
Sub Command1_Click()
   Command1.Enabled = False
   '
   If TRIM$(TXTNFACTURA) = "" Then GoTo 99

   CONDI$ = "FACTURAPROV = '" & TXTNFACTURA & "' and TIPO_FACTURA = 'M' AND NUME_CLI = " & Me.TXTNUMEPROVE
   Call BORRAREGISTROS("FACTURAYGASTOS", CONDI$, REGAF&, "NOMESS")
   CONDI$ = "FACTURAPROV_ASOC = '" & TXTNFACTURA & "' and TIPO_FACTURA = 'G' AND NUME_CLI_ASOC = " & Me.TXTNUMEPROVE
   Call BORRAREGISTROS("FACTURAYGASTOS", CONDI$, REGAF&, "NOMESS")
   '
   MODO$ = "BORRAR"
   On Error Resume Next
   Hide
   On Error GoTo 0
   '
99 Command1.Enabled = True
End Sub

Private Sub Command2_Click()
End Sub

Private Sub Form_Load()

    Call CARGA_ENCABEZADO(MSF1, "Fecha/D8;Factura/A20;Importe/F12.2;Yá Aplicado/F12.2;Nueva Aplic. /F12.2", SELFACTUMAT)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Cancel = 0 Then
     MODO$ = "SALIR"
  End If
End Sub

Private Sub MSF1_DblClick()
   If (MSF1.MouseRow = 0) Then Exit Sub
   REG& = MSF1.Row
   
End Sub

Private Sub msf2_Click()

          COLUM = msf2.MouseCol
          If COLUM <> 5 Then
            TXTMOVIL.Visible = False
            Exit Sub
          End If
          msf2.COL = 5
          With TXTMOVIL
            .ZOrder 0
            .Width = msf2.CellWidth
            .Left = msf2.CellLeft + msf2.Left
            .Top = msf2.CellTop + msf2.Top
            .TEXT = msf2.TEXT
            .Visible = True
            .ToolTipText = msf2.Row
            .SetFocus
          End With

End Sub

'Private Sub MSF2_DblClick()
'   If (MSF2.MouseRow = 0) Then Exit Sub
'
'   REG& = MSF2.Row
'   MSF1.Rows = MSF1.Rows + 1
'   MSF1.Cols = MSF2.Cols
'   MSF1.TextMatrix(MSF1.Rows - 1, 1) = MSF2.TextMatrix(REG&, 1)
'   MSF1.TextMatrix(MSF1.Rows - 1, 2) = MSF2.TextMatrix(REG&, 2)
'   MSF1.TextMatrix(MSF1.Rows - 1, 3) = MSF2.TextMatrix(REG&, 3)
'   MSF1.TextMatrix(MSF1.Rows - 1, 4) = MSF2.TextMatrix(REG&, 4)
'   MSF1.TextMatrix(MSF1.Rows - 1, 5) = MSF2.TextMatrix(REG&, 5)
'
'   Call borrafila(Me.MSF2.Row, Me, Me.MSF2)
'   Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.MSF2, 5), "F12.2")
'End Sub

Private Sub TXTFECHA_Change()
   
End Sub



Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     
      TXTMOVIL.ToolTipText = REG&
      msf2.TEXT = FORMATNUM$(XVALO(TXTMOVIL), "F12.2")
      TXTMOVIL.Visible = False
      Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.msf2, 5), "F12.2")
   End If
End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
       REG& = XVALO(TXTMOVIL.ToolTipText)
       msf2.TextMatrix(REG&, 5) = FORMATNUM$(XVALO(TXTMOVIL), "F10.1")
       TXTMOVIL.Visible = False
     End If
     Label5 = FORMATNUM$(SUMA_COLUMNA_SELECCION(SELFACTUMAT.msf2, 5), "F12.2")
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
    MSF.Rows = 1
    Label8 = FORFAC07.Label43 ' EL TOTAL DE LA FACTURA
    '
    '
    SQLX1$ = "SELECT * FROM FACTURAYGASTOS WITH(NOLOCK)"
    If ANUFACTGAS.TXTTIPOFACTURA = "MATERIALES" Then
        SQLX1$ = SQLX1$ + " WHERE Nume_Cli_Asoc = " & NPROV & " "
        SQLX1$ = SQLX1$ + " AND  FACTURAPROV_ASOC ='" & ANUFACTGAS.TXTNFACTURA & "' AND TIPO_FACTURA = 'G' AND Nume_Cli_Asoc = " & NPROV
    End If
    Set RS1 = READSET(SQLX1$)
    With RS1
       Do While Not .EOF
         'ESCRIBO SIEMPRE LAS 3 CAMPOS
         j& = j& + 1
         MSF.Rows = j& + 1
         '
         MSF.TextMatrix(j&, 1) = CVDAT(CVINT(!FECHEMISI))
         If ANUFACTGAS.TXTTIPOFACTURA = "MATERIALES" Then
           FACTURAPROV$ = TRIM$(SAFETEXT$(!FACTURAPROV))
           IMPOFAC# = XVALO(!IMPORTE_FACTURA)
           IMPOGASTO# = XVALO(!IMPO_GASTO)
           MONEX% = XVALO(!MoneVal)
           If MONEX% > 1 Then
             TICAMB# = XVALO(!TipoCam)
'             COEFCON = TICAMB# / TICAMONE(1)
             IMPOFAC# = IMPOFAC# 'COEFCON *
             IMPOGASTO# = IMPOGASTO#  'COEFCON *
           End If
'           IMPORTE$ = FORMATNUM$(IMPOGASTO#, "F12.2")
'           IMPORTE2$ = FORMATNUM$(IMPOFAC#, "F12.2")
         Else
           FACTURAPROV$ = TRIM$(SAFETEXT$(!FACTURAPROV_ASOC))
           CONDI$ = "FACTURAPROV = '" & FACTURAPROV$ & "' AND TIPO_FACTURA = 'M'"
           IMPORTE$ = FORMATNUM$(XVALO(VALOBADA("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI$, "NOMESS")), "F12.2")
           MONEX% = XVALO(VALOBADA("FACTURAYGASTOS", "MoneVal", CONDI$, "NOMESS"))
           TICAMB# = XVALO(VALOBADA("FACTURAYGASTOS", "TipoCam", CONDI$, "NOMESS"))
           If MONEX% > 1 Then
             COEFCON = TICAMB# / TICAMONE(1)
             IMPOFAC# = COEFCON * IMPOFAC#
             IMPOGASTO# = COEFCON * IMPOGASTO#
           End If
'           IMPORTE$ = IMPORTE$
'           IMPORTE2$ = FORMATNUM$(XVALO(!IMPO_GASTO), "F12.2")
         End If
         '
         IMPORTE$ = FORMATNUM$(IMPOGASTO#, "F12.2")
         IMPORTE2$ = FORMATNUM$(IMPOFAC#, "F12.2")
         '
         MSF.TextMatrix(j&, 2) = FACTURAPROV$
         MSF.TextMatrix(j&, 3) = IMPORTE$
         MSF.TextMatrix(j&, 4) = IMPORTE2$
         MSF.TextMatrix(j&, 5) = ""
         '
         .MoveNext
       Loop
    End With
    '
    ANUFACTGAS.Show 1
       
End Sub
Function ACUMULA_SELEC()
   ACUMULA = 0
   For I& = 1 To MSF1.Rows - 1
     ACUMULA = ACUMULA + XVALO(MSF1.TextMatrix(I&, 3))
   Next I&
   ACUMULA_SELEC = ACUMULA
End Function
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    MSF1.COL = 5
    Select Case KeyCode
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF1.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F12.2")
            'MSF1.SetFocus
            If MSF1.Row > 1 Then MSF1.Row = MSF1.Row - 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            
             .TEXT = FORMATNUM$(MSF1.TEXT, "F12.2")
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            'MSF1.SetFocus
            MSF1.TEXT = FORMATNUM$(TXTMOVIL.TEXT, "F12.2")
            If MSF1.Row < MSF1.Rows Then MSF1.Row = MSF1.Row + 1
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
             .TEXT = FORMATNUM$(MSF1.TEXT, "F12.2")
            .Visible = True
            .ToolTipText = MSF1.Row
            .SetFocus
          End With
            On Error GoTo 0
    End Select
    '
End Sub

