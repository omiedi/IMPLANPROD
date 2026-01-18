VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CTRCORREL 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Correlatividad"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Ver Cuenta Corriente"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generar Documento"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8160
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5000
      Left            =   0
      TabIndex        =   21
      Top             =   3120
      Width           =   6045
      Begin VB.TextBox TXTCLIENTE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         TabIndex        =   27
         Top             =   4580
         Width           =   5055
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   4320
         Left            =   105
         TabIndex        =   22
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   195
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   7620
         _Version        =   393216
         Cols            =   4
         FixedCols       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         MergeCells      =   4
         AllowUserResizing=   1
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   26
         Top             =   4720
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00624E28&
      Height          =   3210
      Left            =   5280
      ScaleHeight     =   3150
      ScaleWidth      =   1020
      TabIndex        =   1
      Top             =   -105
      Width           =   1080
      Begin VB.CommandButton Command4 
         Caption         =   "Pro- ceso"
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
         TabIndex        =   20
         ToolTipText     =   "Lista todas las Facturas Faltantes con Respecto a la Correlatividad."
         Top             =   1125
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   6
         Top             =   2880
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   7
            Top             =   -120
            Width           =   12000
         End
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
         Left            =   105
         Picture         =   "CTRCORREL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   240
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
         Left            =   105
         Picture         =   "CTRCORREL.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Reporte Filtrado"
         Top             =   2010
         Width           =   650
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5205
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Text            =   " "
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command12 
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
         Left            =   2385
         TabIndex        =   15
         Top             =   585
         Width           =   175
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
         Left            =   4665
         TabIndex        =   14
         Top             =   585
         Width           =   175
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2715
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   2715
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2835
         TabIndex        =   17
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   16
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P.venta"
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
         Index           =   2
         Left            =   600
         TabIndex        =   13
         Top             =   2235
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Letra"
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
         Index           =   1
         Left            =   600
         TabIndex        =   12
         Top             =   1755
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         Left            =   600
         TabIndex        =   11
         Top             =   1155
         Width           =   1125
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CTRCORREL.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CTRCORREL.frx":0688
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aceptar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CTRCORREL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub FACTURASEIDCAE(DESDE%, HASTA%, TIPODOCM$, PVTA%, LETRA$, RET_FACTURAX$, RET_IDCAEX$, RET_CTA$)
    FACTURAX$ = ""
    IDCAEX$ = "": RET_CTA$ = ""
    VUELTA% = 0
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE NUMECOMP < 0"
    SQLX$ = SQLX$ + "  AND FECHEMISI >= '" & FETEXCOR1$(DESDE%) & "'"
    SQLX$ = SQLX$ + "  AND FECHEMISI <= '" & FETEXCOR1$(HASTA%) & "'"
    SQLX$ = SQLX$ + "  AND TIPOMOVIM = 'FVEN'"
    SQLX$ = SQLX$ + "  AND LEFT(CODICOM_LAR,3) = '" & TIPODOCM$ & "'" ' FC-, NC- , NC-
    SQLX$ = SQLX$ + "  AND NUMETALO = " & PVTA%
    SQLX$ = SQLX$ + "  AND VARIMOVIM = '" & LETRA$ & "'"
    SQLX$ = SQLX$ + "  order by NUCLIEN"
    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
          FACTURAX$ = SAFETEXT$(!codicom_lar)
          IDCAEX$ = SAFETEXT$(!IDCAE)
          CTA& = XVALO(!NUCLIEN)
          If CTA_ANT& = CTA& Then GoTo 50 'NEXT
          If VUELTA% < 1 Then
            ACUFACTU$ = ACUFACTU$ & FACTURAX$
            ACUIDCAEX$ = ACUIDCAEX$ & IDCAEX$
            ACUCTA$ = ACUCTA$ & SAFETEXT$(CTA&)
            VUELTA% = 1
          Else
            ACUFACTU$ = ACUFACTU$ & "; " & FACTURAX$
            ACUIDCAEX$ = ACUIDCAEX$ & "; " & IDCAEX$
            ACUCTA$ = ACUCTA$ & "; " & SAFETEXT$(CTA&)
          End If
          CTA_ANT& = CTA&
50       .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    RET_FACTURAX$ = ACUFACTU$
    RET_IDCAEX$ = ACUIDCAEX$
    RET_CTA$ = ACUCTA$
    '
    
End Sub

Sub LEO_ARCHIVO_RECUPERA()
    MSF.Rows = 1
    NARCHI1% = FILEOPEN%(LUCOM$ + "FACELEC\RECUPERA.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       CONT% = CONT% + 1
       XX$ = TRIM$(XX$)
       Pos% = InStr(XX$, ">")
       POS2% = InStr(XX$, "</")
       If Pos% > 0 And POS2% > Pos% Then
          TXTLEIDO$ = TRIM$(Mid$(XX$, Pos% + 1, (POS2% - 1) - Pos%))
       End If
       
          DATO$ = ""
          CAMPO$ = ""
       If CONT% = 17 Then
            NUMTCOMPBTE1% = XVALO(TXTLEIDO$)
            Select Case NUMTCOMPBTE1%
              Case 1, 6: TC$ = "FC-"
              Case 2, 7: TC$ = "ND-"
              Case 3, 8: TC$ = "NC-"
            End Select
            If NUMTCOMPBTE1% <= 3 Then LETRA$ = "A"
            If NUMTCOMPBTE1% >= 6 Then LETRA$ = "B"
       End If

          Select Case CONT%
           'Case 4: DATO$ = TRIM$(FORFACPRINT.Label18)          ' Cantidad de registros del detalle de la factura
           Case 4: DATO$ = TXTLEIDO$: CAMPO$ = "ID"       ' ID
           Case 16:
              DATO$ = TXTLEIDO$: CAMPO$ = "CUIT"    '  CUIT
           Case 17:
              DATO$ = LETRA$: CAMPO$ = "LETRA"
           Case 18:
             DATO$ = TXTLEIDO$: CAMPO$ = "PUNTO DE VENTA": PREFIJO$ = FORMATOCON0$(XVALO(TXTLEIDO$), 4)
           Case 19:
             DATO$ = TXTLEIDO$: CAMPO$ = "NUMERO DE COMPROBANTE":  COMPROBANTEX$ = FORMATOCON0$(XVALO(TXTLEIDO$), 8)          'NUMERO DE COMPROBANTE
           Case 21:
             DATO$ = TRIM$(FORMATNUM$(XVALO(TXTLEIDO$), "F12.2")): CAMPO$ = "IMPORTE TOTAL" 'IMPORTE TOTAL
           Case 23:
             DATO$ = TRIM$(FORMATNUM$(XVALO(TXTLEIDO$), "F12.2")): CAMPO$ = "IMPORTE NETO" 'IMPORTE NETO
           Case 24:
             DATO$ = TRIM$(FORMATNUM$(XVALO(TXTLEIDO$), "F12.2")): CAMPO$ = "IVA" 'IMPORTE IMPUESTO
           Case 28:
             DATO$ = TXTLEIDO$: CAMPO$ = "CAE" 'CAE
           Case 29:
             DATO$ = TXTLEIDO$: CAMPO$ = "FECHA COMPROBANTE" 'FECHA COMPROBANTE
          End Select
          If DATO$ <> "" Then
              i& = i& + 1
              MSF.Rows = i& + 1
              MSF.TextMatrix(i&, 1) = CAMPO$
              MSF.TextMatrix(i&, 2) = DATO$
          End If
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
     Loop
    '
    'AGREGO EL N FACTURA COMPLETO
    i& = i& + 1
    MSF.Rows = i& + 1
    MSF.TextMatrix(i&, 1) = "Factura"
    MSF.TextMatrix(i&, 2) = TC$ & PREFIJO$ & "-" & COMPROBANTEX$ & "-" & LETRA$
    Close #NARCHI1%

End Sub

Private Sub aceptar_Click()
   Call Command2_Click
End Sub

Private Sub Command1_Click()
    '
    Unload Me
    '
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text1 = VDEVV$
    End If
    Command12.Enabled = True

End Sub

Private Sub Command2_Click()
   '
   DES% = CVINT(Text1)
   HAS% = CVINT(Text1)
   If DES% < 1 Then Call COMUNI("Falta Fecha Desde"): GoTo 99
   If HAS% < 1 Then Call COMUNI("Falta Fecha Hasta"): GoTo 99
   If TRIM$(Combo1.TEXT) = "" Then Call COMUNI("Falta Tipo de Documento"): GoTo 99
   If TRIM$(Combo2.TEXT) = "" Then Call COMUNI("Falta Letra"): GoTo 99
   If TRIM$(Combo3.TEXT) = "" Then Call COMUNI("Falta Punto de Venta"): GoTo 99
   CAMPOS$ = "Documento/A20;Emisión/D12;Cta/F7;Razon Social/A64 "
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO, , , ANTOT)
   FRMZELECHO.Width = 12500
   '
   SQLX$ = "SELECT codicom_lar,fECHEMISI,nuclien, Refecom FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM = 'FVEN' "
   SQLX$ = SQLX$ + " AND FECHEMISI >= '" & FETEXCOR1$(CVINT(Text1)) & "'"
   SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FETEXCOR1$(CVINT(Text2)) & "'"
   
   SQLX$ = SQLX$ + " AND OPCIMOVIM = '" & TRIM$(Combo1) & "'"
   SQLX$ = SQLX$ + " AND Varimovim = '" & TRIM$(Combo2) & "'"
   SQLX$ = SQLX$ + " AND NumeTalo = '" & TRIM$(Combo3) & "'"
   SQLX$ = SQLX$ + " order by codicom_lar asc"
   Set rs = READSET(SQLX$)
   FRMZELECHO.MSF2.Redraw = False
   J& = 0: i& = 0
   With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.Rows = i& + 1
        numeroComp& = XVALO(Mid$(SAFETEXT$(!codicom_lar), 9, 8))
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = SAFETEXT$(!codicom_lar)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!FECHEMISI)
        FRMZELECHO.MSF2.TextMatrix(i&, 3) = SAFETEXT$(!NUCLIEN)
        FRMZELECHO.MSF2.TextMatrix(i&, 4) = SAFETEXT$(!RefeCom)
        If numeroComp_ant& + 1 <> numeroComp& And i& > 1 Then
           Call COLOREAR_GRILLA_SOLO_UNA_AFILA(FRMZELECHO.MSF2, &HC0E0FF, i&)
           FRMZELECHO.MSF2.TextMatrix(i&, 4) = "**** " + SAFETEXT$(!RefeCom)
           J& = J& + 1
        End If

        numeroComp_ant& = numeroComp&
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.MSF2.Redraw = True
    '
    FRMZELECHO.Caption = "Tabla de Correlatividad (Detectadas: " & J& & " sobre " & i& & ")"

50  FRMZELECHO.Show 1
99
    '
End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
    '

    '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text2 = VDEVV$
    End If
    Command3.Enabled = True

End Sub

Private Sub Command4_Click()
   '
  Command4.Enabled = False
  PVTANORMAL1% = XVALO(Control("FACTURA", "PUVENELE")) 'P VTA. NORMAL
  PVTAFAEXP1% = XVALO(Control("FACTURA", "PUVEFAEX"))  'P.VTA. FACT.EXP.
  PVTABCPTAL1% = XVALO(Control("FACTURA", "PUVENCAP")) 'P.VTA. B.CAPITAL
  If (PVTANORMAL1% + PVTAFAEXP1% + PVTABCPTAL1%) < 1 Then
     Call COMUNI("Falta Configurar los Números de Punto de Venta de Factura Electónica")
     GoTo 99
  End If
  '
  
   DES% = CVINT(Text1)
   HAS% = CVINT(Text2)
   AÑO2017& = Year(Text1)
   If AÑO2017& < "2017" Then
      Call COMUNI("El Sistema Solo Mostrará Información a Partir del Año 2017")
      DES% = CVINT("01/01/17")
   End If
   
   CAMPOS$ = "Documento/A20;Posible Emisión/D18;Cta/F7;Factura (-1)/A18;IDCAE/F16.0;IdCae Posible/F16.0 "
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, CAMPOS$, FRMZELECHO, , , ANTOT)
   FRMZELECHO.Width = 12500
   '
   SQLX$ = "SELECT codicom_lar,fECHEMISI,nuclien, Refecom,VARIMOVIM,NUMETALO FROM CAJABANC WITH(NOLOCK) WHERE TIPOMOVIM = 'FVEN' "
   SQLX$ = SQLX$ + " AND FECHEMISI >= '" & FETEXCOR1$(DES%) & "'"
   SQLX$ = SQLX$ + " AND FECHEMISI <= '" & FETEXCOR1$(HAS%) & "'"
   SQLX$ = SQLX$ + " AND NUMECOMP >= 0"
   SQLX$ = SQLX$ + " AND (LEFT(CODICOM_LAR,3) = 'FC-' OR LEFT(CODICOM_LAR,3) = 'NC-' OR LEFT(CODICOM_LAR,3) = 'ND-')"
   If PVTANORMAL1% > 0 Then SQLX$ = SQLX$ + " AND (NUMETALO = " & PVTANORMAL1%
   If PVTAFAEXP1% > 0 Then SQLX$ = SQLX$ + " OR NUMETALO = " & PVTAFAEXP1%
   If PVTABCPTAL1% > 0 Then SQLX$ = SQLX$ + " OR NUMETALO = " & PVTABCPTAL1%
   SQLX$ = SQLX$ + ")"
   SQLX$ = SQLX$ + " order by varimovim,codicom_lar,numetalo,FECHEMISI"
   Set rs = READSET(SQLX$)
   FRMZELECHO.MSF2.Redraw = False
   J& = 0: i& = 0: L& = 0
   With rs
      Do While Not .EOF
        'LECTURA
        numeroComp& = XVALO(Mid$(SAFETEXT$(!codicom_lar), 9, 8))
        LETRA$ = SAFETEXT$(!VARIMOVIM)
        PUNTOVTA% = XVALO(!NUMETALO)
        TIPODEDOCUM$ = Left(SAFETEXT$(!codicom_lar), 3)
        FACTU$ = SAFETEXT$(!codicom_lar)
        FECHEMISION% = CVINT(!FECHEMISI)
        L& = L& + 1
        If numeroComp_ant& = numeroComp& Then GoTo 40  'por si hay facturas repetidas
30      If numeroComp_ant& + 1 <> numeroComp& And numeroComp_ant& > 0 Then 'CONTROLO LA CORRELATIVIDAD
           'SI ES MISMA LETRA-PUNTO DE VENTA Y TIPO DE DOCUMENTO Y SE ROMPIO LA CORRELATIVIDAD LO ESCRIBO
           If LETRA$ = LETRA_ANT$ And PUNTOVTA% = PUNTOVTA_ANT% And TIPODEDOCUM_ANT$ = TIPODEDOCUM$ Then
'            i& = i& + 1
'            FRMZELECHO.msf2.Rows = i& + 1
            numeroComp& = XVALO(Mid$(SAFETEXT$(!codicom_lar), 9, 8))
            numeroComp_ant& = numeroComp_ant& + 1
            NfactFaltante$ = Left$(SAFETEXT$(!codicom_lar), 8) & FORMATOCON0$(numeroComp_ant&, 8) & "-" & LETRA$
            Call FACTURASEIDCAE(FECHEMISION_ANT%, FECHEMISION%, TIPODEDOCUM$, PUNTOVTA%, LETRA$, RET_FACTURAX$, RET_IDCAEX$, RET_CTA$)
            Call IDCAES_POSIBLES(FECHEMISION_ANT%, FECHEMISION%, RET_IDCAEFALT$)
            On Error Resume Next
            strArray1 = Split(RET_FACTURAX$, ";")
            strArray2 = Split(RET_IDCAEX$, ";")
            strArray3 = Split(RET_CTA$, ";")
            strArray4 = Split(RET_IDCAEFALT$, ";")
            
            CANTFAC& = UBound(strArray1)
'            CANTIDPOSIBLES& = UBound(strArray4)
            
            CANTIX& = CANTFAC&
            If CANTFAC& < 1 Then CANTIX& = 0

            For R& = 0 To CANTIX&
                    i& = i& + 1
                    FACTURAX$ = strArray1(R&)
                    FRMZELECHO.MSF2.Rows = i& + 1
                    IDCAE$ = strArray2(R&)
                    CTA& = XVALO(strArray3(R&))
'                        IDCAEPOSIBLES$ = strArray4(R&)
                    FRMZELECHO.MSF2.TextMatrix(i&, 1) = NfactFaltante$
                      STRINGRANGO$ = FECHATEX$(FECHEMISION_ANT%) & "al " & FECHATEX$(FECHEMISION%)
                    FRMZELECHO.MSF2.TextMatrix(i&, 2) = STRINGRANGO$
                    FRMZELECHO.MSF2.TextMatrix(i&, 3) = strArray3(i&)
                    FRMZELECHO.MSF2.TextMatrix(i&, 4) = strArray1(i&)
                    FRMZELECHO.MSF2.TextMatrix(i&, 5) = strArray2(i&)
             Next R&
             CANTIDPOSIBLES& = UBound(strArray4)
             If CANTIDPOSIBLES& < 0 Then CANTIDPOSIBLES& = 0
             For R& = 0 To CANTIDPOSIBLES&
                    If strArray4(R&) <> "" Then
                        
                        FRMZELECHO.MSF2.Rows = i& + 1
                        FRMZELECHO.MSF2.TextMatrix(i&, 1) = NfactFaltante$
                          STRINGRANGO$ = FECHATEX$(FECHEMISION_ANT%) & "al " & FECHATEX$(FECHEMISION%)
                        FRMZELECHO.MSF2.TextMatrix(i&, 2) = STRINGRANGO$
                        FRMZELECHO.MSF2.TextMatrix(i&, 3) = 0
                        FRMZELECHO.MSF2.TextMatrix(i&, 4) = ""
                        FRMZELECHO.MSF2.TextMatrix(i&, 5) = ""
                        FRMZELECHO.MSF2.TextMatrix(i&, 6) = strArray4(R&)
                        If i& > FRMZELECHO.MSF2.Rows Then
                          i& = i& + 1
                          FRMZELECHO.MSF2.Rows = i& + 1
                        End If
                    End If
             Next R&
                    Me.Caption = L&
            On Error GoTo 0
'            FRMZELECHO.msf2.TextMatrix(i&, 1) = NfactFaltante$
'            FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!FECHEMISI)
            
            LETRA$ = SAFETEXT$(!VARIMOVIM)
            PUNTOVTA% = XVALO(!NUMETALO)
            TIPODEDOCUM$ = Left(SAFETEXT$(!codicom_lar), 3)
            J& = J& + 1
            If numeroComp_ant& < numeroComp& Then GoTo 30 ' VUELVO A CONTROLAR POR SI EL HUECO ES DE MAS DE UN DOCUMENTO
           End If
        End If
        'GUARDO ULTIMOS DATOS LEIDOS
        numeroComp_ant& = numeroComp&
        LETRA_ANT$ = SAFETEXT$(!VARIMOVIM)
        PUNTOVTA_ANT% = XVALO(!NUMETALO)
        TIPODEDOCUM_ANT$ = TIPODEDOCUM$
        FECHEMISION_ANT% = CVINT(!FECHEMISI)
 
40      .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.MSF2.Redraw = True
    '
45  FRMZELECHO.Caption = "Tabla de Correlatividad (Detectadas: " & J& & " Facturas No Registradas sobre " & L& & " Facturas Registradas)"
    
50  FRMZELECHO.Show 1
    'LEO LA FACTURA
    
    FACTURAX$ = RESP_ZELE_VECT(1)
    PVTA% = XVALO(Mid(FACTURAX$, 4, 4))
    If PVTA% < 1 Then
       Call COMUNI("No Se Puede Identificar el Punto de Venta")
       GoTo 99
    End If
    '
    If PVTANORMAL1% = PVTA% Then 'FACTURA COMUN A O B PERO LA BASICA
        IDCAE = XVALO(RESP_ZELE_VECT(5)) 'TOMA EL VALOR DE LA FACTURA EN -1
        If IDCAE < 1 Then 'SI ESTA NO DEVUELVE UN IDCAE
         IDCAE = XVALO(RESP_ZELE_VECT(6)) ' TOMA EL IDCAE POSIBLE (QUE VIENE DE UN HUECO QUE HAYA EN IDCAE)
        End If
        If IDCAE < 1 Then
           Call COMUNI("IDCae Inválido")
           GoTo 99
        End If
        IDENCAE$ = SAFETEXT$(IDCAE)
        If EXISTE%(LUCOM$ + "FACELEC\CorrelatividadDeFactura.EXE") = 1 Then
           XXX = Shell(LUCOM$ + "FACELEC\CorrelatividadDeFactura.EXE " & IDENCAE$ & " " & FACTURAX$ & ", 4")
           Call LEO_ARCHIVO_RECUPERA
           'AGREGO EL N DE CLIENTE SEGUN EL CUIT.
           NC = CtaCliSegunCuit(MSF.TextMatrix(2, 2))
           TXTCLIENTE.TEXT = SAFETEXT$(NC) & " - " & rasocli$(NC)
        End If
        
    ElseIf PVTABCPTAL1% = PVTA% Then 'FACTURA DE BIENES DE CAPITALES
    
    
    ElseIf PVTAFAEXP1% = PVTA% Then  'FACTURA DE EXPORTACION
    
    
    End If
    
    
99  Me.Caption = ""
    Command4.Enabled = True
    '

End Sub
Sub IDCAES_POSIBLES(DESDE%, HASTA%, RET_IDCAEFALT$)
    FACTURAX$ = ""
    IDCAEX$ = ""
    VUELTA% = 0: ACUMUFALTAID$ = ""
    SQLX$ = "SELECT IDCAE FROM CAJABANC WITH(NOLOCK)  "
    SQLX$ = SQLX$ + "  WHERE FECHEMISI >= '" & FETEXCOR1$(DESDE%) & "'"
    SQLX$ = SQLX$ + "  AND FECHEMISI <= '" & FETEXCOR1$(HASTA%) & "'"
    SQLX$ = SQLX$ + "  AND IDCAENUM > 0 "
    SQLX$ = SQLX$ + "  GROUP BY IDCAE "
    SQLX$ = SQLX$ + "  ORDER BY IDCAE"
    
    Set rs = READSET(SQLX$)
    With rs
       Do While Not .EOF
          NUMIDCAE = XVALO(!IDCAE)
30        If NUMIDCAE_ANT > 0 And NUMIDCAE_ANT + 1 < NUMIDCAE Then
            NUMIDCAE_ANT = NUMIDCAE_ANT + 1
            If VUELTA% < 1 Then
              ACUMUFALTAID$ = ACUMUFALTAID$ & SAFETEXT$(NUMIDCAE_ANT)
              VUELTA% = 1
            Else
              ACUMUFALTAID$ = ACUMUFALTAID$ & "; " & SAFETEXT$(NUMIDCAE_ANT)
            End If
            If NUMIDCAE_ANT < NUMIDCAE Then GoTo 30
            
          End If
          NUMIDCAE_ANT = XVALO(!IDCAE)
         .MoveNext
       Loop
    End With
    rs.Close
    Set rs = Nothing
    RET_IDCAEFALT$ = ACUMUFALTAID$
End Sub
    
Private Sub Command5_Click()
    If CLACONTRA% > 0 Then
       Command5.Visible = False
    End If
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If MSF.Rows > 2 Then
      RECUEC07.Text1.TEXT = CtaCliSegunCuit(MSF.TextMatrix(2, 2))
      RECUEC07.Show 1
    End If
    Command6.Enabled = True

End Sub

Private Sub Form_Load()
  Condi$ = "Tipomovim = 'FVEN' Group By Opcimovim Order By Opcimovim asc"
  Call CargaCombo(Combo1, "Cajabanc", "Opcimovim", Condi$)
  
  Condi$ = "Tipomovim = 'FVEN' Group By Varimovim Order By Varimovim asc"
  Call CargaCombo(Combo2, "Cajabanc", "Varimovim", Condi$)
'  Condi$ = "Tipomovim = 'FVEN' Group By NumeTalo Order By NumeTalo asc"
'  Call CargaCombo(Combo3, "Cajabanc", "NumeTalo", Condi$)

  PVTANORMAL$ = TRIM$(Control("FACTURA", "PUVENELE")) 'P VTA. NORMAL
  PVTABCPTAL$ = TRIM$(Control("FACTURA", "PUVENCAP")) 'P.VTA. B.CAPITAL
  PVTAFAEXP$ = TRIM$(Control("FACTURA", "PUVEFAEX"))  'P.VTA. FACT.EXP.
  If PVTANORMAL$ <> "" Then Combo3.AddItem PVTANORMAL$
  If PVTAFAEXP$ <> "" Then Combo3.AddItem PVTAFAEXP$
  If PVTABCPTAL$ <> "" Then Combo3.AddItem PVTABCPTAL$
  X1$ = HOY(HO$)
  Text1.TEXT = "01/" & Mid$(HO$, 4)
  Text2.TEXT = FECHATEX$(HOY(HO$))
  CAMPOS$ = "Identificador/A24;Valor/A24"
  Call CARGA_ENCABEZADO(MSF, CAMPOS$, Me)

  
End Sub
Sub CargaCombo(CMB As ComboBox, Tablax$, Campox$, Condi$)
    SQLX$ = "Select " & Campox$ & " as Campo1"
    SQLX$ = SQLX$ + " from " & Tablax$
    SQLX$ = SQLX$ + " with(nolock)"
    SQLX$ = SQLX$ + " Where " & Condi$
    Set rs = READSET(SQLX$)
    CMB.Clear
    With rs
        Do While Not .EOF
           CMB.AddItem SAFETEXT$(!CAMPO1)
           .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub
Private Sub Picture1_Click()
    '
    '
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text10_DblClick()
   '\\\OT-06-0570-RG-27/12/06///
   '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text10_GotFocus()

End Sub

Private Sub Text8_Change()

End Sub

Private Sub Text8_GotFocus()

End Sub

Private Sub Text9_Click()
 
End Sub

