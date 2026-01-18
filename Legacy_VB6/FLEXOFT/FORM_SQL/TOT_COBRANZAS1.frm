VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form TOT_COBRANZAS1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Total Cobranzas"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exp.Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   8280
      TabIndex        =   13
      ToolTipText     =   "Exportar a Excel"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1150
      Left            =   0
      Picture         =   "TOT_COBRANZAS1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "iMPRIME REPORTE ACTIVO"
      Top             =   3840
      Width           =   8175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6120
      OleObjectBlob   =   "TOT_COBRANZAS1.frx":5C12
      Top             =   720
   End
   Begin VB.TextBox Text15 
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
      Index           =   8
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Cuotas:"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "Total:"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text16 
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Matrículas:"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1095
   End
   Begin VB.TextBox Text15 
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "TOTALES DE COBRANZAS X RANGO DE FECHAS"
      Top             =   0
      Width           =   12675
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   6960
      TabIndex        =   2
      Top             =   620
      Width           =   5655
      Begin MSComCtl2.DTPicker DESDE 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   130351105
         CurrentDate     =   40875
      End
      Begin MSComCtl2.DTPicker HASTA 
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CalendarBackColor=   14737632
         CalendarTitleBackColor=   14737632
         Format          =   130351105
         CurrentDate     =   40875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   14737632
      BackColorFixed  =   16501405
      ForeColorSel    =   16501405
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "TOT_COBRANZAS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command25_Click()
   Command25.Enabled = False
   'valido
   If MSF.Rows <= 1 Then
      MsgBox "GRILLA DE IMPRESIÓN VACIA"
      GoTo 99
   End If
   '
   Set GRILLA_IMPRESION = MSF
   '
   PREIMPRE.txtEncabezado2.TEXT = "Desde: " & DESDE.Value & " Hasta: " & HASTA.Value

   '
   
   ENCABEZADO1$ = "Totales de Cobranzas"
   
   PREIMPRE.txtEncabezado1.TEXT = ENCABEZADO1$
   PREIMPRE.txtPiePagina = "Fin de Reporte"
   PREIMPRE.Combo1.TEXT = "Vertical"
   PREIMPRE.Show 1
   '
99 Command25.Enabled = True

End Sub

 Private Sub Command3_Click()
   Command3.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
   ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(TOT_COBRANZAS.MSF, ARCHIEX$)
   End If
   
99 Command3.Enabled = True

End Sub


Private Sub DESDE_Change()
   Call CARGA_GRILLA_COB_TOT
End Sub

Private Sub DESDE_Click()
   Call CARGA_GRILLA_COB_TOT
End Sub

Private Sub Form_Load()

    Call APLICAR_SKIN(Me, App.Path & "\SKINS\SteelBlue.skn")
    Call ENCABEZA_GRILLA
    DESDE.Value = Date
    HASTA.Value = Date
    Call CARGA_GRILLA_COB_TOT
    Call CARGAVECTOR_MES
End Sub


Sub CARGA_GRILLA_COB_TOT()
   '
   MSF.Rows = 1
   MSF.Cols = 4
   '
   DES% = CVINT(DESDE.Value)
   HAS% = CVINT(HASTA.Value)
   DESDEX = FETEXCOR1$(DES%)
   HASTAX = FETEXCOR1$(HAS%)
   FILA& = 0

10 CONDI$ = " FECHEMISI >= '" & DESDEX & "' AND FECHEMISI <= '" & HASTAX & "'"
   '
   If CantRegistros&("CAJABANC_COL", CONDI$, "NOMESS") < 1 Then Exit Sub
   '
   'OBTENGO LOS IMPORTES ACUMULADOS POR COBRANZAS DE MATRICULAS Y CUOTAS
   CONDI1$ = CONDI$ + "  AND (VARIMOVI = 'COBMATRI' OR VARIMOVI = 'MATRICULA') AND CUENTA_CONT  = 1"
   CONDI2$ = CONDI$ + "  AND (VARIMOVI = 'RCOB' OR VARIMOVI = 'RECIBO COBRANZA') AND CUENTA_CONT  = 1"
   SUINGRE_MAT# = XVALO(VALOADMIN("CAJABANC_COL", "SUM(HABER)", CONDI1$))
   SUINGRE_RCOB# = XVALO(VALOADMIN("CAJABANC_COL", "SUM(HABER)", CONDI2$))
   '
   'IMPRIMO EN LA GRILLA
   FILA& = FILA& + 1
   MSF.Rows = FILA& + 1
   MSF.TextMatrix(FILA&, 1) = FORMATNUM$(SUINGRE_MAT#, "F14.2")
   MSF.TextMatrix(FILA&, 2) = FORMATNUM$(SUINGRE_RCOB#, "F14.2")
   MSF.TextMatrix(FILA&, 3) = FORMATNUM$((SUINGRE_RCOB# + SUINGRE_MAT#), "F14.2")
   '
   'IMPRIMO EN CASILLEROS A PIE DE FRM
   Text15(0) = FORMATNUM$(SUINGRE_MAT#, "F14.2")
   Text15(1) = FORMATNUM$(SUINGRE_RCOB#, "F14.2")
   Text15(2) = FORMATNUM$((SUINGRE_RCOB# + SUINGRE_MAT#), "F14.2")
   '
End Sub
Sub ENCABEZA_GRILLA()
    MSF.Cols = 4
    NUCOLU% = 1: TEPRUEBA = "1": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
    NUCOLU% = 2: TEPRUEBA = "COBRANZA X MATRICULAS000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "COBRANZA X MATRICULAS"
    NUCOLU% = 3: TEPRUEBA = "COBRANZA X CUOTAS000S000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "COBRANZA X CUOTAS"
    NUCOLU% = 4: TEPRUEBA = "TOTALCOBRANZA00000000000 ": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "TOTAL COBRANZAS"
    '
End Sub

Private Sub HASTA_Change()
    Call CARGA_GRILLA_COB_TOT
End Sub

Private Sub HASTA_Click()
   Call CARGA_GRILLA_COB_TOT
End Sub



