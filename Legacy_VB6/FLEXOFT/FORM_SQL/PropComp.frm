VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form PropComp 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7725
   ClientLeft      =   1905
   ClientTop       =   3300
   ClientWidth     =   16590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   136.261
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   292.629
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5280
      Top             =   3600
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   3
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Propuesta de Compras"
      Top             =   0
      Width           =   10695
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   8400
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3240
      OleObjectBlob   =   "PropComp.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "PropComp.frx":0234
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   7245
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   12779
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   4200
      TabIndex        =   9
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Decripción:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   405
      Width           =   900
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   1200
      TabIndex        =   7
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D8BD05&
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   1320
      TabIndex        =   5
      Top             =   45
      Width           =   1455
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "PropComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ENTER EN CUALQUIER LUGAR DEL FORMULARIO Y REALIZA CLICK EN EL BOTON
'SE DEBE MARCAR EN TRUE LA PROPIEDAD DEFAULT DEL BOTON


Private Sub cmdcancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
 Campos$ = "CI/A0"                       '  1
 Campos$ = Campos$ + ";Código/A16"                   '  2
 Campos$ = Campos$ + ";Descripción/A32"              '  3
 Campos$ = Campos$ + ";U -m/ A3"                    '  4
 Campos$ = Campos$ + ";Necesidad/F10.4"              '  5
 Campos$ = Campos$ + ";M/A4"                         '  6
 Campos$ = Campos$ + ";Pr.Unit/F10.4"                '  7
 Campos$ = Campos$ + ";Importe/F12.4"                '  8
 Campos$ = Campos$ + ";Cant-Indiv/F10.4"             '  9
 Campos$ = Campos$ + ";Importe/F12.4"                '  10
 Campos$ = Campos$ + ";Stock/F9.4"                   '  11
 Campos$ = Campos$ + ";Cant-Stock//F10.4"            '  12
 Campos$ = Campos$ + ";Importe/F12.4"                '  13
 Campos$ = Campos$ + ";Cta/f8"                       '  14
 Campos$ = Campos$ + ";R.Social/A24"                 '  15
 Campos$ = Campos$ + ";Cant.Compra/F10.4"                 '  16
 Campos$ = Campos$ + ";Imp.Compra/F12.4"                 '  17
 

Call CARGA_ENCABEZADO(Me.MSF1, Campos$, MUESTRADATOS3, 1, 1)
End Sub
Private Sub Form_Resize()
   On Error Resume Next
   Me.MSF1.Width = Me.Width - 300
   Me.MSF1.Height = Me.Height - 1600
   Text17(3) = Me.Width - 5000
   On Error GoTo 0
End Sub

Sub Label8_Change()
    CI1% = CODINT%(Label8)
    If CI1% < 1 Then Exit Sub
    Label2 = DESCRIT0$(CI1%)
    Call CargaPropuestaCompras(CI1%)

End Sub

Sub CargaPropuestaCompras(CI1%)
   '
   If CI1% < 1 Then Exit Sub
   '
   CANTIACALCULAR = XVALO(PROPUCOM.Label5)
   
   CODD$ = CODEXT$(CI1%)
   Screen.MousePointer = 11
   PORMILLAR% = 0
   Call LEEEXPLO(CODD$, 1)
   '
'   List1.Clear
'   Call SETULTREG("!BOLREDET", 0)
   For i% = 1 To CAIT%
     ZZ1$ = EXPLOLIN$(i%)
     CICOMPO% = CODINT%(TRIM$(Left$(ZZ1$, 15)))
     Codigo$ = TRIM$(Left$(ZZ1$, 15))
     DESCRIX$ = TRIM$(Mid$(ZZ1$, 17, 28))
     UNMED$ = Unimed$(CICOMPO%)
     CANTIUNI# = XVALO(Mid$(ZZ1$, 49, 12))
     CANTICALCULADA# = CANTIUNI# * CANTIACALCULAR
     MONECO% = MONECOSTO%(CICOMPO%)
     SIMBOLOMONEDA$ = SIMBOLPES$(MONECO%)
     COSTOUNITARIO = FORMATNUM$(COSUNI(CICOMPO%), "F9.4")
     IMPORTESEGUNCANTIDADINGRESADA = COSPESOS(CICOMPO%) * CANTICALCULADA#
     CANTISEGUNLOTEINDIVISIBLE# = REDON_INDIVIS(CICOMPO%, CANTICALCULADA#) 'COMPRA SIN CONSIDERAR STOCK
     XX = COSPESOS(CICOMPO%)
     IMPORTESEGUNCANTIDALOTEINDIVISIBLE# = COSPESOS(CICOMPO%) * CANTISEGUNLOTEINDIVISIBLE#
     STOCKDISPONIBLE = STOCKDIS(CICOMPO%)
     DifEntreStockYNecesidad# = 0
     If STOCKDISPONIBLE > 0 Then
        DifEntreStockYNecesidad# = CANTICALCULADA# - STOCKDISPONIBLE
     Else
        DifEntreStockYNecesidad# = CANTICALCULADA#
        STOCKDISPONIBLE = 0
     End If
     
     DeberiaComprarConsiderandoElStock = 0: ImporteConsiderandoStockYLoteIndivis# = 0
'     If DifEntreStockYNecesidad# > 0 Then
        DeberiaComprarConsiderandoElStock = REDON_INDIVIS(CICOMPO%, DifEntreStockYNecesidad#)
        ImporteConsiderandoStockYLoteIndivis# = COSPESOS(CICOMPO%) * DeberiaComprarConsiderandoElStock
'     End If
     CuentaProv = PROVHABI(CICOMPO%)
     razonSocial$ = rasocli$((-1) * CuentaProv)
     MSF1.Rows = i% + 1
     MSF1.TextMatrix(i%, 1) = CICOMPO%
     MSF1.TextMatrix(i%, 2) = Codigo$ 'CODIGO
     MSF1.TextMatrix(i%, 3) = DESCRIX$
     MSF1.TextMatrix(i%, 4) = UNMED$
     MSF1.TextMatrix(i%, 5) = FORMATNUM$(CANTICALCULADA#, "F12.4")
     MSF1.TextMatrix(i%, 6) = SIMBOLOMONEDA$
     MSF1.TextMatrix(i%, 7) = FORMATNUM$(COSPESOS(CICOMPO%), "F12.4")
     MSF1.TextMatrix(i%, 8) = FORMATNUM$(IMPORTESEGUNCANTIDADINGRESADA, "F12.4")
     MSF1.TextMatrix(i%, 9) = FORMATNUM$(CANTISEGUNLOTEINDIVISIBLE#, "F12,4")
     MSF1.TextMatrix(i%, 10) = FORMATNUM$(IMPORTESEGUNCANTIDALOTEINDIVISIBLE#, "F12,4")
     MSF1.TextMatrix(i%, 11) = FORMATNUM$(STOCKDISPONIBLE, "F12,4")
     MSF1.TextMatrix(i%, 12) = FORMATNUM$(DeberiaComprarConsiderandoElStock, "F12,4")
     MSF1.TextMatrix(i%, 13) = FORMATNUM$(ImporteConsiderandoStockYLoteIndivis#, "F12,4")
     MSF1.TextMatrix(i%, 14) = CuentaProv
     MSF1.TextMatrix(i%, 15) = razonSocial$

   Next i%
   '
   Screen.MousePointer = 1
   PROPUCOM.Show 1
   '


End Sub


Private Sub MSF1_DblClick()
 Campos$ = "CI/A0"                       '  1
 Campos$ = Campos$ + ";Código/A16"                   '  2
 Campos$ = Campos$ + ";Descripción/A32"              '  3
 Campos$ = Campos$ + ";U -m / A3"                    '  4
 Campos$ = Campos$ + ";Necesidad/F12.4"              '  5
 Campos$ = Campos$ + ";M/A4"                         '  6
 Campos$ = Campos$ + ";Pr.Unit/F12.4"                '  7
 Campos$ = Campos$ + ";Importe/F12.4"                '  8
 Campos$ = Campos$ + ";Cant-Indiv/F12.4"             '  9
 Campos$ = Campos$ + ";Importe/F12.4"                '  10
 Campos$ = Campos$ + ";Stock/F9.4"                   '  11
 Campos$ = Campos$ + ";Cant-Stock//F12.4"            '  12
 Campos$ = Campos$ + ";Importe/F12.4"                '  13
 Campos$ = Campos$ + ";Cta/f8"                       '  14
 Campos$ = Campos$ + ";R.Social/A32"                 '  15
 Campos$ = Campos$ + ";Cant.Compra/F12.4"                 '  16
 Campos$ = Campos$ + ";Imp.Compra/F12.4"                 '  17
    Set VENTANA = New VentaNew

    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    'VENTANA.formula(15) = "=EXCEL;(L1/K1)*(L1/K1)*M1"
    'VENTANA.formula(14) = "=EXCEL;(12/11)*(12/11)*13"
    VENTANA.Formula(17) = "7;*;16"
    VENTANA.HabilitaBorrar = 1
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Propuesta de Compra"

    VENTANA.ColEcoCampo(14) = "nume_cli;Raso_Cli;PROVED12"
    VENTANA.CampEditables = "14;16"
    VENTANA.Inicializar = 1
    '
    Call CopiarMsfAOtroMsf(MSF1, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
30  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
    End If
     Unload VENTABNEW
End Sub


