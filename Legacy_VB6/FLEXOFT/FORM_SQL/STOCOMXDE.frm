VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form STOCOMXDE 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Stocks de Componentes"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "VISTA DETALLE"
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
      Left            =   11760
      TabIndex        =   22
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   9000
      Width           =   2130
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ACUMULAR X CÓDIGO"
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
      Left            =   13920
      TabIndex        =   21
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   9000
      Width           =   2130
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Seleccionar Depósitos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   10440
      TabIndex        =   15
      ToolTipText     =   "Selección de Depósitos para visualizar su stock actual"
      Top             =   300
      Width           =   1170
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15930
      Begin VB.TextBox TXTNIVELES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6960
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exportar a Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1440
         TabIndex        =   23
         Top             =   1050
         Width           =   2130
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9120
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   315
         Width           =   3375
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   11640
         TabIndex        =   16
         Top             =   160
         Width           =   4095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Niveles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   25
         Top             =   1120
         Width           =   780
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Simular      Cantidad     (x defecto 1)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   8040
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -30
         TabIndex        =   8
         Top             =   810
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -10
         TabIndex        =   7
         Top             =   400
         Width           =   1380
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stocks"
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
         Left            =   3435
         TabIndex        =   6
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1470
         TabIndex        =   5
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4410
         TabIndex        =   4
         Top             =   315
         Width           =   1770
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1470
         TabIndex        =   3
         Top             =   720
         Width           =   6600
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   9330
      Left            =   16200
      ScaleHeight     =   9300
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   11
         Top             =   7200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "STOCOMXDE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   210
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "STOCOMXDE.frx":014A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STOCOMXDE.frx":037E
      TabIndex        =   9
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   7155
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   12621
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
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
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   7275
      Left            =   120
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   15945
      _ExtentX        =   28125
      _ExtentY        =   12832
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421504
      BackColorSel    =   4210752
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
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
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "STOCOMXDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CamposOriginal$

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   SELDEPOS.Show 1
   STRINGDEPOS$ = ""
   List1.Clear
   
   For U& = 1 To SELDEPOS.LISNUM.ListCount
       If SELDEPOS.LISNUM.Selected(U& - 1) = True Then
        TXTX$ = SELDEPOS.LISNUM.List(U& - 1)
        List1.AddItem TXTX$
        NDEP$ = TRIM$(Str$(Mid$(TXTX$, 1, 6)))
        STRINGDEPOS$ = STRINGDEPOS$ + NDEP$ + ","
      End If
     Next U&
     If TRIM$(STRINGDEPOS$) <> "" Then
        Text1 = Mid$(STRINGDEPOS$, 1, Len(STRINGDEPOS$) - 1) 'para quitar la coma
     Else
        Text1 = ""
     End If

   Command15.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Text3.Enabled = False
   Command15.Enabled = False
   List1.Enabled = False
   msf2.Rows = 1
   Dim VectCodigoAcumulaCantSimulada()
   CantInicial% = 0
   ReDim Preserve VectCodigoAcumulaCantSimulada(CantInicial%)
   '
   'ACA QUEDARIAN CARGADOS SOLO UNA VEZ CADA CODIGO Y APARTE ACUMULO EN UN VECTOR LA CANTIDAD ACUMULADA DE LA COL 6
   For i& = 1 To MSF.Rows - 1 'RECORRO
    CI1% = XVALO(MSF.TextMatrix(i&, 1))
        If CI1% >= UBound(VectCodigoAcumulaCantSimulada) Then
           ReDim Preserve VectCodigoAcumulaCantSimulada(CI1%)
        End If
       If VectCodigoAcumulaCantSimulada(CI1%) > 0 Then 'ES QUE A ESTE CODIGO  YA LO LEYO
          'ACUMULA LA CANTIDAD SIMULADA
          VectCodigoAcumulaCantSimulada(CI1%) = VectCodigoAcumulaCantSimulada(CI1%) + XVALO(MSF.TextMatrix(i&, 6))
       Else
          REG& = REG& + 1
          msf2.Rows = REG&
          VectCodigoAcumulaCantSimulada(CI1%) = XVALO(MSF.TextMatrix(i&, 6))
          Call AGREGAR_FILA_OTRO_MSF(MSF, msf2, REG&)
       End If
       
   
   Next i&
   '
   'ACA APLICO LA CANTIDAD ACUMULADA
   For i& = 1 To msf2.Rows - 1 'RECORRO
       CI1% = XVALO(msf2.TextMatrix(i&, 1))
      msf2.TextMatrix(i&, 6) = FORMATNUM$(VectCodigoAcumulaCantSimulada(CI1%), "F12.3") 'LE APLICA EN LA CANTIDAD SIMULADA EL ACUMULADO QUE CARGO EN EL FOR ANTERIOR
   Next i&
   MSF.Visible = False
   MSF.ZOrder 1
   msf2.Visible = True
   msf2.ZOrder 0
   
99 Command3.Enabled = True

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Command2.Enabled = True
   Text3.Enabled = True
   Command15.Enabled = True
   List1.Enabled = True
   MSF.Visible = True: MSF.ZOrder 0
   Text1.Enabled = True
End Sub

Private Sub Command4_Click()
    Dim MSF As Object
    Set MSF = Me.MSF

    If msf2.Visible = True Then Set MSF = Me.msf2
   
   Call ImprimirEnExcel(MSF)
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture3)
    CamposOriginal$ = "CI1/Aa0"                              '1
    CamposOriginal$ = CamposOriginal$ + ";CodXEelem/A15"    '2
    CamposOriginal$ = CamposOriginal$ + ";Descripción/A32"  '3
    CamposOriginal$ = CamposOriginal$ + ";U.M./A3"           '4
    CamposOriginal$ = CamposOriginal$ + ";Uso/F9.4"        '5
    CamposOriginal$ = CamposOriginal$ + ";Cant.Simulada/F10.4" '6
    CamposOriginal$ = CamposOriginal$ + ";St.Acumulado./F12.4"   '7
    CamposOriginal$ = CamposOriginal$ + ";LIBRE/A0"   '8
    '
End Sub

Private Sub Label4_Change()
   '
   'DESCRIPCION DE ARTICULO SELECCIONADO
   'DATOS DE ENCABEZADO
   CI1% = CODINT%(Label4)
   Label6 = DESCRIT$(CI1%)
   CODDX$ = Label4
   Label5 = FORMATNUM$(STOCKACT(CI1%), "F15." & CNTDC$(CI1%) & "")
   If CI1% < 1 Then Exit Sub
   
   Call CARGA_ENCABEZADO(MSF, CamposOriginal$, STOCOMXDE)
   Call CARGA_ENCABEZADO(msf2, CamposOriginal$, STOCOMXDE)
   'If XVALO(TXTNIVELES) < 1 Then TXTNIVELES = "TODOS"
   'Call CargarGrilla(CI1%)
   Call CargaEstructuraMultinivel(Me.MSF, CI1%, XVALO(TXTNIVELES), 1)
   
End Sub
Sub CargarGrilla(CI1%)
   '
   Screen.MousePointer = 11
'   Call EXPLOMUL(CI1%)
'   If CONTAMUL% < 2 Then Exit Sub
   '
   cantidadSimulada = XVALO(Text3)
   If cantidadSimulada < 1 Then cantidadSimulada = 1
   COEFCA = 1: If PORMILLAR% = 1 Then COEFCA = 1000
'   Call CargaEstructuraMultinivel(MSF, CI1%, XVALO(TXTNIVELES), cantidadSimulada)
'   Fin& = CONTAMUL
'   For KKI% = 2 To CONTAMUL
'     Z$ = Space$(64)
'     RESUL = NIVE%(KKI%)
'     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
'     SP1$ = String$(RESUL, ">")
'     MSF.Rows = KKI%
'     MSF.TextMatrix(KKI% - 1, 1) = CIK%(KKI%)
'     MSF.TextMatrix(KKI% - 1, 2) = CODEXT$(CIK%(KKI%))
'     MSF.TextMatrix(KKI% - 1, 3) = SP$ & SP$ & DESCRIT0$(CIK%(KKI%))
'     MSF.TextMatrix(KKI% - 1, 4) = Unimed$(CIK%(KKI%))
'     MSF.TextMatrix(KKI% - 1, 5) = ROUND(COEFCA * CANT(KKI%), 4)
'     MSF.TextMatrix(KKI% - 1, 6) = ROUND(COEFCA * CANT(KKI%), 4) * Int(cantidadSimulada)
'     'MSF.TextMatrix(KKI% - 1, 7) = FORMATNUM$(STOCKACT(CIK%(KKI%)), "F12.1")
'     MSF.TextMatrix(KKI% - 1, 7) = ""
'   Next KKI%
   Screen.MousePointer = 1
End Sub

    '

Private Sub List1_Change()

End Sub

Private Sub Text1_Change()
  Call ActualizarInfo
  
End Sub
Sub ActualizarInfo()
   Campos$ = CamposOriginal$
   Erase CADENATEX$
   If TRIM$(Text1) <> "" Then
      Call TEXTOLOTES(Text1, ",")
        For i& = 1 To UBound(CADENATEX$)
           Campos$ = Campos$ + ";Dep." & CADENATEX$(i&) & "/A12"
        Next i&
   End If
   Call CARGA_ENCABEZADO(MSF, Campos$, STOCOMXDE)
   Call CARGA_ENCABEZADO(msf2, Campos$, STOCOMXDE)
   CI1% = CODINT%(Label4)
   indDepo% = 0
'   Call CargarGrilla(CI1%)
   Call CargaEstructuraMultinivel(Me.MSF, CI1%, XVALO(TXTNIVELES), XVALO(Text3))
   If TRIM$(Text1) = "" Then Exit Sub
   CantEnDepositoAcumulado = 0
   cantidadColumnasOriginal = 8
   cantidadColumnasAgregadas = XVALO(UBound(CADENATEX$))
   totalColumnas = cantidadColumnasOriginal + cantidadColumnasAgregadas
   For REG& = 1 To MSF.Rows - 1
   CantEnDepositoAcumulado = 0
    For COL& = cantidadColumnasOriginal + 1 To totalColumnas
        CI1% = XVALO(MSF.TextMatrix(REG&, 1))
        indDepo% = indDepo% + 1
        If Int(CADENATEX$(indDepo%)) = 5 Then 'esto es por que arcon en el deposito 5 tiene lo pendiente de C.C.
           CantEnDeposito = DepositoRecepcionArcon(CI1%)
        Else
           CantEnDeposito = STODEPOS(CI1%, Int(CADENATEX$(indDepo%))) ' Y quiere que lo presente
        End If
        
        MSF.TextMatrix(REG&, COL&) = FORMATNUM$(CantEnDeposito, "F9.1")
        CantEnDepositoAcumulado = CantEnDepositoAcumulado + CantEnDeposito
    Next COL&
    MSF.TextMatrix(REG&, 7) = FORMATNUM$(CantEnDepositoAcumulado, "F12.1")
    indDepo% = 0
   Next REG&

End Sub
Function DepositoRecepcionArcon(CI1%)
        TORECEP = 0
        
     SQLX$ = " SELECT CANTNOM, NumProv, NUMEBORE , FECHRECEP, NumRem_Prov, CanToApro  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     SQLX$ = SQLX$ + " AND Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
     '

     Set COBRECEPTMP = READSET(SQLX$)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         JJ& = 0
         k& = 0
         Do While (.EOF = False)
            CARECEP = XVALO(!CantNom) - XVALO(!CanToApro)
            If CARECEP > 0 Then
              NCI% = XVALO(!NumProv)
              REMIPRO$ = SAFETEXT$(!NumRem_Prov)
              NROBOL& = XVALO(!NUMEBORE)
              FEX = CVINT(!FechRecep)
              TORECEP = TORECEP + CARECEP
            End If
         .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     COBRECEPTMP.Close
     Set COBRECEPTMP = Nothing
     
    Call COPYSTRU("BRECEPEN", "BRECEPEL")
    Call FILESORT("BRECEPEN", "BRECEPEL", 1, 1)
    FIN& = ULTREG&("BRECEPEL")
    
    For U& = 1 To FIN&
      TXTX$ = REGLEIDO$("BRECEPEL", U&)
      CIII% = CVI(Mid$(TXTX$, 3, 2))
      If CIII% = CI1% Then
           DOPRI$ = TRIM$(Mid$(TXTX$, 5, 16))
           CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
           CANTSALID = CVD(Mid$(TXTX$, 96, 8))
           CANTI = CANTINGRE - CANTSALID
           If XVALO(CANTI) > 0.05 Then
             TORECEP = TORECEP + XVALO(CANTI)
           End If
      End If
    Next U&
    DepositoRecepcionArcon = TORECEP

End Function
Private Sub Text3_Change()
   If XVALO(Text3) < 1 And TRIM$(Text3) <> "" Then
      Call COMUNI("Cantidad Ingresada No Válida")
      Exit Sub
   End If
   Call T_Espera(1)

   Call ActualizarInfo
End Sub

Private Sub TXTNIVELES_Change()
Call T_Espera(1)

   Call ActualizarInfo
End Sub
