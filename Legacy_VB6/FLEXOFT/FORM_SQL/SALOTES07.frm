VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form SALOTES07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Saldos por Lote"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Ajuste a 0 de Seleccionados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   39
      ToolTipText     =   "Ajuste a 0 de Seleccionados"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Seleccionar Todo  o Revertir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   38
      ToolTipText     =   "Realiza Control y Recupera Saldos por Lote"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   3270
      Left            =   105
      TabIndex        =   37
      Top             =   1680
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   5768
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1680
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   -120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Control Saldos por Lote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   34
      ToolTipText     =   "Realiza Control y Recupera Saldos por Lote"
      Top             =   5340
      Width           =   2225
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   5
      Left            =   7450
      TabIndex        =   19
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   4
      Left            =   6400
      TabIndex        =   18
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   1
      Left            =   3045
      TabIndex        =   11
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   3
      Left            =   5460
      TabIndex        =   10
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   2
      Left            =   4180
      TabIndex        =   9
      Top             =   1785
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3200
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   1785
      Width           =   15
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      FillColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   1600
      Left            =   8235
      ScaleHeight     =   1575
      ScaleWidth      =   2640
      TabIndex        =   0
      Top             =   -30
      Width           =   2670
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
         Left            =   120
         Picture         =   "SALOTES07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Terminar - Salir de Consulta de Stocks"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Left            =   105
         Picture         =   "SALOTES07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Ficha de Movimientos de Material (Ficha Kardex)"
         Top             =   840
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "SALOTES07.frx":07B4
      Top             =   720
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   3840
      OleObjectBlob   =   "SALOTES07.frx":09E8
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2040
      Width           =   7650
   End
   Begin VB.Label Label24 
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
      Left            =   5280
      TabIndex        =   33
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Total del Lote"
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
      Left            =   4170
      TabIndex        =   32
      Top             =   6810
      Width           =   1485
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Etiquetas"
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
      Left            =   5280
      TabIndex        =   31
      Top             =   5700
      Width           =   1485
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas Incompletas"
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
      Left            =   1545
      TabIndex        =   30
      Top             =   6375
      Width           =   1650
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas Completas"
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
      Left            =   1560
      TabIndex        =   29
      Top             =   6060
      Width           =   1485
   End
   Begin VB.Label Label19 
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
      Left            =   5280
      TabIndex        =   28
      Top             =   6270
      Width           =   1335
   End
   Begin VB.Label Label18 
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
      Left            =   5280
      TabIndex        =   27
      Top             =   5940
      Width           =   1335
   End
   Begin VB.Label Label17 
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
      Left            =   3825
      TabIndex        =   26
      Top             =   6285
      Width           =   1335
   End
   Begin VB.Label Label16 
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
      Left            =   3150
      TabIndex        =   25
      Top             =   6285
      Width           =   480
   End
   Begin VB.Label Label14 
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
      Left            =   3150
      TabIndex        =   24
      Top             =   5955
      Width           =   480
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cajas"
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
      Left            =   3150
      TabIndex        =   23
      Top             =   5730
      Width           =   690
   End
   Begin VB.Label Label12 
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
      Left            =   3825
      TabIndex        =   22
      Top             =   5955
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Etiq. x Caja"
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
      Left            =   3840
      TabIndex        =   21
      Top             =   5730
      Width           =   1485
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
      Height          =   495
      Left            =   1050
      TabIndex        =   20
      Top             =   570
      Width           =   5790
   End
   Begin VB.Label Label11 
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
      Left            =   1680
      TabIndex        =   17
      Top             =   1140
      Width           =   5160
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito"
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
      Left            =   -70
      TabIndex        =   16
      Top             =   1250
      Width           =   960
   End
   Begin VB.Label Label9 
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
      Left            =   1050
      TabIndex        =   15
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTALES"
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
      Left            =   1920
      TabIndex        =   14
      Top             =   5115
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3255
      TabIndex        =   13
      Top             =   5040
      Width           =   1335
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
      Left            =   1050
      TabIndex        =   12
      Top             =   210
      Width           =   2850
   End
   Begin VB.Label Label5 
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
      Left            =   6090
      TabIndex        =   7
      Top             =   210
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "U-M"
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
      Left            =   5150
      TabIndex        =   6
      Top             =   300
      Width           =   855
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
      Left            =   -50
      TabIndex        =   5
      Top             =   300
      Width           =   960
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Fecha           Id.Lote                               F.Vencim                     Saldo          Ancho            Largo"
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
      Left            =   105
      TabIndex        =   4
      Top             =   1790
      Width           =   7650
   End
End
Attribute VB_Name = "SALOTES07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Cancel = 0
   'List1.Clear
   Me.GRID.Rows = 1
   Unload Me
End Sub

Private Sub Command2_Click()
    If DERACCE%("CONSALOT", "Control de Saldos por Lote") < 2 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    CIXI% = CODINT%(Label4)
    NDEP% = XVALO(Label9)
    SQLX$ = "SELECT IDENLOTE, "
    SQLX$ = SQLX$ + " SUM(CANTINGRE) - SUM(CANTSALID) AS SALDO "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE COD_INTE = " & CIXI%
    SQLX$ = SQLX$ + " AND DEPOMOVI = " & NDEP%
    SQLX$ = SQLX$ + " GROUP BY IDENLOTE "
    '
    Dim SAMOVI As ADODB.Recordset
    Set SAMOVI = New ADODB.Recordset
    Set SAMOVI = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With SAMOVI
       Do While Not .EOF
         IDLO$ = TRIM$(SAFETEXT$(!idenlote))
         SALOTIN = XVALO(VALOBADA("LOTINGRE", "CANALTA - CANCONSU", "COD_INTE = " & CIXI% & " AND DEPOUBIC = " & NDEP% & " AND IDENLOTE = '" & IDLO$ & "'"))
         If Abs(SALOTIN - XVALO(!SALDO)) > 0.05 Then
            Call ACSALOTE(CIXI%, IDLO$)
         End If
       .MoveNext
       Loop
    End With
    '
    ' aqui faltaría tambien recorrer el lotingre con los que tienen saldo
    ' y contrastar contra el saldo del movisto
    '
    Screen.MousePointer = 1
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Screen.MousePointer = 11
    FIN& = GRID.Rows - 1
    GRID.Redraw = False
    For i& = 1 To GRID.Rows - 1
        Call TRACE(20, i&, FIN&)
        Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, i&)
    Next i&
    GRID.Redraw = True
    Screen.MousePointer = 1
    Command3.Enabled = 1
End Sub

Private Sub Command4_Click()
            Command4.Enabled = False
            
            If XVALO(Label9.Caption) <> 2 Then
               Call COMUNI("Este Proceso Esta Solo Reservado para Ajustar el Depósito 2")
               Exit Sub
            End If
             
            For i& = 1 To GRID.Rows - 1
               If GRID.TextMatrix(i&, 0) = "*" Then
                 CIXI% = CODINT%(Label4.Caption)
                 CANTING = XVALO(Me.GRID.TextMatrix(i&, 4))
   
                 '
                 If CANTING > 0 Then
                   IDLO$ = TRIM$(Me.GRID.TextMatrix(i&, 2))
                   'IDLOT$ = TRIM$(Mid$(TTXX$, 11, 16))
                   If IDLO$ <> "" Then
                     CONSUMLOT07.Label4 = Label4
                     CONSUMLOT07.Label1 = IDLO$
                     CONSUMLOT07.Text4 = XVALO(Label9.Caption)
                     Call CONSUMLOT07.Command2_Click
'                     CONSUMLOT07.Show 1
                   End If
                 End If
               End If
            Next i&
            Command4.Enabled = True
            Unload Me
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Then
     Call CopyToExcel(GRID, "REPORTE", "")
     GoTo 99
   End If
   Call MENSERR(24, "No Disponible en la Presente Versión.")
99 Command8.Enabled = True
End Sub

Private Sub Form_Load()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Campos$ = "Fecha/D6;Id Lote/A10;F.Vencim/D6;Saldo/F7.1;Ancho/F7.1;Largo/F7.1;Ref.(1)/A9"
   If InStr(EPAC$, "ACHERNAR") > 0 Then Campos$ = Campos$ + ";Ref.(2)/A9"
   
   If InStr(EPAC$, "CAMBACERES") > 0 Then
      Campos$ = "Fecha/D6;Id Lote/A10;F.Vencim/D6;Saldo/F7.1;Kilos/F7.1;Metros/F7.1;Ref.(1)/A9"
   End If
   Call CARGA_ENCABEZADO(Me.GRID, Campos$, FRMZELECHO, , , ANTOT)
   If InStr(EMPREAC$, "ACHERNAR") > 0 Then
      SALOTES07.Height = 7485
   End If
   
   If InStr(EPAC$, "CONTEPLAST") > 0 Then 'BOTONES PARA REALIZAR AJUSTES DE LOTES SELECCIONADOS.
      Command3.Visible = True
      Command4.Visible = True
   End If
   
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub GRID_Click()
  REG& = Me.GRID.MouseRow
  If InStr(EMPREAC$, "ACHERNAR") > 0 Then
    CI% = CODINT(Label4)
    'IDLO$ = TRIM(Mid$(List1.TEXT, 11, 16))
    IDLO$ = Me.GRID.TextMatrix(REG&, 2)
    CONDI$ = "IDENLOTE ='" & IDLO$ & "' AND COD_INTE=" & CI%
    '
    VETIQXCAJ# = XVALO(VALOBADA("LOTINGRE", "EtiqxCaja", CONDI$))
    VCANALTA# = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
    VCanConsu# = XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDI$))
    
    VCANSALDO# = VCANALTA# - VCanConsu#
    If VCANSALDO# < 0 Then VCANSALDO# = 0
    If VETIQXCAJ# < 1 Then VETIQXCAJ# = VCANALTA#
    '
    'calculo el valor de cajas
    On Error Resume Next
       CAJ# = VCANSALDO# / VETIQXCAJ#
       CAJENT& = Int(CAJ#)
       CAJRES# = CAJ# - CAJENT&
       If Err > 0 Then
          CAJ# = 0: CAJENT& = 0: CAJRES# = 0
       End If
    On Error GoTo 0
    '
    'calculo el valor de cantidades
    CantPri# = CAJENT& * VETIQXCAJ# 'CANTIDAD Etiquetas por caja, DEL PRIMER REGISTRO
    CANTRES# = VCANSALDO# - CantPri#  'CANTIDAD RESTANTE
    '
    Label14 = CAJENT& ' cajas completas
    If CAJRES# > 0 Or CANTRES# > 0 Then
         Label16 = 1 'CAJA INCOMPLETA
       Else
         Label16 = "0"
    End If
    '
    '
    Label12 = VETIQXCAJ#  ' ETIQUETA X CAJAS completas
    Label17 = CANTRES#  ' ETIQUETA X CAJAS incompletas
    '
    Label18 = VETIQXCAJ# * CAJENT& 'Total etiq completas
    Label19 = CANTRES#           'Total etiq incompletas
    '
    Label24 = XVALO(Label18) + XVALO(Label19)
    '
  End If
  If InStr(EMPREAC$, "CONTEPLAST") > 0 Then
     Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
  End If
  
End Sub

Private Sub GRID_DblClick()
   REG& = Me.GRID.MouseRow
   'TTXX$ = List1.TEXT
   'CANTING = XVALO(Mid$(TTXX$, 27, 12))
   CANTING = XVALO(Me.GRID.TextMatrix(REG&, 4))
   If SALOTES07.Height < 6000 Then 'ENTONCES VINO DE AJUSTK07
       IDLO$ = TRIM$(Me.GRID.TextMatrix(REG&, 2))
       AJUSTK07.Label17 = TRIM$(IDLO$)
       Unload Me
       Exit Sub
   End If
   '
   If CANTING > 0 Then
      IDLO$ = TRIM$(Me.GRID.TextMatrix(REG&, 2))
      'IDLOT$ = TRIM$(Mid$(TTXX$, 11, 16))
      If IDLO$ <> "" Then
        CONSUMLOT07.Label4 = Label4
        CONSUMLOT07.Label1 = IDLO$
        CONSUMLOT07.Show 1
      End If
   End If

End Sub

Private Sub Label4_Change()
    CIXI% = CODINT%(Label4)
    Label5 = Unimed$(CIXI%)
    Label6 = DESCRIT0$(CIXI%)
End Sub

Sub MUESTRASALOT()
   'List1.Clear
   GRID.Rows = 1
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
'   Call ARMALILOT(CIXI%, DEPO%, FAJUI%)

   CIXI% = CODINT%(Label4)
   DEPO% = XVALO(Label9)
   CADECI% = XVALO(CNTDC$(CIXI%))
   '
   FE1% = FECHANUM("01/01/80")
'   FEINIKAR$ = TRIM$(CONTROL("", "FEINIKAR"))
'   FF% = FECHANUM(FEINIKAR$)
'   If FF% > 0 Then
'      Text11 = FECHATEX$(FF%)
'
'   Else
'      Text11 = FECHATEX$(FF1%)
'      FF% = FE1%
'   End If

'   Call ARMALILOT(CIXI%, DEPO%, FF1%)
   J& = 0
   ESACHERNAR% = 0
   If InStr(EMPREAC$, "ACHERNAR") > 0 Then ESACHERNAR% = 1
   For U& = 1 To ULTREG&("!LILOTSE")
       TTYY$ = REGLEIDO$("!LILOTSE", U&)
       TTXX$ = Space$(128)
         FEALTA$ = Mid$(TTYY$, 17, 8)
'       Call REPLA(TTXX$, FEALTA$, 1, 8)
         IDLO$ = Mid$(TTYY$, 1, 16)
'       Call REPLA(TTXX$, IDLO$, 11, 16)
         FEVEX$ = Mid$(TTYY$, 47, 8)
'       Call REPLA(TTXX$, FEVEX$, 30, 8)
         ANCHX = XVALO(Mid$(TTYY$, 55, 8)) ' se corrige la posiciones de lectura
         LARGX = XVALO(Mid$(TTYY$, 66, 8))
         If ANCHX >= 0.5 And LARGX >= 0.05 And Abs(XVALO(Mid$(TTYY$, 35, 12))) >= 1 Then
              CANSA = ANCHX * LARGX / 1000
            Else
              CANSA = XVALO(Mid$(TTYY$, 35, 12))
         End If
'       Call REPLA(TTXX$, FORMATNUM$(CANSA, "F12." & CADECI%), 39, 12)
         SUMASAL = SUMASAL + CANSA
'       Call REPLA(TTXX$, FORMATNUM$(ANCHX, "F9.1"), 51, 9)
'       Call REPLA(TTXX$, FORMATNUM$(LARGX, "F10.1"), 60, 10)
        REF1$ = "": REF2$ = ""
        If Left$(TRIM$(IDLO$), 3) = "OF-" Then
          CONDI$ = "IDSUBOR = '" & IDLO$ & "'"
          REF1$ = VALOBADA("ORDEFABR", "REFEREN", CONDI$, "NOMESS")
          If ESACHERNAR% > 0 Then REF2$ = VALOBADA("ORDEFABR", "ID_ORFA", CONDI$, "NOMESS")
        End If
       'List1.AddItem TTXX$
       If Abs(CANSA) > 0.05 Then
            J& = J& + 1
            GRID.Rows = J& + 1
            GRID.TextMatrix(J&, 1) = FEALTA$
            GRID.TextMatrix(J&, 2) = IDLO$
            GRID.TextMatrix(J&, 3) = FEVEX$
            GRID.TextMatrix(J&, 4) = FORMATNUM$(CANSA, "F12." & CADECI%)
            If InStr(EPAC$, "CAMBACERES") > 0 Then
               CONDI$ = "IDENLOTE = '" & IDLO$ & "' AND (CODIMOVI = 'OF' OR CODIMOVI = 'BR') AND (METROS > 0 OR PESOKG > 0)"
               METROS# = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
               KG1# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
              GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(KG1#), "F10.2")
              GRID.TextMatrix(J&, 6) = FORMATNUM$(XVALO(METROS), "F10.2")
            Else
              GRID.TextMatrix(J&, 5) = FORMATNUM$(ANCHX, "F9.1")
              GRID.TextMatrix(J&, 6) = FORMATNUM$(LARGX, "F10.1")
            End If
            GRID.TextMatrix(J&, 7) = REF1$
            If ESACHERNAR% > 0 Then GRID.TextMatrix(U&, 8) = REF2$
       End If
   Next U&
   '
   'EN ESTE BLOQUE SE AGREGA SI HAY RESULTADOS SIN LOTE EN EL CASO DE LOS TRAZABLES PARA QUE COINCIDA
   'CON LA CANTIDAD QUE PRESENTA EN EL LIST DE EXISTENCIAS.
   Resultado = 0
   If ESTRAZABLE%(CIXI%) = 1 Then
    SQLX$ = "SELECT SUM(CANTSALID) AS SALI , SUM(CANTINGRE) AS INGRE FROM MOVISTO WITH(NOLOCK) WHERE COD_INTE = " & CIXI%
    SQLX$ = SQLX$ + " AND (IDENLOTE IS NULL OR IDENLOTE = '') AND DEPOMOVI = " & DEPO%
'    FE% = FECHANUM(CONTROL("", "FEINIKAR"))
'    If FE% > 0 Then SQLX$ = SQLX$ + " AND FECHMOV = '" & FETEXCOR1$(FE%) & "'"
    
    Set rs = READSET(SQLX$)
    Resultado = XVALO(rs!INGRE) - XVALO(rs!sali)
    If Abs(Resultado) > 0 Then
        J& = J& + 1
        GRID.Rows = J& + 1
        GRID.TextMatrix(J&, 1) = "S/Lote"
        GRID.TextMatrix(J&, 2) = ""
        GRID.TextMatrix(J&, 3) = ""
        GRID.TextMatrix(J&, 4) = FORMATNUM$(Resultado, "F12." & CADECI%)
        GRID.TextMatrix(J&, 5) = ""
        
        GRID.TextMatrix(J&, 5) = ""
        GRID.TextMatrix(J&, 6) = ""
        GRID.TextMatrix(J&, 7) = ""
    End If
    ST1 = STODEPOS(CIXI%, DEPO%) '
    'SI QUEDA ALGUNA DIFERENCIA FUERZA A QUE CIERRE LA CANTIDAD GENERANDO UN REGISTRO POR LA DIFERENCIA
    DIFEREX = ST1 - (SUMASAL + Resultado)
    If Abs(DIFEREX) > 0.05 Then
        J& = J& + 1
        GRID.Rows = J& + 1
        GRID.TextMatrix(J&, 1) = "Ajustes"
        GRID.TextMatrix(J&, 2) = " y Otros"
        GRID.TextMatrix(J&, 3) = ""
        GRID.TextMatrix(J&, 4) = FORMATNUM$(DIFEREX, "F12." & CADECI%)
        GRID.TextMatrix(J&, 5) = ""
        GRID.TextMatrix(J&, 6) = ""
        GRID.TextMatrix(J&, 7) = ""
     End If
   End If
   Label7 = TRIM$(FORMATNUM$(SUMASAL + Resultado + DIFEREX, "F12." & CADECI%))
   Screen.MousePointer = 1
   'CAMPOS$ = "Fecha/D6;Id Lote/A10;F.Vencim/D6;Saldo/F7.1;Ancho/F7.1;Largo/F7.1;Ref.(1)/A9;Ref.(2)/A9"

End Sub

'
Private Sub Label9_Change()
    If XVALO(Label9) > 0 Then
         Label11 = ECOARCH$("TADEPOSI", MKI$(XVALO(Label9)))
      Else
         Label11 = "Todos los Depositos"
    End If
End Sub

Private Sub List1_Click()
  '
  REG& = Me.GRID.MouseRow
  If InStr(EMPREAC$, "ACHERNAR") > 0 Then
    CI% = CODINT(Label4)
    'IDLO$ = TRIM(Mid$(List1.TEXT, 11, 16))
    IDLO$ = Me.GRID.TextMatrix(REG&, 2)
    CONDI$ = "IDENLOTE ='" & IDLO$ & "' AND COD_INTE=" & CI%
    '
    VETIQXCAJ# = XVALO(VALOBADA("LOTINGRE", "EtiqxCaja", CONDI$))
    VCANALTA# = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
    VCanConsu# = XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDI$))
    
    VCANSALDO# = VCANALTA# - VCanConsu#
    If VCANSALDO# < 0 Then VCANSALDO# = 0
    If VETIQXCAJ# < 1 Then VETIQXCAJ# = VCANALTA#
    '
    'calculo el valor de cajas
    On Error Resume Next
       CAJ# = VCANSALDO# / VETIQXCAJ#
       CAJENT& = Int(CAJ#)
       CAJRES# = CAJ# - CAJENT&
       If Err > 0 Then
          CAJ# = 0: CAJENT& = 0: CAJRES# = 0
       End If
    On Error GoTo 0
    '
    'calculo el valor de cantidades
    CantPri# = CAJENT& * VETIQXCAJ# 'CANTIDAD Etiquetas por caja, DEL PRIMER REGISTRO
    CANTRES# = VCANSALDO# - CantPri#  'CANTIDAD RESTANTE
    '
    Label14 = CAJENT& ' cajas completas
    If CAJRES# > 0 Or CANTRES# > 0 Then
         Label16 = 1 'CAJA INCOMPLETA
       Else
         Label16 = "0"
    End If
    '
    '
    Label12 = VETIQXCAJ#  ' ETIQUETA X CAJAS completas
    Label17 = CANTRES#  ' ETIQUETA X CAJAS incompletas
    '
    Label18 = VETIQXCAJ# * CAJENT& 'Total etiq completas
    Label19 = CANTRES#           'Total etiq incompletas
    '
    Label24 = XVALO(Label18) + XVALO(Label19)
    '
  End If
  '
End Sub

Private Sub List1_DblClick()
   '
   REG& = Me.GRID.MouseRow
   TTXX$ = List1.TEXT
   'CANTING = XVALO(Mid$(TTXX$, 27, 12))
   CANTING = XVALO(Me.GRID.TextMatrix(REG&, 4))
   
   '
   If CANTING > 0 Then
      IDLO$ = Me.GRID.TextMatrix(REG&, 2)
      'IDLOT$ = TRIM$(Mid$(TTXX$, 11, 16))
      If IDLOT$ <> "" Then
        CONSUMLOT07.Label4 = Label4
        CONSUMLOT07.Label1 = IDLOT$
        CONSUMLOT07.Show 1
      End If
   End If
   '
End Sub
