VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RepoRega 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programa de Fabricación Por Destino / Cliente"
   ClientHeight    =   9090
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   18750
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   18750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox DestinoTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0D0D0&
      Caption         =   "Todos los Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   26
      Top             =   240
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   3195
      Left            =   120
      TabIndex        =   8
      Top             =   3800
      Width           =   18555
      _ExtentX        =   32729
      _ExtentY        =   5636
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   2355
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   18555
      _ExtentX        =   32729
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXTVENCIMIENTO2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox TXTDIFERENCIA2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   17640
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1020
   End
   Begin VB.TextBox TXTDIFERENCIA1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   17640
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   1020
   End
   Begin VB.TextBox TXTVENCIMIENTO1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15120
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TXTECODESTINO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton Command6 
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
      Height          =   330
      Left            =   2370
      TabIndex        =   6
      ToolTipText     =   "Despliega Lista de Selección"
      Top             =   120
      Width           =   160
   End
   Begin VB.TextBox TXTDESTINO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "RepoRega.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "RepoRega.frx":0234
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   2355
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   4154
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
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
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1995
      Left            =   14040
      TabIndex        =   10
      Top             =   1320
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   3519
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1275
      Left            =   14040
      TabIndex        =   13
      Top             =   5520
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   2249
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF4 
      Height          =   1635
      Left            =   120
      TabIndex        =   16
      Top             =   7400
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   2884
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   25
      Top             =   5115
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14040
      TabIndex        =   24
      Top             =   915
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUBTOTALES Y TOTALES"
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
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   7080
      Width           =   18540
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia en Días"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16650
      TabIndex        =   22
      Top             =   5120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia en Días"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16680
      TabIndex        =   21
      Top             =   920
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PENDIENTE DE ENTREGA"
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
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   18540
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTREGADOS"
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
      Height          =   280
      Left            =   120
      TabIndex        =   19
      Top             =   650
      Width           =   18540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Hoy"
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
      Left            =   11400
      TabIndex        =   18
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino / Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   50
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuExpoExcel 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu Buscar 
      Caption         =   "Buscar"
      Begin VB.Menu BuscarEntregados 
         Caption         =   "Buscar en Entregados"
      End
      Begin VB.Menu BuscarPendientes 
         Caption         =   "Buscar en Pendientes"
      End
   End
End
Attribute VB_Name = "RepoRega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextoBuscarEntregados$
Dim TextoBuscarPendientes$

Sub CARGA_VENCIMIENTO_MSF0()
   MAXIMOFE% = 33
   'CONSIDERAR LA FECHA MAS CERCANA
   For i& = 1 To MSF0.Rows - 1
      FE% = FECHANUM(MSF0.TextMatrix(i&, 2))
      If FE% > MAXIMOFE% Then
         MAXIMOFE% = FE%
      End If
   Next i&
   '
   If MAXIMOFE% = 33 Then
      Me.TXTVENCIMIENTO1 = ""
      Me.TXTDIFERENCIA1 = ""
      Exit Sub
   End If
   
   FE% = HOY(HO$)
   Me.TXTVENCIMIENTO1 = FECHATEX$(MAXIMOFE%)
   Me.TXTDIFERENCIA1 = DiferenciaDiasEntreFechas(MAXIMOFE%, HOY(HO$))
End Sub

Sub CARGA_VENCIMIENTO_MSF3()
   MAXIMOFE% = 33
   'CONSIDERAR LA FECHA MAS CERCANA
   For i& = 1 To MSF3.Rows - 1
      FE% = FECHANUM(MSF3.TextMatrix(i&, 2))
      If FE% > MAXIMOFE% Then
         MAXIMOFE% = FE%
      End If
   Next i&
   '
   If MAXIMOFE% = 33 Then
      Me.TXTVENCIMIENTO1 = ""
      Me.TXTDIFERENCIA1 = ""
      Exit Sub
   End If
   
   FE% = HOY(HO$)
   Me.TXTVENCIMIENTO2 = FECHATEX$(MAXIMOFE%)
   Me.TXTDIFERENCIA2 = DiferenciaDiasEntreFechas(MAXIMOFE%, FE%)
End Sub

Sub CARGATOTALES()

    MSF4.Rows = 1: J& = 0
    ACUMU2 = 0
    ACUMU3 = 0
    ACUMU4 = 0
    ACUMU5 = 0

    SQLX$ = " SELECT M.FAMILIA, SUM(CANAPRO) AS CANP, SUM (CANPROY) AS CAPROY"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    If DestinoTodos.Value = 0 Then
        SQLX$ = SQLX$ + " AND  O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    End If
    SQLX$ = SQLX$ + " AND (O.StatuOrf <= 3)"
    SQLX$ = SQLX$ + " GROUP BY M.FAMILIA"
    SQLX$ = SQLX$ + " ORDER BY M.FAMILIA"
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
          TAFAM$ = SAFETEXT$(!familia)
          FAMIX$ = ECOARCH$("TAFAMIL", TAFAM$)
          CANTIAP = XVALO(!CANP)
          CANPRO = XVALO(!CAPROY)
          CANTI = XVALO(!CAPROY)
          J& = J& + 1
          MSF4.Rows = J& + 1
          MSF4.TextMatrix(J&, 1) = FAMIX$
          MSF4.TextMatrix(J&, 2) = FORMATNUM$(CANTI, "F12.1")
          MSF4.TextMatrix(J&, 3) = FORMATNUM$(CANTI - CANTIAP, "F12.1")
          MSF4.TextMatrix(J&, 4) = FORMATNUM$(CANTIAP, "F12.1")
          MSF4.TextMatrix(J&, 5) = FORMATNUM$(CANTIAP * 100 / CANTI, "F12.1")
          ACUMU2 = ACUMU2 + XVALO(MSF4.TextMatrix(J&, 2))
          ACUMU3 = ACUMU3 + XVALO(MSF4.TextMatrix(J&, 3))
          ACUMU4 = ACUMU4 + XVALO(MSF4.TextMatrix(J&, 4))
          ACUMU5 = ACUMU5 + XVALO(MSF4.TextMatrix(J&, 5))
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
    MSF4.Rows = MSF4.Rows + 2
    MSF4.TextMatrix(MSF4.Rows - 1, 1) = "Totales:"
    MSF4.TextMatrix(MSF4.Rows - 1, 2) = FORMATNUM$(ACUMU2, "F12.1")
    MSF4.TextMatrix(MSF4.Rows - 1, 3) = FORMATNUM$(ACUMU3, "F12.1")
    MSF4.TextMatrix(MSF4.Rows - 1, 4) = FORMATNUM$(ACUMU4, "F12.1")
    MSF4.TextMatrix(MSF4.Rows - 1, 5) = FORMATNUM$(ACUMU5, "F12.1")
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(Me.MSF4, &HE0E0E0, MSF4.Rows - 1)
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For h& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, h&) = GRID.TextMatrix(i&, h&)
            Next h&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub
Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub BuscarEntregados_Click()

    TextoBuscarEntregados$ = TRIM$(InputBox("Texto a buscar", "Búsqueda en Entregados", ""))
    If TextoBuscarEntregados$ <> "" Then
        For FilaGrid& = 1 To MSF0.Rows - 1
            For ColumnaGrid& = 1 To MSF0.Cols - 1
                MSF0.col = ColumnaGrid&
                If MSF0.Width > 0 Then
                    TextoCelda$ = SAFETEXT$(MSF0.TextMatrix(FilaGrid&, ColumnaGrid&))
                    If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarEntregados$)) > 0 Then
                        Call SELECCIONARFILA(MSF0, FilaGrid&)
                        GoTo 99
                    End If
                End If
            Next ColumnaGrid&
        Next FilaGrid&
    End If
    
99  Exit Sub

End Sub

Private Sub BuscarPendientes_Click()

    TextoBuscarPendientes$ = TRIM$(InputBox("Texto a buscar", "Búsqueda en Entregados", ""))
    If TextoBuscarPendientes$ <> "" Then
        For FilaGrid& = 1 To MSF3.Rows - 1
            For ColumnaGrid& = 1 To MSF3.Cols - 1
                MSF3.col = ColumnaGrid&
                If MSF3.Width > 0 Then
                    TextoCelda$ = SAFETEXT$(MSF3.TextMatrix(FilaGrid&, ColumnaGrid&))
                    If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarPendientes$)) > 0 Then
                        Call SELECCIONARFILA(MSF3, FilaGrid&)
                        GoTo 99
                    End If
                End If
            Next ColumnaGrid&
        Next FilaGrid&
    End If
    
99  Exit Sub

End Sub

Private Sub DestinoTodos_Click()

    TXTDESTINO.TEXT = ""
    If DestinoTodos.Value = 1 Then
        TXTDESTINO.TEXT = "Sin Filtro"
        TXTECODESTINO.TEXT = "Todos los Clientes"
    End If
    
End Sub

Private Sub Form_Load()

     Campos$ = "CODINT/F0;Vencimiento/D9;N° Plano/A16;Descripción/A44;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;OK/D8;Operador/A12"
     Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
     Call CARGA_ENCABEZADO(Me.MSF_2, Campos$, Me)
     Text1 = FECHATEX$(HOY(HO$))
     '
     Campos$ = "Grupo-Familia/A24;Cantidad/F9.1;CODFAMILIA/A0"
     Call CARGA_ENCABEZADO(Me.MSF1, Campos$, Me)
     '
     Campos$ = "Grupo-Familia/A24;Cantidad/F9.1;CODFAMILIA/A0"
     Call CARGA_ENCABEZADO(Me.MSF2, Campos$, Me)
     '
     Campos$ = "CODINT/F0;Vencimiento/D9;Calculada/D9;N° Plano/A16;Descripción/A36;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;Conjunto/A16;Calidad/F8.1;Operación/A32;NumeroOrfa/A0"
     Call CARGA_ENCABEZADO(Me.MSF3, Campos$, Me)
     '
     Campos$ = "Grupo-Familia/A32;Totales/F24.1;Pendientes de Entrega/F24.1;Entregados/F24.1;% Entregados/F24.1"
     Call CARGA_ENCABEZADO(Me.MSF4, Campos$, Me)

End Sub

Private Sub TXTDATO1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDATO2.SetFocus
   End If
End Sub


Private Sub TXTDATO2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDATO3.SetFocus
   End If

End Sub


Private Sub TXTDATO3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Command1.SetFocus
   End If

End Sub

Private Sub Command6_Click()
  '
  DestinoTodos.Value = 0
  MSF0.Rows = 1
  MSF1.Rows = 1
  MSF2.Rows = 1
  MSF3.Rows = 1
  MSF4.Rows = 1
  
  Call SELECHO("DESTFAB")
  COD$ = VALACT1$("DESTFAB")
  Me.TXTDESTINO = COD$
  '
End Sub

Private Sub mnuExpoExcel_Click()
   '
   If TRIM$(TXTDESTINO) = "" Then
        Call COMUNI("Imposible Exportar.\Debe Elegir un Destino de Fabricación")
        Exit Sub
   End If
   '
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call ExportaProgramaDeFabricacion(AREXPO$)
   End If
   '
99: Exit Sub
End Sub
Sub ExportaProgramaDeFabricacion(NOMBREDEARCHIVO$)
   
    Dim XLAPP As Excel.Application
    Dim XLBOOK As Excel.Workbook
    Dim XLSHEET As Excel.Worksheet
    '
    On Error Resume Next
    Set XLAPP = New Excel.Application
    Set XLBOOK = XLAPP.Workbooks.Add
    Set XLSHEET = XLBOOK.Worksheets.Add
    '
    ' *** DATOS DE ENCABEZADO
    XLSHEET.Cells(1, 1) = TXTECODESTINO
    With XLSHEET.Cells(1, 1)
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Underline = True
        .Font.Italic = True
    End With
    XLSHEET.Range("A1:Z1").Merge
    XLSHEET.Range("A1:Z1").HorizontalAlignment = xlLeft
    
    ' *** ESTABLECE LA FILA DE INICIO Y OBTIENE LA DE FIN PARA EL BLOQUE DE LOS ENTREGADOS
    filaInicio& = 4
    FilaFin& = MSF0.Rows - 1
    If MSF1.Rows > MSF0.Rows Then FilaFin& = MSF1.Rows - 1
    
    ' *** EXPORTA ENTREGADOS
    XLSHEET.Cells(filaInicio&, 1) = Label3
    XLSHEET.Range("A" & CStr(filaInicio&) & ":" & "L" & CStr(filaInicio&)).Interior.COLOR = RGB(239, 228, 176)
    filaInicio& = 5
    For FilaEntregados& = 0 To MSF0.Rows - 1
        ColumnasVisibles& = 0
        For ColumnaEntregados& = 1 To MSF0.Cols - 1
                        
            If MSF0.ColWidth(ColumnaEntregados&) > 0 Then
                ColumnasVisibles& = ColumnasVisibles& + 1
                CELDA$ = MSF0.TextMatrix(FilaEntregados&, ColumnaEntregados&)
                XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&) = CELDA$
                
                Call EstiloCeldaSegunGrilla(MSF0, XLSHEET, FilaEntregados&, ColumnaEntregados&, filaInicio& + FilaEntregados&, ColumnasVisibles&)
                If FilaEntregados& = 0 Then XLSHEET.Cells(filaInicio&, ColumnasVisibles&).Interior.COLOR = RGB(223, 233, 245)
                
                If FilaEntregados& > 0 Then
                    With XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&)
                        Select Case ColumnasVisibles&
                            Case 1, 7: XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&) = FETEXCOR2$(FECHANUM(CELDA$))
                                       .NumberFormat = "dd/mm/yy"
                        End Select
                    End With
                End If

            End If
        
        Next ColumnaEntregados&
    Next FilaEntregados&
    
    For FilaEntregados& = 0 To MSF1.Rows - 1
        columnaInicio& = ColumnasVisibles& + 2
        For ColumnaEntregados& = 1 To MSF1.Cols - 1
            
            If MSF1.ColWidth(ColumnaEntregados&) > 0 Then
                columnaInicio& = columnaInicio& + 1
                CELDA$ = MSF1.TextMatrix(FilaEntregados&, ColumnaEntregados&)
                XLSHEET.Cells(filaInicio& + FilaEntregados&, columnaInicio&) = CELDA$
                
                Call EstiloCeldaSegunGrilla(MSF1, XLSHEET, FilaEntregados&, ColumnaEntregados&, filaInicio& + FilaEntregados&, columnaInicio&)
                If FilaEntregados& = 0 Then XLSHEET.Cells(filaInicio&, columnaInicio&).Interior.COLOR = RGB(223, 233, 245)
            End If
        
        Next ColumnaEntregados&
    Next FilaEntregados&
   
    
    ' *** ESTABLECE LA FILA DE INICIO Y OBTIENE LA DE FIN PARA EL BLOQUE DE PENDIENTES DE ENTREGA
    filaInicio& = FilaFin& + 7
    FilaFin& = filaInicio& + MSF3.Rows - 1
    If MSF2.Rows > MSF3.Rows Then FilaFin& = filaInicio& + MSF2.Rows - 1
    ColumnasVisibles& = 0
    
    ' *** EXPORTA PENDIENTES DE ENTREGA
    XLSHEET.Cells(filaInicio&, 1) = Label4
    XLSHEET.Range("A" & CStr(filaInicio&) & ":" & "L" & CStr(filaInicio&)).Interior.COLOR = RGB(239, 228, 176)
    filaInicio& = filaInicio& + 1
    For FilaEntregados& = 0 To MSF3.Rows - 1
        ColumnasVisibles& = 0
        For ColumnaEntregados& = 1 To MSF3.Cols - 1
        
            If MSF3.ColWidth(ColumnaEntregados&) > 0 Then
                ColumnasVisibles& = ColumnasVisibles& + 1
                CELDA$ = MSF3.TextMatrix(FilaEntregados&, ColumnaEntregados&)
                XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&) = CELDA$
                
                Call EstiloCeldaSegunGrilla(MSF3, XLSHEET, FilaEntregados&, ColumnaEntregados&, filaInicio& + FilaEntregados&, ColumnasVisibles&)
                If FilaEntregados& = 0 Then XLSHEET.Cells(filaInicio&, ColumnasVisibles&).Interior.COLOR = RGB(223, 233, 245)
                
                If FilaEntregados& > 0 Then
                    With XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&)
                        Select Case ColumnasVisibles&
                            Case 1: XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&) = FETEXCOR2$(FECHANUM(CELDA$))
                                    .NumberFormat = "dd/mm/yy"
                        End Select
                    End With
                End If
            End If
        
        Next ColumnaEntregados&
    Next FilaEntregados&
    
    For FilaEntregados& = 0 To MSF2.Rows - 1
        columnaInicio& = ColumnasVisibles& + 2
        For ColumnaEntregados& = 1 To MSF2.Cols - 1
        
            If MSF2.ColWidth(ColumnaEntregados&) > 0 Then
                columnaInicio& = columnaInicio& + 1
                CELDA$ = MSF2.TextMatrix(FilaEntregados&, ColumnaEntregados&)
                XLSHEET.Cells(filaInicio& + FilaEntregados&, columnaInicio&) = CELDA$
                
                Call EstiloCeldaSegunGrilla(MSF2, XLSHEET, FilaEntregados&, ColumnaEntregados&, filaInicio& + FilaEntregados&, columnaInicio&)
                If FilaEntregados& = 0 Then XLSHEET.Cells(filaInicio&, columnaInicio&).Interior.COLOR = RGB(223, 233, 245)
            End If
        
        Next ColumnaEntregados&
    Next FilaEntregados&
    
    
    ' *** OBTIENE LA FILA DE INICIO PARA EL BLOQUE DE TOTALES
    filaInicio& = FilaFin& + 10
    FilaFin& = filaInicio& + MSF4.Rows - 1
   
    ' *** EXPORTA TOTALES
    XLSHEET.Cells(filaInicio&, 1) = Label7
    XLSHEET.Range("A" & CStr(filaInicio&) & ":" & "L" & CStr(filaInicio&)).Interior.COLOR = RGB(239, 228, 176)
    filaInicio& = filaInicio& + 1
    For FilaEntregados& = 0 To MSF4.Rows - 1
        ColumnasVisibles& = 0
        For ColumnaEntregados& = 1 To MSF4.Cols - 1
        
            If MSF4.ColWidth(ColumnaEntregados&) > 0 Then
                ColumnasVisibles& = ColumnasVisibles& + 1
                CELDA$ = MSF4.TextMatrix(FilaEntregados&, ColumnaEntregados&)
                XLSHEET.Cells(filaInicio& + FilaEntregados&, ColumnasVisibles&) = CELDA$
                
                Call EstiloCeldaSegunGrilla(MSF4, XLSHEET, FilaEntregados&, ColumnaEntregados&, filaInicio& + FilaEntregados&, ColumnasVisibles&)
                If FilaEntregados& = 0 Then XLSHEET.Cells(filaInicio&, ColumnasVisibles&).Interior.COLOR = RGB(223, 233, 245)
            End If
        
        Next ColumnaEntregados&
    Next FilaEntregados&
    
    XLSHEET.Columns.AutoFit

    On Error Resume Next
    XLSHEET.SaveAs NOMBREDEARCHIVO$
    On Error GoTo 0
    '
    XLAPP.QUIT      ' cierra EXCEL que se esta ejecutando minimizado
    Set XLAPP = Nothing     ') libera memoria
    Set XLBOOK = Nothing    ') utilizada por
    Set XLSHEET = Nothing   ') las variables objeto
    '
    Screen.MousePointer = 1
    Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))

End Sub







Private Sub MSF0_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        If TextoBuscarEntregados$ <> "" Then
            FilaComienzo& = MSF0.row
            If FilaComienzo& = MSF0.Rows - 1 Then FilaComienzo& = 1
20          For FilaGrid& = FilaComienzo& + 1 To MSF0.Rows - 1
                For ColumnaGrid& = 1 To MSF0.Cols - 1
                    MSF0.col = ColumnaGrid&
                    If MSF0.Width > 0 Then
                        TextoCelda$ = SAFETEXT$(MSF0.TextMatrix(FilaGrid&, ColumnaGrid&))
                        If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarEntregados$)) > 0 Then
                            Call SELECCIONARFILA(MSF0, FilaGrid&)
                            GoTo 99
                        End If
                    End If
                Next ColumnaGrid&
            Next FilaGrid&
            If PrimeraVuelta% = 0 Then
                PrimeraVuelta% = 1
                FilaComienzo& = 0
                GoTo 20
            End If
        End If
    End If
    
99  Exit Sub

End Sub

Private Sub MSF3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 114 Then
        If TextoBuscarPendientes$ <> "" Then
            FilaComienzo& = MSF3.row
            If FilaComienzo& = MSF3.Rows - 1 Then FilaComienzo& = 1
20          For FilaGrid& = FilaComienzo& + 1 To MSF3.Rows - 1
                For ColumnaGrid& = 1 To MSF3.Cols - 1
                    MSF3.col = ColumnaGrid&
                    If MSF3.Width > 0 Then
                        TextoCelda$ = SAFETEXT$(MSF3.TextMatrix(FilaGrid&, ColumnaGrid&))
                        If InStr(1, UCase(TextoCelda$), UCase(TextoBuscarPendientes$)) > 0 Then
                            Call SELECCIONARFILA(MSF3, FilaGrid&)
                            GoTo 99
                        End If
                    End If
                Next ColumnaGrid&
            Next FilaGrid&
            If PrimeraVuelta% = 0 Then
                PrimeraVuelta% = 1
                FilaComienzo& = 0
                GoTo 20
            End If
        End If
    End If
    
99  Exit Sub

End Sub

Private Sub MSF3_DblClick()

    If MSF3.row > 0 Then
        OrdenFabricacion$ = TRIM$(MSF3.TextMatrix(MSF3.row, 12))
        Call CONECRUN("IFABRI07/INFONUOF" + OrdenFabricacion$, "Registro de Trazabilidad de O/F's")
    End If

End Sub


Private Sub TXTDESTINO_Change()
    If DestinoTodos.Value = 0 Then
        Me.TXTECODESTINO = TRIM$(ECOARCH$("DESTFAB", Me.TXTDESTINO))
        DestinoTodos.Value = 0
    End If

    Call CARGAMSF0
    Call CARGA_VENCIMIENTO_MSF0
    Call CARGAMSF1
    Call CARGAMSF2
    Call CARGAMSF3
    Call CARGA_VENCIMIENTO_MSF3
    Call CARGATOTALES
End Sub
Sub CARGAMSF0()
    MSF0.Rows = 1
    If DestinoTodos.Value = 0 Then
        Campos$ = "CODINT/F0;Vencimiento/D9;N° Plano/A16;Descripción/A44;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;OK/D8;Operador/A12"
        Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
    Else
        Campos$ = "CODINT/F0;Vencimiento/D9;N° Plano/A16;Descripción/A44;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;OK/D8;Operador/A12;Destino/A24"
        Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
    End If
    SQLX$ = " SELECT M.DESCRIDOS, M.DESCRIPCION,M.NPLANO ,M.Codigo, O.IDSUBOR,M.FAMILIA, O.CANPROY, O.CANAPRO, O.FechEnSol, Id_Orfa,M.CODINT, O.DESTINO"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    If DestinoTodos.Value = 0 Then
        SQLX$ = SQLX$ + " WHERE O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    End If
    SQLX$ = SQLX$ + " AND O.CanApro > 0.05"
    SQLX$ = SQLX$ + " ORDER BY O.FECHENSOL, O.ID_ORFA"
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
          CI1% = XVALO(!CODINT)
          FE% = CVINT(!FechEnSol)
          PLANOX$ = SAFETEXT$(!NPLANO)
          CODI$ = SAFETEXT$(!Codigo)
          DESCRIX$ = SAFETEXT$(!Descripcion)
          OCOMPRAZ$ = SAFETEXT$(!ID_ORFA)
          CANTI = XVALO(!CANAPRO)
            CODIFAM$ = SAFETEXT$(!familia)
          OBSERVAX$ = ECOARCH$("TAFAMIL", CODIFAM$)
          NORFA$ = SAFETEXT$(!idsubor)
          DESTINO$ = SAFETEXT$(!DESTINO)
          J& = J& + 1
          MSF0.Rows = J& + 1
          MSF0.TextMatrix(J&, 1) = CI1%
          MSF0.TextMatrix(J&, 2) = FECHATEX$(FE%)
          MSF0.TextMatrix(J&, 3) = PLANOX$
          MSF0.TextMatrix(J&, 4) = DESCRIX$
          MSF0.TextMatrix(J&, 5) = OCOMPRAZ$
          MSF0.TextMatrix(J&, 6) = FORMATNUM$(CANTI, "F10.1")
          MSF0.TextMatrix(J&, 7) = OBSERVAX$
            FechaInforme% = CVINT(VALOBADA("INFOCALIDAD", "FechaInforme", "DocumentoOrigen='" & NORFA$ & "'"))
          MSF0.TextMatrix(J&, 8) = TRIM$(FECHATEX$(FechaInforme%))
          MSF0.TextMatrix(J&, 9) = "C.CALIDAD"
          If DestinoTodos.Value = 1 Then
            MSF0.TextMatrix(J&, 10) = TRIM$(ECOARCH$("DESTFAB", DESTINO))
          End If
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing

End Sub
Sub CARGAMSF3()
    MSF3.Rows = 1
    If DestinoTodos.Value = 0 Then
        Campos$ = "CODINT/F0;Vencimiento/D9;Calculada/D9;N° Plano/A16;Descripción/A36;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;Conjunto/A16;Calidad/F8.1;Operación/A32;NumeroOrfa/A0"
        Call CARGA_ENCABEZADO(Me.MSF3, Campos$, Me)
    Else
        Campos$ = "CODINT/F0;Vencimiento/D9;Calculada/D9;N° Plano/A16;Descripción/A36;O-C N°/A12;Cantidad/F8.1;Observaciones/A24;Conjunto/A16;Calidad/F8.1;Operación/A32;NumeroOrfa/A0;Destino/A24"
        Call CARGA_ENCABEZADO(Me.MSF3, Campos$, Me)
    End If
    SQLX$ = " SELECT M.DESCRIDOS, M.DESCRIPCION,M.NPLANO ,M.Codigo, O.IDSUBOR,M.FAMILIA, O.CANPROY, O.CANAPRO, O.FechEnSol, O.Id_Orfa, O.FechProyEntre, M.CODINT, O.DESTINO, O.StatuOrf"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    If DestinoTodos.Value = 0 Then
        SQLX$ = SQLX$ + " WHERE  O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    End If
    SQLX$ = SQLX$ + " AND (O.StatuOrf <= 3 AND O.CanApro < (O.CanProy - 0.005)) "
    SQLX$ = SQLX$ + " ORDER BY O.FECHENSOL, O.ID_ORFA"

    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
          NORFA$ = SAFETEXT$(!idsubor)
          CI1% = XVALO(!CODINT)
          FE% = CVINT(!FechEnSol)
          FechaCargaMaquina% = CVINT(!FechProyEntre)
          PLANOX$ = SAFETEXT$(!NPLANO)
          CODI$ = SAFETEXT$(!Codigo)
          DESCRIX$ = SAFETEXT$(!DESCRIDOS)
          OCOMPRAZ$ = SAFETEXT$(!ID_ORFA)
          CANTI = XVALO(!CANPROY) - XVALO(!CANAPRO)
          OBSERVAX$ = SAFETEXT$(!Descripcion)
          DESTINO$ = SAFETEXT$(!DESTINO)
          J& = J& + 1
          MSF3.Rows = J& + 1
          MSF3.TextMatrix(J&, 1) = CI1%
          MSF3.TextMatrix(J&, 2) = FECHATEX$(FE%)
          MSF3.TextMatrix(J&, 3) = FECHATEX$(FechaCargaMaquina%)
          MSF3.TextMatrix(J&, 4) = CODI$
          MSF3.TextMatrix(J&, 5) = DESCRIX$
          MSF3.TextMatrix(J&, 6) = OCOMPRAZ$
          MSF3.TextMatrix(J&, 7) = FORMATNUM$(CANTI, "F10.1")
          MSF3.TextMatrix(J&, 8) = OBSERVAX$
          MSF3.TextMatrix(J&, 9) = PLANOX$
             PendienteCalidad# = XVALO(VALOBADA("MOVITRANSTO", "SUM(CANTINGRE)", "DOPRILAR='" & NORFA$ & "'"))
          MSF3.TextMatrix(J&, 10) = TRIM$(FORMATNUM$(PendienteCalidad#, "F12.1"))
             NOPER% = OPERACIONPENDIENTE%(NORFA$)
             CONDI$ = " CodIConj = " & CStr(CI1%) & " AND NumeOper =" & CStr(NOPER%)
             Descrioper$ = TRIM$(SAFETEXT$(VALOBADA("SECOPER", "descoper", CONDI$, "NOMESS")))
             If Descrioper$ = "" Then Descrioper$ = "C.CALIDAD"
          MSF3.TextMatrix(J&, 11) = Descrioper$
          MSF3.TextMatrix(J&, 12) = NORFA$
          If DestinoTodos.Value = 1 Then
            MSF3.TextMatrix(J&, 13) = TRIM$(ECOARCH$("DESTFAB", DESTINO$))
          End If
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing

End Sub
Sub CARGAMSF1()

    MSF1.Rows = 1: J& = 0
    SQLX$ = " SELECT M.FAMILIA, SUM(CANAPRO) AS CANP"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    If DestinoTodos.Value = 0 Then
        SQLX$ = SQLX$ + " WHERE O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    End If
    SQLX$ = SQLX$ + " AND O.CanApro > 0.05"
    SQLX$ = SQLX$ + " GROUP BY M.FAMILIA"
    SQLX$ = SQLX$ + " ORDER BY M.FAMILIA"
    Set rs = READSET(SQLX$)
    With rs
'         CAMPOS$ = "CODINT/F0;Vencimiento/D11;N° Plano/A18;Descripción/A24;O-C N°/A12;Cant./F8.1;Observaciones/A32;OK/D8;Operador/A18"

        Do While Not .EOF
          TAFAM$ = SAFETEXT$(!familia)
          FAMIX$ = ECOARCH$("TAFAMIL", TAFAM$)
          
          CANTI = XVALO(!CANP)
          J& = J& + 1
          MSF1.Rows = J& + 1
          MSF1.TextMatrix(J&, 1) = FAMIX$
          MSF1.TextMatrix(J&, 2) = FORMATNUM$(CANTI, "F12.1")
          MSF1.TextMatrix(J&, 3) = TAFAM$
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub
Sub CARGAMSF2()
    'MUESTRA LOS DATOS AGRUPADOS POR FAMILIA EN ESTE CASO SOLO LOS PENDIENTES.
    MSF2.Rows = 1: J& = 0
    SQLX$ = " SELECT M.FAMILIA, SUM(CANPROY)AS CANPROY, SUM(CANAPRO) AS CANPRO"
    SQLX$ = SQLX$ + " FROM MASTER M WITH(NOLOCK) INNER JOIN ORDEFABR O WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ON M.CODINT = O.COD_INTE"
    If DestinoTodos.Value = 0 Then
        SQLX$ = SQLX$ + " WHERE  O.DESTINO ='" & TRIM$(Me.TXTDESTINO) & "'"
    End If
    SQLX$ = SQLX$ + " AND O.StatuOrf < 3 AND O.CanApro < (O.CanProy - 0.005) "
    SQLX$ = SQLX$ + " GROUP BY M.FAMILIA"
    SQLX$ = SQLX$ + " ORDER BY M.FAMILIA"
    Set rs = READSET(SQLX$)
    With rs
        Do While Not .EOF
          TAFAM$ = SAFETEXT$(!familia)
          
          FAMIX$ = ECOARCH$("TAFAMIL", TAFAM$)
          CANTI = XVALO(!CANPROY) - XVALO(!CANPRO)
          If CANTI < 0 Then CANTI = 0
          J& = J& + 1
          MSF2.Rows = J& + 1
          MSF2.TextMatrix(J&, 1) = FAMIX$
          MSF2.TextMatrix(J&, 2) = FORMATNUM$(CANTI, "F12.1")
          MSF2.TextMatrix(J&, 3) = TAFAM$
          .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub

