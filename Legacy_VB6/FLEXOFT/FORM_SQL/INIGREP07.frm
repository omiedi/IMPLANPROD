VERSION 5.00
Begin VB.Form INIGREP07 
   Caption         =   "Generador de Reportes"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillStyle       =   2  'Horizontal Line
      Height          =   3210
      Left            =   0
      Picture         =   "INIGREP07.frx":0000
      ScaleHeight     =   3150
      ScaleMode       =   0  'User
      ScaleWidth      =   2695.431
      TabIndex        =   9
      ToolTipText     =   "Doble Click para Generar Consultas"
      Top             =   0
      Width           =   5370
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Generador de Reportes "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   4530
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   360
      Top             =   3960
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   4725
      Picture         =   "INIGREP07.frx":AA3E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Búsqueda por Codigo o Descripcion"
      Top             =   3360
      Width           =   435
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   -1220
      TabIndex        =   4
      Top             =   3360
      Width           =   3855
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   1350
         ScaleHeight     =   75
         ScaleWidth      =   1005
         TabIndex        =   5
         Top             =   120
         Width           =   1065
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   3820
         X2              =   3820
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Consulta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   7
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   50
         Width           =   1065
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   735
      TabIndex        =   3
      Top             =   4515
      Width           =   3060
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -2420
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3360
      Width           =   7085
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   3780
      Width           =   5055
   End
   Begin VB.Label TASELE 
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   1890
      TabIndex        =   1
      Top             =   3990
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   -210
      X2              =   5300
      Y1              =   3210
      Y2              =   3210
   End
End
Attribute VB_Name = "INIGREP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Private Const CurrentUser As String = "HKEY_CURRENT_USER"

' VOLVER A CONECTAR REFERENIA EXCEL9.OLB
'
Private Sub Combo2_Click()
    RCONSU$ = Combo2.TEXT
    Call CARGA_CONSULTA(TRIM$(Mid$(Combo2.TEXT, 49)))
    Label9 = DESQUERY$
    TASELE = TABSELE$
    If TIPQUERY% = 0 Then
          Screen.MousePointer = 11
          SHOWREP07.Label3 = DESQUERY$
          SHOWREP07.Label2 = TRIM$(Mid$(Combo2.TEXT, 49))
          Screen.MousePointer = 1
          If SUBAPLI$ <> "" Then
             Call FINAL("")
             Exit Sub
          End If
          If SHOWREP07.Visible = True Then SHOWREP07.Hide
          SHOWREP07.Show 1
          '
          If APLINVO$ <> "" Then Call FINAL("")
          '
          If MODICONSU% = 1 Then
             MODICONSU% = 0
             Combo2.TEXT = RCONSU$
             FGENREP07.Label1.Caption = TRIM$(Mid$(RCONSU$, 49))
             FGENREP07.Show 1
          End If
       ElseIf TIPQUERY% = 1 Then
          SHOWPLA07.Label2 = TRIM$(Mid$(Combo2.TEXT, 49))
          SHOWPLA07.Label3 = DESQUERY$
          If IMPOSEJEC% < 1 Then SHOWPLA07.Show 1
          '
          If APLINVO$ <> "" Then Call FINAL("")
          '
          If MODICONSU% = 1 Then
             MODICONSU% = 0
             Combo2.TEXT = RCONSU$
             FGENPLA07.Label1.Caption = TRIM$(Mid$(RCONSU$, 49))
             FGENPLA07.Show 1
          End If
    End If
    Combo2.TEXT = RCONSU$
End Sub

Private Sub Combo2_DblClick()
    FGENREP07.Show 1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Combo2_Click
   End If
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    CAMPOS$ = "Código/A16;Descripcion/A64"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "Listado de Reportes"
    '
    SQLX$ = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY Asc "
    Set RS = READSET(SQLX$)
    With RS
      If Not .EOF Then
        Dim CONSULXA$()
        ReDim Preserve CONSULXA$(1)
        J& = 0
      End If
      Do While Not RS.EOF
        J& = J& + 1
        ReDim Preserve CONSULXA$(J&)
        CONSULXA$(J&) = SAFETEXT$(!COD_QUERY)
        FRMZELECHO.msf2.Rows = J& + 1
        FRMZELECHO.msf2.TextMatrix(J&, 1) = Left$(SAFETEXT$(!COD_QUERY), 16)
        FRMZELECHO.msf2.TextMatrix(J&, 2) = Left$(SAFETEXT$(!DES_QUERY), 48)
        .MoveNext
      Loop
    End With
    Call ABRETACON
    JJ& = 0
    '
    STRSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY Asc "
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para esta Instalacion.")
       MIRS.Close
       Set MIRS = Nothing
       GoTo 99
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
        CONSU$ = SAFETEXT$(MIRS!COD_QUERY)
        REPETIDO% = 0
        On Error Resume Next
        For i& = 1 To UBound(CONSULXA$)
          If UCase$(CONSULXA$(i&)) = CONSU$ Then
             REPETIDO% = 1
          End If
        Next i&
        On Error GoTo 0
        If REPETIDO% < 1 Then
          J& = J& + 1
          FRMZELECHO.msf2.Rows = J& + 1
          FRMZELECHO.msf2.TextMatrix(J&, 1) = Left$(SAFETEXT$(MIRS!COD_QUERY), 16)
          FRMZELECHO.msf2.TextMatrix(J&, 2) = Left$(SAFETEXT$(MIRS!DES_QUERY), 48)
        End If

        MIRS.MoveNext
    Loop
    MIRS.Close: Set MIRS = Nothing
    FRMZELECHO.msf2.COL = 1
    FRMZELECHO.msf2.Sort = 1
    FRMZELECHO.Width = 10000

    FRMZELECHO.Show 1
    
    RSEL$ = RESP_ZELE_VECT(1)
    If RSEL$ <> "" Then
        Combo2 = AJUSTI$(Mid$(RESP_ZELE_VECT(2), 1, 46), 48) + Left$(RESP_ZELE_VECT(1), 16)
       Call Combo2_Click
    End If
    
'    Call SETULTREG("!INDIASIE", JJ&)
'    Call FRESHECHO("!INDIASIE")
    '
'10  RSEL$ = REGSEL$("!INDIASIE/Consultas y Reportes Preconfigurados")
'    If Len(RSEL$) > 4 Then
'       Combo2 = AJUSTI$(Mid$(RSEL$, 17, 46), 48) + Left$(RSEL$, 16)
'       Call Combo2_Click
'    End If
    '
99  Command1.Enabled = True
End Sub

Private Sub Command1_NUEVO_Click()
    Command1.Enabled = False
    '
    JJ& = 0
    '
    CAMPOS$ = "Código/A20;Descripción Consulta/A64"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Caption = "Consultas"
    i& = 0
    SQLX$ = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY Asc "
    Set RS = READSET(SQLX$)
    Do While Not RS.EOF
       i& = i& + 1
       FRMZELECHO.msf2.Rows = i& + 1
       COCU$ = RS!COD_QUERY
       DECU$ = RS!DES_QUERY
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       FRMZELECHO.msf2.TextMatrix(i&, 1) = COCU$
       FRMZELECHO.msf2.TextMatrix(i&, 2) = DECU$
    MIRS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    FRMZELECHO.Show 1
    CODIX$ = TRIM$(RESP_ZELE_VECT(1))
    DESCRIX$ = TRIM$(RESP_ZELE_VECT(2))
    If CODIX$ <> "" Then
       Combo2 = AJUSTI$(Mid$(DESCRIX$, 1, 46), 48) + Left$(CODIX$, 16)
       Call Combo2_Click
    End If
    '
99  Command1.Enabled = True
End Sub




Private Sub Form_Load()
   '
   If App.PrevInstance = True Then
      Close: End
   End If
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call InstalaFuenteAnchoFijo
   '
'APLINVO$ = "VEPORCLIRAFE/VEPRI(01/01/06;31/12/09;1;12)"
   If APLINVO$ <> "" Then Height = 3570
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call ABRECONEXION
   KLIK% = 0
   '
   Call ABRETACON     ' crea la tabla de consultas
   Call VerConsul
   Call CREATABLAS_SQLCTRL
   '
   FILTX$ = "": SUBAPLI$ = "": LLAMOCONFILTROS% = 0
   'APLINVO$ = ""
   If APLINVO$ <> "" Then
      PPXX% = InStr(APLINVO$, "(")
      If PPXX% > 0 Then
         PPYY% = InStr(PPXX%, APLINVO$, ")")
         If PPYY% <= PPXX% Then PPYY% = 1 + Len(APLINVO$)
         FILTX$ = Mid$(APLINVO$, PPXX% + 1, PPYY% - PPXX% - 1)
         'SE AGREGO ESTA FILA DE ABAJO POR QUE SI LO PASO CON ';' LO TRUNCA EN VERJOBID
         FILTX$ = REPLACAR$(FILTX$, ",", ";")
         '
         APLINVO$ = Left$(APLINVO$, PPXX% - 1) + Mid$(APLINVO$, PPYY% + 1)
         LLAMOCONFILTROS% = 1
      End If
      PPXX% = InStr(APLINVO$, "/")
      If PPXX% > 0 Then
         SUBAPLI$ = Mid$(APLINVO$, PPXX% + 1)
         APLINVO$ = Left$(APLINVO$, PPXX% - 1)
      End If
      Screen.MousePointer = 11
      Timer1.Enabled = True
'      Combo2.TEXT = Space$(48) + APLINVO$
'      Call Combo2_KeyPress(13)
   End If
   '
End Sub

Sub VerConsul()
    '
    Combo2.Clear
    If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
        SQLX$ = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY ASC"
        Set RS = READSET(SQLX$)
        With RS
          If Not .EOF Then
            TTXX$ = AJUSTI$(SAFETEXT$(RS!DES_QUERY), 46) + "  " + SAFETEXT$(RS!COD_QUERY)
            Combo2.AddItem TTXX$
            RS.MoveNext
          End If
        End With
    End If
    RS.Close
    Set RS = Nothing
    
    Call ABRETACON
    
    
    STRSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then Exit Sub
    '
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          TTXX$ = AJUSTI$(SAFETEXT$(MIRS!DES_QUERY), 46) + "  " + SAFETEXT$(MIRS!COD_QUERY)
          Combo2.AddItem TTXX$
          MIRS.MoveNext
       Loop
    End If
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call COMPACTABD
    ESFINAL = 1
    Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ESFINAL% = 1: Cancel = 0
    If Cancel > 0 Then
       MsgBox "Cancel = " & Cancel & " - Imposible Descargar Formulario."
       Cancel = 0
    End If
    Call FINAL("")
End Sub

Private Sub Label8_Click()
  VENTADISE.Show 1
End Sub

Private Sub Picture1_DBLClick()
    If DERACCE%("CFGREPS", "Configuración de Reportes") > 1 Then
       DEFTREP07.Show 1
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    On Error Resume Next
    Me.ZOrder (0)
    AppActivate Caption
    SendKeys Chr$(0)
    On Error GoTo 0
    Combo2.TEXT = Space$(48) + APLINVO$
    Call Combo2_KeyPress(13)
End Sub
