VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMZELECHO2 
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2640
      Top             =   1920
   End
   Begin MSFlexGridLib.MSFlexGrid msf2 
      Bindings        =   "FRMZELECHO2.frx":0000
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16501405
      ForeColorSel    =   8421504
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Bindings        =   "FRMZELECHO2.frx":0014
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648447
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox TXTBUSCAR 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      TabIndex        =   0
      Top             =   3840
      Width           =   6825
   End
   Begin VB.Label LBLANCHO 
      AutoSize        =   -1  'True
      Caption         =   "                "
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FRMZELECHO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODITABLA$
Public RESP_ZELE$


Sub BUSCAR()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF1.Visible = False
      msf2.Visible = True
      Exit Sub
    End If
    '
    MSF1.Rows = 1
    MSF1.Visible = True
    MSF1.ZOrder 0
    MSF1.Cols = msf2.Cols
    MSF1.Redraw = False
    msf2.Visible = False
    '
    ABUS$ = UCase(TXTBUSCAR.TEXT)
    ABUSEC$ = ""
    PPXYZ% = InStr(ABUS$, "+")
    If PPXYZ% > 1 Then
       ABUSEC$ = UCase$(TRIM$(Mid$(ABUS$, PPXYZ% + 1)))
       ABUS$ = UCase$(TRIM$(Left$(ABUS$, PPXYZ% - 1)))
    End If

    '''
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN MSF1  LOS QUE MACHEAN
    For i& = 1 To msf2.Rows - 1
      For j& = 1 To msf2.Cols - 1
        If ABUSEC$ <> "" Then
          If InStr(UCase$(msf2.TextMatrix(i&, j&)), ABUS$) Or InStr(UCase$(msf2.TextMatrix(i&, j&)), ABUSEC$) Then GoTo 50
        End If
        If InStr(UCase$(msf2.TextMatrix(i&, j&)), ABUS$) Then
50        REG& = REG& + 1
          MSF1.Rows = MSF1.Rows + 1
          For H& = 1 To msf2.Cols - 1
            MSF1.TextMatrix(REG&, H&) = msf2.TextMatrix(i&, H&)
          Next H&
          Exit For
        End If
      Next j&
    Next i&
    '
    MSF1.Redraw = True
    '
End Sub
    '



Private Sub Form_Activate()
    If Me.Width > 12000 Then Me.Width = 12000
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
    RESP_ZELE$ = ""
    For i& = 0 To UBound(RESP_ZELE_VECT)
       RESP_ZELE_VECT(i&) = ""
    Next i&
End Sub

Private Sub Form_Load()
   If VENZELECHO.Inicializar = 1 Then
      'SI SE LE PASA NOMBRE DE ZELECHO PREVIAMENTE ARMADO BUSCA POR EL NOMBRE
      If TRIM$(VENZELECHO.NombreZelechoArmado) <> "" Then Call VENZELECHOFUN.ArmarZelechoArmado
   End If
End Sub
'''''Sub ArmarZelechoArmado()
'''''   If TRIM$(UCase$(VENZELECHO.NombreZelechoArmado)) = "DEUDOR12" Then
'''''        VENZELECHO.CamposCFormato = "Cta.Cliente/F8;Razon Social/A48"
'''''        VENZELECHO.Titulo = "Clientes Activos"
'''''        VENZELECHO.Condix = "STAT_CLI >=0"
'''''        VENZELECHO.Campox = "Nume_cli,Raso_Cli"
'''''        VENZELECHO.Tablax = "DEUDOR12"
'''''        VENZELECHO.Anchox = 10000
'''''   End If
'''''   Call ArmaConsultaRecord
'''''   Call CountRegMsf
'''''   Call CargarInfo
'''''
'''''End Sub
'''''Sub ArmaConsultaRecord()
'''''    Call CARGA_ENCABEZADO(Me.msf2, VENZELECHO.CamposCFormato, Me)
'''''    Call CARGA_ENCABEZADO(Me.MSF1, VENZELECHO.CamposCFormato, Me)
'''''    Me.Caption = VENZELECHO.Titulo
'''''    SQLX0$ = " SELECT " & VENZELECHO.Campox
'''''    SQLX1$ = " SELECT count(*) "
'''''    SQLX2$ = SQLX2$ + " FROM " & VENZELECHO.Tablax
'''''    If TRIM$(VENZELECHO.Condix) <> "" Then SQLX2$ = SQLX2$ + " WHERE " & VENZELECHO.Condix
'''''    VENZELECHO.SelectCompleta = SQLX0$ & SQLX2$
'''''    VENZELECHO.SelectCount = SQLX1$
'''''End Sub
'''''Sub CountRegMsf()
'''''   Set RS = READSET(VENZELECHO.SelectCount)
'''''   VENZELECHO.CantiReg = XVALO(RS.Fields(0))
'''''   RS.Close
'''''   Set RS = Nothing
'''''End Sub


''''''Sub CargarInfo()
''''''    Set RS = READSET(VENZELECHO.SelectCompleta)
''''''    With RS
''''''      j& = 0
''''''      Do While Not .EOF
''''''         j& = j& + 1
''''''         For i& = 1 To RS.Fields.Count
''''''            Me.msf2.TextMatrix(j&, i& - 1) = RS.Fields(i& - 1)
''''''         Next i&
''''''         .MoveNext
''''''      Loop
''''''    End With
''''''End Sub
Private Sub Form_Resize()
   TXTBUSCAR.Width = FRMZELECHO2.Width - 250
   MSF1.Width = FRMZELECHO2.Width - 250
   msf2.Width = FRMZELECHO2.Width - 250
End Sub

Private Sub MSF1_Click()
  If PINTAFILACLICK% > 0 Then

   Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF1, vbYellow, MSF1.Row)
  End If
End Sub

Private Sub MSF1_DblClick()
    MSF1.COL = 0
    RESP_ZELE$ = MSF1.TEXT

    RESP_ZELE$ = MSF1.TEXT
    REG& = MSF1.MouseRow
    If REG& < 1 Then Exit Sub
    For i& = 1 To MSF1.Cols - 1
      RESP_ZELE_VECT(i&) = MSF1.TextMatrix(REG&, i&)
    Next i&
    Call T_Espera(1 / 10) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
                          ' por que cuando se oculta queda en click con el frm que esta debajo
    
    Hide
    
End Sub

Private Sub msf2_Click()
   If PINTAFILACLICK% > 0 Then
     Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.msf2, vbYellow, msf2.Row)
   End If
End Sub

Private Sub MSF2_DblClick()
    MSF1.COL = 0
    RESP_ZELE$ = msf2.TEXT
    REG& = msf2.MouseRow
    If REG& < 1 Then Exit Sub
    For i& = 1 To msf2.Cols - 1
      RESP_ZELE_VECT(i&) = msf2.TextMatrix(REG&, i&)
    Next i&
    Call T_Espera(1 / 10) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
                          ' por que cuando se oculta queda en click con el frm que esta debajo
    Hide
End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       RESP_ZELE$ = msf2.TEXT
'       REG& = msf2.MouseRow
       REG& = msf2.RowSel
       If REG& < 1 Then Exit Sub
       For i& = 1 To msf2.Cols - 1
          RESP_ZELE_VECT(i&) = msf2.TextMatrix(REG&, i&)
       Next i&
       Hide
    ElseIf KeyAscii >= 32 And KeyAscii <= 126 Then
       On Error Resume Next
       TXTBUSCAR.SetFocus
       On Error GoTo 0
       'agrego el caracter por que si no solo toma el foco
       CARACTERX$ = Chr$(KeyAscii)
       TXTBUSCAR.TEXT = TXTBUSCAR.TEXT & CARACTERX$
       TXTBUSCAR.SelStart = Len(TXTBUSCAR)
    ElseIf KeyAscii = 8 Then
       TXTBUSCAR = Left$(TXTBUSCAR, Len(TXTBUSCAR) - 1)
    End If
End Sub
Private Sub MSF1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       RESP_ZELE$ = MSF1.TEXT
       'REG& = MSF1.MouseRow
       REG& = MSF1.RowSel
       If REG& < 1 Then Exit Sub
       For i& = 1 To MSF1.Cols - 1
          RESP_ZELE_VECT(i&) = MSF1.TextMatrix(REG&, i&)
       Next i&
       Hide
    ElseIf KeyAscii >= 32 And KeyAscii <= 126 Then
       On Error Resume Next
       TXTBUSCAR.SetFocus
       On Error GoTo 0
       'agrego el caracter por que si no solo toma el foco
       CARACTERX$ = Chr$(KeyAscii)
       TXTBUSCAR.TEXT = TXTBUSCAR.TEXT & CARACTERX$
       TXTBUSCAR.SelStart = Len(TXTBUSCAR)
    ElseIf KeyAscii = 8 Then
       TXTBUSCAR = Left$(TXTBUSCAR, Len(TXTBUSCAR) - 1)
    End If
End Sub


Private Sub Timer1_Timer()
    Call PINTATEXT(TXTBUSCAR)
    Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()
    Static TEXTA$, FUEAMOSTRAR%
    '
    If TXTBUSCAR = TEXTA$ Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '

    FUEAMOSTRAR% = 1
    Call BUSCAR
    FUEAMOSTRAR% = 0
    TEXTA$ = TXTBUSCAR
    
End Sub



Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
    '
    
    If TXTBUSCAR.TEXT <> "" Then
        If KeyAscii = 13 Then
           If msf2.Visible = True Then
            msf2.SetFocus
           ElseIf MSF1.Visible = True Then
            MSF1.SetFocus
           End If
        End If
    End If
    '
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
    Select Case KeyCode 'SI TIENE UN TEXTO DE BUSQUEDA BORRA
        Case 27 'pulso  escape
          If TRIM$(TXTBUSCAR.TEXT) <> "" Then
            TXTBUSCAR.TEXT = ""
          Else          'SI NO TIENE TEXTO OCULTA EL FORMULARIO
            Hide
          End If
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
          If msf2.Visible = True Then
            On Error Resume Next
            msf2.SetFocus
            If msf2.Row > 1 Then msf2.Row = msf2.Row - 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE SUBA UNA POSICION
            On Error GoTo 0
          ElseIf MSF1.Visible = True Then
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.Row > 1 Then MSF1.Row = MSF1.Row - 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE SUBA UNA POSICION
            On Error GoTo 0
          End If
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
          If msf2.Visible = True Then
            On Error Resume Next
            msf2.SetFocus
            If msf2.Row < msf2.Rows Then msf2.Row = msf2.Row + 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE BAJE UNA POSICION
            On Error GoTo 0
          ElseIf MSF1.Visible = True Then
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.Row < MSF1.Rows Then MSF1.Row = MSF1.Row + 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE BAJE UNA POSICION
            On Error GoTo 0
          End If
    End Select
    '
End Sub
