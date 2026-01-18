VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SHOWREP13 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Composición de Deuda"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   600
      Left            =   10680
      Picture         =   "SHOWREP13.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Emite Listado"
      Top             =   6600
      Width           =   700
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   820
      Left            =   120
      TabIndex        =   5
      Top             =   7320
      Width           =   10515
      Begin VB.TextBox TXTBUSCAR 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   400
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
         Height          =   315
         Left            =   3000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   400
         Width           =   2055
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
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
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   210
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   210
         Width           =   705
      End
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
      OleObjectBlob   =   "SHOWREP13.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "SHOWREP13.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   6645
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   480
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   6645
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   600
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   11721
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Titulo de la Consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   720
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
End
Attribute VB_Name = "SHOWREP13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
      Label5.FontName = GRID.CellFontName
      Label5.FontSize = GRID.CellFontSize
      Label5.FontBold = GRID.CellFontBold
      '
      If GRID.Rows > 1 Then
      Call PRINREP13.CARGALIMPRE
      If PRINREP13.Visible = True Then
         PRINREP13.Hide
      End If
      PRINREP13.Show 1
   End If

99 Command1.Enabled = True

End Sub


Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
   GRID.Cols = 5
   For I = 1 To 100
      GRID.Rows = I + 1
      GRID.TextMatrix(I, 1) = "AAAA" & I
      GRID.TextMatrix(I, 2) = "BBBB" & I
      GRID.TextMatrix(I, 3) = "CCCC" & I
      GRID.TextMatrix(I, 4) = "DDDD" & I
   Next I
   
   'EN LA FUNCION DE CARGAR LOS ENCABEZADOS HABRIA QUE ASIGNAR LOS ANCHOS DE LABEL 5 Y ASIGNARSELO A LABEL8
   For U& = 1 To GRID.Cols
          JI& = JI& + 1
          ' PROVISORIAMENTE A VER SI FUNCIONA
          'AQUI AGREGAR A LABEL5 LA CANTIDAD DE CARACTERES
          ANCH& = 20
          Label5.Caption = FORMATNUM$("9999999.9999999", "A25" + "*")
          GRID.ColWidth(JI& - 1) = Label5.Width
25        If JI& >= Label8.Count Then
             NUPIC% = Label8.Count
             Load Label8(NUPIC%)
             GoTo 25
          End If
          Label8(U&) = Label5.Width
   Next U&
   '
   Label7 = 100
   '
End Sub

Private Sub Label13_Change()
   Call ACTUTOTA2
End Sub

Private Sub Label5_Click()
   Call ACTUTOTA2
End Sub


Private Sub Label9_Change()
    Call ACTUTOTA1
End Sub
Sub ACTUTOTA1()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) + XVALO(Label5), "F12.2")
End Sub
Sub ACTUTOTA2()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) - XVALO(Label5), "F12.2")
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
End Sub

Private Sub MSF1_Click()
    Static MODO  As Boolean
    If (MSF1.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF1.Col = MSF1.MouseCol
            MSF1.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF1.Col = MSF1.MouseCol
            MSF1.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.Col = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.Col = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    I& = MSF1.MouseRow
    j& = MSF1.MouseCol
    If I& = 0 Or I& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = I&
    MSF1.Col = j&
    On Error Resume Next

    NPX& = TRIM$(MSF1.TextMatrix(I&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub
Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    I& = MSF_2.MouseRow
    j& = MSF_2.MouseCol
    If I& = 0 Or I& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.Row = I&
    MSF_2.Col = j&
    On Error Resume Next

    NPX& = TRIM$(MSF_2.TextMatrix(I&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub MSF2_Click()
    Static MODO  As Boolean
    If (MSF2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF2.Col = MSF2.MouseCol
            MSF2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF2.Col = MSF2.MouseCol
            MSF2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    I& = MSF2.MouseRow
    j& = MSF2.MouseCol
    If I& = 0 Or I& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = I&
    MSF2.Col = j&
    On Error Resume Next

    NPX& = TRIM$(MSF2.TextMatrix(I&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub


Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      MSF1.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = MSF1.Cols
    MSF_2.Redraw = False
    MSF1.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To MSF1.Rows - 1
      For j& = 1 To MSF1.Cols - 1
        If (InStr(UCase$(MSF1.TextMatrix(I&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF1.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = MSF1.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next I&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
