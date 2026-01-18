VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRMZELECHO 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msf2 
      Height          =   3735
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   -2147483645
      ForeColorSel    =   16777215
      AllowBigSelection=   -1  'True
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2640
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   240
      Top             =   0
   End
   Begin VB.TextBox TXTBUSCAR 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   3820
      Width           =   6600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Bindings        =   "FRMZELECHO.frx":0000
      Height          =   3735
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   -2147483645
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "FRMZELECHO.frx":0014
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   255
      Index           =   0
      Left            =   240
      OleObjectBlob   =   "FRMZELECHO.frx":0248
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.Label TITULO 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label LBLANCHO 
      AutoSize        =   -1  'True
      Caption         =   "                "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExpExcel 
         Caption         =   "Exportar a Excel"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FRMZELECHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODITABLA$
Public RESP_ZELE$
Public NOCERRAR%
Public IMPRIMIR%
Public EXPORTAREXCEL%


Sub Buscar()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF1.Visible = False
      MSF2.Visible = True
      Exit Sub
    End If
    '
    MSF1.Rows = 1
    MSF1.Visible = True
    MSF1.ZOrder 0
    MSF1.Cols = MSF2.Cols
    MSF1.Redraw = False
    MSF2.Visible = False
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
    For i& = 1 To MSF2.Rows - 1
      For J& = 1 To MSF2.Cols - 1
        If InStr(UCase$(MSF2.TextMatrix(i&, J&)), ABUS$) Then
50        REG& = REG& + 1
          MSF1.Rows = MSF1.Rows + 1
          For h& = 1 To MSF2.Cols - 1
            MSF1.TextMatrix(REG&, h&) = MSF2.TextMatrix(i&, h&)
          Next h&
          Exit For
        End If

      Next J&
    Next i&
    '
    'MODIFICACION REALIZADA 20/12/18 SI HAY UNA BUSQUEDA SECUNDARIA RECOREE LOS REGISTROS YA FILTRADOS
    'Y VUELVE A FILTRAR POR EL SIGUIENTE DATO.
    REG& = 0
    If ABUSEC$ <> "" Then
        While REG& < MSF1.Rows - 1
          REG& = REG& + 1
          Encontro% = 0
          For J& = 1 To MSF1.Cols - 1
            If InStr(UCase$(MSF1.TextMatrix(REG&, J&)), ABUSEC$) Then
               Encontro% = 1
               Exit For
            End If
          Next J&
          If Encontro% < 1 Then
              Call cmdEliminaItem(MSF1, REG&)
              REG& = REG& - 1
          End If
        Wend
    End If

    MSF1.Redraw = True
    '
End Sub
    '



Private Sub Form_Activate()
    'If Me.Width > 12000 Then Me.Width = 12000
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
    RESP_ZELE$ = ""
    For i& = 0 To UBound(RESP_ZELE_VECT)
       RESP_ZELE_VECT(i&) = ""
    Next i&
    
    If Me.Caption <> TITULO Then
       TITULO.Caption = Me.Caption
       TXTBUSCAR = ""
    End If
    
    If TRIM$(TXTBUSCAR) <> "" Then
       Call Buscar ' 07/02/2017 PARA QUE REFRESQUE LA BUSQUEDA POR QUE SE QUEDABA VACIA LA GRILLA.
    End If
    IMPRIMIR% = 0
    

End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Cancel = 0 And NOCERRAR% > 0 Then
     Cancel = 1
     TXTBUSCAR.TEXT = ""
     Hide
  End If
End Sub

Private Sub Form_Resize()
   TXTBUSCAR.Width = FRMZELECHO.Width - 200
   MSF1.Width = FRMZELECHO.Width - 300
   MSF2.Width = FRMZELECHO.Width - 300
   TXTBUSCAR.Top = FRMZELECHO.Height - (TXTBUSCAR.Height + 935)
   
   MSF1.Height = FRMZELECHO.Height - 1400
   MSF2.Height = FRMZELECHO.Height - 1400
End Sub

Private Sub mnuExpExcel_Click()
   EXPORTAREXCEL% = 1
   Hide
End Sub

Private Sub mnuimprimir_Click()
   IMPRIMIR% = 1
   Hide
End Sub

Private Sub MSF1_Click()
'  If PINTAFILACLICK% > 0 Then
'   Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF1, vbYellow, MSF1.Row)
'  End If
End Sub

Private Sub MSF1_DblClick()
    MSF1.col = 0
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

Private Sub MSF2_Click()
   If PINTAFILACLICK% > 0 Then
     Call COLOREAR_GRILLA_SOLAFILA(FRMZELECHO.MSF2, vbYellow, MSF2.row)
   End If
End Sub

Private Sub MSF2_DblClick()
    If MSF2.Rows <= 1 Then Exit Sub
    MSF1.col = 0
    RESP_ZELE$ = MSF2.TEXT
    REG& = MSF2.MouseRow
    If REG& < 1 Then Exit Sub
    For i& = 1 To MSF2.Cols - 1
      RESP_ZELE_VECT(i&) = MSF2.TextMatrix(REG&, i&)
    Next i&
    Call T_Espera(1 / 5) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
                          ' por que cuando se oculta queda en click con el frm que esta debajo
    Hide
End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       RESP_ZELE$ = MSF2.TEXT
'       REG& = msf2.MouseRow
       REG& = MSF2.RowSel
       If REG& < 1 Then Exit Sub
       For i& = 1 To MSF2.Cols - 1
          RESP_ZELE_VECT(i&) = MSF2.TextMatrix(REG&, i&)
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
       On Error Resume Next
       TXTBUSCAR = Left$(TXTBUSCAR, Len(TXTBUSCAR) - 1)
       On Error GoTo 0
    End If
End Sub


Private Sub Timer1_Timer()
    Call PINTATEXT(TXTBUSCAR)
    ' LA RUTINA ANTERIOR LE AGREGA UNOS ESPACIOS VACIOS CUANDO LA CAJA DE TEXTO TXTBUSCAR ESTA VACIA.
    ' SI EL USUARIO HACE UN CLIC SOBRE TXTBUSCAR Y LUEGO INGRESA LO QUE ESTÁ BUSCANDO, DICHOS ESPACIOS
    ' FORMAN PARTE DE LA BUSQUEDA Y POR LO TANTO EL FILTRO NO DEVUELVE NADA. SI SE QUITAN LOS ESPACIOS
    ' AHI SI ENCUENTRA LO INGRESADO.
    If TRIM$(TXTBUSCAR.TEXT) = "" Then TXTBUSCAR.TEXT = ""
    Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()

    Static TextA$, FUEAMOSTRAR%
    '
    If TXTBUSCAR = TRIM$(TextA$) Then Exit Sub     ' NO CAMBIO EL CODIGO DE ARTICULO
    If FUEAMOSTRAR% > 0 Then Exit Sub   ' ESTÁ PRESENTANDO EL ANTERIOR
    '

    FUEAMOSTRAR% = 1
    Call Buscar
    FUEAMOSTRAR% = 0
    TextA$ = TXTBUSCAR
    
End Sub



 
 Sub TXTBUSCAR_Change()

End Sub

Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
    '
    
    If TXTBUSCAR.TEXT <> "" Then
        If KeyAscii = 13 Then
           If MSF2.Visible = True Then
            MSF2.SetFocus
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
          If MSF2.Visible = True Then
            On Error Resume Next
            MSF2.SetFocus
            If MSF2.row > 1 Then MSF2.row = MSF2.row - 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE SUBA UNA POSICION
            On Error GoTo 0
          ElseIf MSF1.Visible = True Then
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.row > 1 Then MSF1.row = MSF1.row - 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE SUBA UNA POSICION
            On Error GoTo 0
          End If
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
          If MSF2.Visible = True Then
            On Error Resume Next
            MSF2.SetFocus
            If MSF2.row < MSF2.Rows Then MSF2.row = MSF2.row + 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE BAJE UNA POSICION
            On Error GoTo 0
          ElseIf MSF1.Visible = True Then
            On Error Resume Next
            MSF1.SetFocus
            If MSF1.row < MSF1.Rows Then MSF1.row = MSF1.row + 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE BAJE UNA POSICION
            On Error GoTo 0
          End If
    End Select
    '
End Sub
