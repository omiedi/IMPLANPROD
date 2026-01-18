VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form VENTABUX 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción AHP - Valores expresados en Dolares"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   17040
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   17040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton BOTON 
      Caption         =   "."
      Height          =   300
      Left            =   4920
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.CommandButton BOTERASE 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   16340
      TabIndex        =   14
      ToolTipText     =   "Elimina Registro Activo"
      Top             =   840
      Width           =   660
   End
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   16320
      Picture         =   "VENTABUX.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aprobar"
      Top             =   4440
      Width           =   660
   End
   Begin VB.CommandButton BOTEXIT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   16340
      Picture         =   "VENTABUX.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cierra Formulario "
      Top             =   120
      Width           =   660
   End
   Begin VB.TextBox TXTMOVIL 
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exp. a Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7800
      TabIndex        =   7
      Top             =   7320
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Maximizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7920
      Picture         =   "VENTABUX.frx":0614
      TabIndex        =   6
      ToolTipText     =   "Emite Listado"
      Top             =   7320
      Visible         =   0   'False
      Width           =   2025
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
      OleObjectBlob   =   "VENTABUX.frx":091E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "VENTABUX.frx":0B52
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4605
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
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
   Begin MSFlexGridLib.MSFlexGrid MSFFORMULA 
      Height          =   1245
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar / Doble Click Para Seleccionar"
      Top             =   5160
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   2196
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFTOTALES 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      ToolTipText     =   "Encabezado"
      Top             =   4680
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   450
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      Appearance      =   0
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
   Begin VB.Label AGREGAFILASNUEVAS 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label EDITABLES 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   4
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
Attribute VB_Name = "VENTABUX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APROBADO%
Dim UNICAEDITABLE%
Dim ACTIVOX%
Sub ACTUGRILLA()
   If MSFFORMULA.Rows > 1 Then
      For i& = 1 To MSFFORMULA.Cols - 1
          Formulax$ = TRIM$(MSFFORMULA.TextMatrix(1, i&))
          If Formulax$ <> "" And Left(Formulax$, 1) <> "=" Then
             Call TEXTOLOTES$(Formulax$, ";")
             On Error Resume Next
             SIMB$ = CADENATEX$(2)
             COLU1& = CADENATEX$(1)
             COLU2& = CADENATEX$(3)
             For J& = 1 To MSF.Rows - 1
              Select Case SIMB$
                  Case "+"
                    MSF.TextMatrix(J&, i&) = XVALO(MSF.TextMatrix(J&, COLU1&)) + XVALO(MSF.TextMatrix(J&, COLU2&))
                  Case "-"
                    MSF.TextMatrix(J&, i&) = XVALO(MSF.TextMatrix(J&, COLU1&)) - XVALO(MSF.TextMatrix(J&, COLU2&))
                  Case "*"
                    MSF.TextMatrix(J&, i&) = XVALO(MSF.TextMatrix(J&, COLU1&)) * XVALO(MSF.TextMatrix(J&, COLU2&))
                  Case "/"
                    If XVALO(MSF.TextMatrix(J&, COLU2&)) <> 0 Then
                       MSF.TextMatrix(J&, i&) = XVALO(MSF.TextMatrix(J&, COLU1&)) / XVALO(MSF.TextMatrix(J&, COLU2&))
                    Else
                       MSF.TextMatrix(J&, i&) = 0
                    End If
                 End Select
               Next J&
          ElseIf Left(Formulax$, 1) = "=" Then
             Call TEXTOLOTES$(Formulax$, ";")
             On Error Resume Next
             SIMB$ = CADENATEX$(2)
             COLU1& = CADENATEX$(1)
             COLU2& = CADENATEX$(3)
             On Error GoTo 0
             Call EJECUTAR_FUNCIONES(Formulax$, i&)
          End If
      Next i&
   End If
   If MSFTOTALES.Rows > 1 Then
     For i& = 1 To MSFTOTALES.Cols - 1
      TOTAX$ = TRIM$(MSFTOTALES.TextMatrix(1, i&))
          If TOTAX$ <> "" Then
            MSFTOTALES.TextMatrix(0, i&) = ""
            For J& = 1 To MSF.Rows - 1
               MSFTOTALES.TextMatrix(0, i&) = XVALO(MSFTOTALES.TextMatrix(0, i&)) + XVALO(MSF.TextMatrix(J&, i&))
            Next J&
          End If
      Next i&
      'MSFTOTALES.TextMatrix(0, i&) = FORMATNUM$(XVALO(MSFTOTALES.TextMatrix(0, i&)), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
      MSF.TextMatrix(0, i&) = VALORSEGFORMATO(XVALO(MSFTOTALES.TextMatrix(0, i&)), Me.MSFFORMULA, 2, COLU1&)
   End If
End Sub
Sub EJECUTAR_FUNCIONES(Formulax$, COLUMNARESULTADO&)
      Dim VALOX()
      Call TEXTOLOTES$(Formulax$, ";")
      
      FUNXION$ = Mid$(CADENATEX$(1), 2)
      
      Select Case FUNXION$
         Case "PORCENT_ENTRE_2_VALORES"
            COLU1& = XVALO(CADENATEX$(2))
            COLU2& = XVALO(CADENATEX$(3))
            For J& = 1 To MSF.Rows - 1
                MSF.TextMatrix(J&, COLUMNARESULTADO&) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(XVALO(MSF.TextMatrix(J&, COLU1&)), XVALO(MSF.TextMatrix(J&, COLU2&))), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
            Next J&
      End Select

End Sub

Function VALORSEGFORMATO(DATOX$, MSF As msFlexGrid, FILACONFIG&, COLU1&)
   'SI ES FORMATO F REALIZA EL FORMATNUM$
   'SI ES OTRO FORMATO LO DEJA COMO ESTA.
   VALORSEGFORMATO = DATOX$
   FORMATX$ = FormatoSegunCampo$(MSF, FILACONFIG&, COLU1&)
   If UCase$(Left$(FORMATX$, 1)) = "F" Then
        VALORSEGFORMATO = FORMATNUM$(XVALO(DATOX$), FORMATX$)
   End If
End Function

Private Sub BOTERASE_Click()
   Call cmdEliminaItem(Me.MSF, MSF.Row)
   TXTMOVIL.Visible = False
   BOTON.Visible = False
End Sub

Private Sub BOTEXIT_Click()
   Me.APROBADO = 0
   ACTIVOX% = 0
   Unload Me
End Sub

Private Sub BOTON_Click()
     'POSICION DE LA CAJA DE TEXTO EN LA GRILLA
     BOTON.Enabled = False
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     '
     FORMATO1$ = Me.MSFFORMULA.TextMatrix(2, COLU1&)
     FORMATO2$ = Me.MSFFORMULA.TextMatrix(2, COLU1& + 1)
     ENCA1$ = MSF.TextMatrix(0, COLU1&)
     ENCA2$ = MSF.TextMatrix(0, COLU1& + 1)
     Campos$ = ENCA1$ & "/" & FORMATO1$ & ";" & ENCA2$ & "/" & FORMATO2$
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
     'CARGO ENCABEZADO
     CAMPMAS1$ = MSF.TextMatrix(0, COLU1&)
     CAMPMAS2$ = MSF.TextMatrix(0, COLU1& + 1)
     
     If FRMZELECHO.Caption = "Selección de: " & ENCA1$ Then
        GoTo 80 'si no cambio muestra sin volver a cargar la tabla.
     End If
     FRMZELECHO.Caption = "Selección de " & ENCA1$
     
     J& = 0
     On Error GoTo 0
     'FILTRO DE CONSULTA EJ.: CODIGO,DESCRIPCION,MASTER
     ARGUMENBUS$ = TRIM$(MSFFORMULA.TextMatrix(3, COLU1&))
     If ARGUMENBUS$ <> "" Then
         Screen.MousePointer = 11
         FRMZELECHO.MSF2.Redraw = False
         Call TEXTOLOTES$(ARGUMENBUS$, ";")
         CAMPO1$ = CADENATEX$(1)
         campo2$ = CADENATEX$(2)
         TABLA$ = CADENATEX$(3)
         SQLX$ = "SELECT " & CAMPO1$ & " AS CAMP1, " & campo2$ & " AS CAMP2"
         SQLX$ = SQLX$ + " FROM " & TABLA$ & " WITH(NOLOCK)"
         SQLX$ = SQLX$ + " ORDER BY " & campo2$
         Set RS = READSET(SQLX$)
         With RS
           Do While Not .EOF
             J& = J& + 1
             FRMZELECHO.MSF2.Rows = J& + 1
             FRMZELECHO.MSF2.TextMatrix(J&, 1) = SAFETEXT$(!CAMP1)
             FRMZELECHO.MSF2.TextMatrix(J&, 2) = SAFETEXT$(!CAMP2)
             .MoveNext
            Loop
          End With
     End If
     '
     FRMZELECHO.MSF2.Redraw = True
     Screen.MousePointer = 1
'     If TRIM$(FRMZELECHO.TXTBUSCAR) <> "" Then
'        Call FRMZELECHO.TXTBUSCAR_Change
'     End If
     
80   FRMZELECHO.Show 1
     LECTURA1X$ = TRIM$(RESP_ZELE_VECT(1))
     If LECTURA1X$ = "" Then GoTo 99
     TXTMOVIL = LECTURA1X$
     'MSF.TextMatrix(REG1&, COLU1&) = LECTURA1X$
     Call TXTMOVIL_KeyPress(13)
99   BOTON.Enabled = True

End Sub

Private Sub BOTREC_Click()
    BOTREC.Enabled = False
    APROBADO% = 1
    Hide
99  BOTREC.Enabled = True

End Sub

Private Sub Command4_Click()
End Sub

Private Sub Form_Activate()
  If ACTIVOX% > 0 Then Exit Sub
  ACTIVOX% = 1
  If MSF.Rows > 1 Then
    Call ACTUGRILLA
    Call COLOREAR_CELDA_EDITAB_BLANCO(Me.MSF, Me.EDITABLES.Caption)

    'MSF.SetFocus
    UNICAEDITABLE% = 0
    COLU_EDIT$ = TRIM$(EDITABLES.Caption)
    If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
       Call TEXTOLOTES$(COLU_EDIT$, ";")
       On Error Resume Next
       COLU1& = XVALO(CADENATEX$(1))
       REG1& = XVALO(CADENATEX$(4))
       If UBound(CADENATEX$) = 1 Then UNICAEDITABLE% = 1
       '
       MSF.COL = COLU1&
       MSF.Row = 1

       With TXTMOVIL
          .ZOrder 0
          .Width = MSF.CellWidth
          .Left = MSF.CellLeft + MSF.Left
          .Top = MSF.CellTop + MSF.Top
          .TEXT = MSF.TEXT
          .Visible = True
          .ToolTipText = MSF.Row & "," & COLU1&
          Call AGREGABOTON(COLU1&, 1)
          .SetFocus
       End With
    End If
    
  End If

End Sub

Private Sub Form_Load()
   'PINTO DE BLANCO LAS CELDAS EDITABLES
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.APROBADO = 0
   ACTIVOX% = 0
End Sub


Private Sub Form_Resize()
   MSF.Width = Me.Width - 855
   Me.MSFTOTALES.Width = Me.Width - 855
   Me.BOTERASE.Left = Me.Width - 750
   Me.BOTEXIT.Left = Me.Width - 750
   Me.BOTREC.Left = Me.Width - 750
   
End Sub

Private Sub MSF_Click()
          COLUM = MSF.MouseCol
          TXTMOVIL.ToolTipText = MSF.Row & "," & COLUM
          
          COLU_EDIT$ = TRIM$(EDITABLES.Caption)
          If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
            Call TEXTOLOTES$(COLU_EDIT$, ";")
            On Error Resume Next
            For i& = 1 To UBound(CADENATEX$)
               COLU1& = XVALO(CADENATEX$(i&))
               If COLUM = COLU1& Then
                 
                 On Error GoTo 0
                 GoTo 50
               End If
            Next i&
            On Error GoTo 0
          End If
          'SI LLEGA AQUI ES POR QEU NO MACHEA CON NINGUNO DE LOS EDITABLES
'          If COLUM <> 4 And COLUM <> 10 Then
          TXTMOVIL.Visible = False
          Exit Sub
'          End If

50        MSF.COL = COLUM
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            .TEXT = MSF.TEXT
            .Visible = True
            .ToolTipText = MSF.Row & "," & COLUM
            Call TXTMOVIL_Change
            .SetFocus
          End With
          Call ACTUGRILLA

End Sub

Private Sub MSF_Scroll()
   If TXTMOVIL.Visible = True Then
     'TXTMOVIL.Left = MSF1.CellLeft
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF.TEXT = TXTMOVIL
     TXTMOVIL.Visible = False
   End If
'   ScrollActivo% = 1
'   MSFTOTALES.LeftCol = MSF.LeftCol
End Sub





Private Sub MSFTOTALES_Scroll()
   MSF.LeftCol = MSFTOTALES.LeftCol
End Sub


Private Sub TXTMOVIL_Change()
     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     ARGUMENBUS$ = TRIM$(MSFFORMULA.TextMatrix(3, COLU1&))
     If ARGUMENBUS$ <> "" Then
         Call TEXTOLOTES(ARGUMENBUS$, ";")
         CAMPO1$ = CADENATEX$(1)
         campo2$ = CADENATEX$(2)
         TABLA$ = CADENATEX$(3)
         CONDI$ = CAMPO1$ & " = '" & TXTMOVIL & "'"
         EXO$ = VALOBADA(TABLA$, campo2$, CONDI$, "NOMESS")
         MSF.TextMatrix(REG1&, COLU1& + 1) = EXO$
     End If
End Sub

Private Sub TXTMOVIL_DblClick()
     'para ver si es columna editable con lista de seleccion
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If TRIM$(MSFFORMULA.TextMatrix(3, COLU1&)) <> "" Then
       Call BOTON_Click
     End If
End Sub

'Private Sub Timer1_Timer()
''hay que verlo bien tambien tenia codigo en scroll
''''     'POSICION CUANDO SE AMPLIA UNA COLUMNA SE EMPAREJA CON EL MSF SUPERIOR
''''     For I& = 1 To MSF.Cols - 1
''''       MSFTOTALES.ColWidth(I&) = MSF.ColWidth(I&)
''''     Next I&
''''     '
''''     'MUESTRA EL SCROLL CUANDO OCUPA MAS DE LA POSICION VISIBLE
''''     columax = MSF.Cols - 1
''''     POS& = MSF.ColPos(columax)
''''     If POS& > MSF.Width - MSF.CellWidth Then
''''        MSFTOTALES.Height = 600
''''        Exit Sub
''''     End If
''''     If MSFTOTALES.Height > 375 Then MSFTOTALES.Height = 375
'
'End Sub

Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     COLU1& = XVALO(CADENATEX$(2))
     REG1& = XVALO(CADENATEX$(1))
     On Error GoTo 0
'     T_Espera ((0.2))
     Call AGREGABOTON(COLU1&, REG1&)
End Sub

Private Sub TXTMOVIL_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
     'valido si es columna editable con lista de seleccion
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     If TRIM$(MSFFORMULA.TextMatrix(3, COLU1&)) <> "" Then
       Call BOTON_Click
     End If

   End If
End Sub

Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   BOTON.Visible = False
   
   NUEVAFILA% = XVALO(AGREGAFILASNUEVAS) 'label PARA ASIGANR SI SUMA FILAS
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
'     If XVALO(TXTMOVIL) > XVALO(MSF.TextMatrix(MSF.Row, 12)) Then
'       Call COMUNI("Cantidad Ingresada Supera la Disponible, Verifique!!!!")
'       TXTMOVIL = ""
'     Else
      If TXTMOVIL.Visible Then
        'MSF.TextMatrix(REG&, COLU1&) = TRIM$(FORMATNUM$(XVALO(TXTMOVIL), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&, VALOR)))
         MSF.TextMatrix(REG&, COLU1&) = VALORSEGFORMATO(TXTMOVIL, Me.MSFFORMULA, 2, COLU1&)
         TXTMOVIL.Visible = False
         On Error Resume Next
         Command1.SetFocus
         On Error GoTo 0
       End If
       Call ACTUGRILLA
       
       ColumActual& = COLU1&
       COLU_EDIT$ = TRIM$(EDITABLES.Caption)
       If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
         Call TEXTOLOTES$(COLU_EDIT$, ";")
         On Error Resume Next
         UltCOLEditable& = XVALO(CADENATEX$(UBound(CADENATEX$)))

         If ColumActual& < UltCOLEditable& Then
           For i& = 1 To UBound(CADENATEX$)
               COLU1& = XVALO(CADENATEX$(i&))
               If COLU1& > ColumActual& Then
                 On Error GoTo 0
                 MSFTOTALES.SetFocus
                 MSF.COL = COLU1&
                 MSF.Row = REG&
                 With TXTMOVIL
                  .ToolTipText = MSF.Row & "," & COLU1&
                  .ZOrder 0
                  .Width = MSF.CellWidth
                  .Left = MSF.CellLeft + MSF.Left
                  .Top = MSF.CellTop + MSF.Top
                  Call AGREGABOTON(COLU1&, REG&)
                   .TEXT = TRIM$(MSF.TEXT)
                  .Visible = True
                  Call PINTATEXT(TXTMOVIL)
                  .SetFocus
                  Exit Sub
                 End With
               End If
           Next i&
         Else
            'ACA ES CUANDO LLEGO AL FINAL DE LA FILA
            If MSF.Row = MSF.Rows - 1 Then  'SI ES LA ULTIMA FILA Y VARIABLE NUEVAVILA% > 0
             If NUEVAFILA% > 0 Then
              MSF.Rows = MSF.Rows + 1
              MSF.Row = MSF.Rows - 1
              MSF.TextMatrix(MSF.Row, 1) = ""
              ' por si la última fila no está visible
'              MSF.TopRow = MSF.Rows - 8
'              MSF.LeftCol = 1
             End If
            End If
            For i& = 1 To UBound(CADENATEX$) 'To 1 Step -1
               COLU1& = XVALO(CADENATEX$(i&))
               If COLU1& < ColumActual& Or UNICAEDITABLE% = 1 Then
                 On Error GoTo 0
                 'SE VE SI LLEGO AL FINAL DE LA GRILLA
                 If MSF.Rows - 1 >= MSF.Row Then
                      REG& = REG& + 1
                      If MSF.Rows - 1 = MSF.Row Then 'si llego al fin valida si esta configurado para que agregue nuevas filas
                        If NUEVAFILA% < 1 Then REG& = 1
                      End If
                 Else
                      If NUEVAFILA% < 1 Then
                        COLU1& = CADENATEX$(1)
                        REG& = 1
                      End If
                 End If
                 
                 MSF.COL = COLU1&
'                 If MSF.Rows <= MSF.Row Then MSF.Rows = MSF.Rows + 1
                 MSF.Row = REG&
                 With TXTMOVIL
                  .ToolTipText = MSF.Row & "," & COLU1&
                  .ZOrder 0
                  .Width = MSF.CellWidth
                  .Left = MSF.CellLeft + MSF.Left
                  .Top = MSF.CellTop + MSF.Top
                   .TEXT = TRIM$(MSF.TEXT)
                  .Visible = True
                  Call AGREGABOTON(COLU1&, REG&)
                  Call PINTATEXT(TXTMOVIL)
                  Call COLOREAR_CELDA_EDITAB_BLANCO(Me.MSF, Me.EDITABLES.Caption)
                  .SetFocus
                  Exit Sub
                 End With
               End If
             Next i&
           End If
          On Error GoTo 0
       End If
   End If

End Sub
Private Sub TXTMOVIL_LostFocus()
     If BOTON.Enabled = False Then Exit Sub
     If TXTMOVIL.Visible = True Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
        MSF.TextMatrix(REG&, COLU1&) = VALORSEGFORMATO(TXTMOVIL.TEXT, Me.MSFFORMULA, 2, COLU1&)
     End If
'     TXTMOVIL.Visible = False
     Call ACTUGRILLA
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
    MSF.COL = COLU1&
    Select Case KeyCode
        Case 38 '(38 flecha arriba)
            On Error Resume Next
            Call PINTATEXT(TXTMOVIL)
            TXTMOVIL.ToolTipText = MSF.Row & "," & COLU1&
            MSF.TEXT = TXTMOVIL.TEXT

            'MSF.SetFocus
            If MSF.Row > 1 Then MSF.Row = MSF.Row - 1
          MSF.COL = COLU1&
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
            
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            .SetFocus
          End With
'            Call VERIFICA_CANTI(MSF.Row - 1)
             Call ACTUGRILLA
            On Error GoTo 0
        Case 40 '(flecha 40 abajo)
            On Error Resume Next
            'MSF.SetFocus
            Call PINTATEXT(TXTMOVIL)
            TXTMOVIL.ToolTipText = MSF.Row & "," & COLU1&
            MSF.TEXT = TXTMOVIL.TEXT

            If MSF.Row < MSF.Rows Then MSF.Row = MSF.Row + 1
          MSF.COL = COLU1&

          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
            
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            .SetFocus
          End With
'          Call VERIFICA_CANTI(MSF.Row - 1)
          Call ACTUGRILLA
          On Error GoTo 0
    End Select

    '
End Sub
Private Sub Command1_Click()
'   Command1.Enabled = False
'   '
'
'   CommonDialog1.CancelError = True
'   CommonDialog1.DialogTitle = "Exportar a Excel"
'       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
'   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
'   CommonDialog1.FilterIndex = 1
'   On Error GoTo 99
'   CommonDialog1.ShowOpen
'   On Error GoTo 0
'   '
'   DoEvents
'   Screen.MousePointer = 11
'   AREXPO$ = UCase$(CommonDialog1.FileName)
'   If TRIM$(AREXPO$) <> "" Then
'     If MSF_2.Visible = True Then
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
'     Else
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.GRID)
'     End If
'   End If
   '

99 Command1.Enabled = True

End Sub
Private Sub Command3_Click()
    Command3.Enabled = False
    APROBADO% = 1
    Hide
99  Command3.Enabled = True
End Sub
Sub AGREGABOTON(ColumActual&, REG1&)
 BOTON.Visible = False
 
 CAMBIO = ColumActual& & REG1&
 If TRIM$(MSFFORMULA.TextMatrix(3, ColumActual&)) <> "" Then
    BOTON.Visible = True
    BOTON.ZOrder 0
    BOTON.Height = TXTMOVIL.Height
    BOTON.Left = TXTMOVIL.Left + (TXTMOVIL.Width - BOTON.Width)
    BOTON.Top = TXTMOVIL.Top
  End If

End Sub
                 
