VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VENTABNEW 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción AHP - Valores expresados en Dolares"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   17265
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8040
      Top             =   2280
   End
   Begin VB.CommandButton CMDINSERTAR 
      Caption         =   "Insertar"
      Enabled         =   0   'False
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
      Left            =   16300
      TabIndex        =   17
      ToolTipText     =   "Inserta una Fila en Blanco en  laPosición Seleccionada"
      Top             =   3960
      Width           =   930
   End
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
      Enabled         =   0   'False
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
      Left            =   16320
      TabIndex        =   14
      ToolTipText     =   "Elimina Registro Activo"
      Top             =   840
      Width           =   930
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
      Left            =   16300
      Picture         =   "MSTRDATO_BASE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aprobar"
      Top             =   4440
      Width           =   950
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
      Left            =   16300
      Picture         =   "MSTRDATO_BASE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cierra Formulario "
      Top             =   120
      Width           =   930
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
      Top             =   240
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
      Visible         =   0   'False
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
      Picture         =   "MSTRDATO_BASE.frx":0614
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
      OleObjectBlob   =   "MSTRDATO_BASE.frx":091E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "MSTRDATO_BASE.frx":0B52
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4605
      Left            =   120
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
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar / Doble Click Para Seleccionar"
      Top             =   5280
      Visible         =   0   'False
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
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label EDITABLES 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
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
Attribute VB_Name = "VENTABNEW"
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
      On Error Resume Next
      MSF.TextMatrix(0, i&) = VALORSEGFORMATO(XVALO(MSFTOTALES.TextMatrix(0, i&)), Me.MSFFORMULA, 2, COLU1&)
      On Error GoTo 0
   End If
   For JJ% = 0 To MSF.Cols - 1
      J& = JJ%
      If VENTANA.Enumerar(JJ%) = "S" Then
         Call numerargrilla(Me.MSF, J&)
      End If
      If VENTANA.EnumerarRepUltimo(JJ%) = "S" Then
         Call numerargrilla(Me.MSF, J&, 0, 1)
      End If
      
   Next JJ%
   '
   
End Sub
Sub EJECUTAR_FUNCIONES(Formulax$, COLUMNARESULTADO&)
      Dim VALOX()
      Call TEXTOLOTES$(Formulax$, ";")
      
      FUNXION$ = Mid$(CADENATEX$(1), 2)
      CALCULOX$ = CADENATEX(2)
      Select Case FUNXION$
         Case "PORCENT_ENTRE_2_VALORES"
            COLU1& = XVALO(CADENATEX$(2))
            COLU2& = XVALO(CADENATEX$(3))
            For J& = 1 To MSF.Rows - 1
                MSF.TextMatrix(J&, COLUMNARESULTADO&) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(XVALO(MSF.TextMatrix(J&, COLU1&)), XVALO(MSF.TextMatrix(J&, COLU2&))), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
            Next J&
          Case "EXCEL"
            CALCULOX$ = CADENATEX(2)
            For J& = 1 To MSF.Rows - 1
                UBICAX$ = UBICACIONX1$(J&, COLUMNARESULTADO&)
                MSF.TextMatrix(J&, COLUMNARESULTADO&) = TRIM$(FORMATNUM$(VALORCELDA1#(UBICAX$, VENTABNEW.MSF, J&), "F16.2"))
                'MSF.TextMatrix(J&, COLUMNARESULTADO&) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(XVALO(MSF.TextMatrix(J&, COLU1&)), XVALO(MSF.TextMatrix(J&, COLU2&))), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
            Next J&

'             Dim COLUMNAX(16)
'             Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
'             REG& = XVALO(CADENATEX$(1))
'             Call TEXTOLOTES$(Formulax$, "(", ")")
'             '(12;/;11);*;(12/11);*;13
'             'For i = 1 To 16
'             If InStr(Formulax$, "(") > 0 Then
'               If InStr(Formulax$, ")") > 0 Then
'
'                  COLUMNAX(12) = 1
'               End If
'             End If
             
      End Select

End Sub

Function VALORSEGFORMATO(DATOX$, MSF As MSFlexGrid, FILACONFIG&, COLU1&)
   'SI ES FORMATO F REALIZA EL FORMATNUM$
   'SI ES OTRO FORMATO LO DEJA COMO ESTA.
   VALORSEGFORMATO = DATOX$
   FORMATX$ = FormatoSegunCampo$(MSF, FILACONFIG&, COLU1&)
   If UCase$(Left$(FORMATX$, 1)) = "F" Then
        VALORSEGFORMATO = FORMATNUM$(XVALO(DATOX$), FORMATX$)
   ElseIf UCase$(Left$(FORMATX$, 1)) = "H" Then
         If XVALO(DATOX$) > 0 Then
           VALORSEGFORMATO = HORATEX$(HORANUM(DATOX$))
           If XVALO(VALORSEGFORMATO) > 24 Or XVALO(VALORSEGFORMATO) <= 0 Then
              Call COMUNI("Horario No Válido")
           End If
         End If
   End If
   
End Function

Private Sub BOTERASE_Click()
   Call cmdEliminaItem(Me.MSF, MSF.Row)
   TXTMOVIL.Visible = False
   BOTON.Visible = False
   Call ACTUGRILLA

End Sub

Private Sub BOTEXIT_Click()
   Me.APROBADO = 0
   ACTIVOX% = 0
   Call T_Espera(1 / 10)
   Unload Me
End Sub

Private Sub BOTON_Click()
     'POSICION DE LA CAJA DE TEXTO EN LA GRILLA
     BOTON.Enabled = False
     TXTMOVIL_GotFocus
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     '
     FORMATO1$ = Me.MSFFORMULA.TextMatrix(2, COLU1&)
     LETFORM$ = UCase$(Left(TRIM$(FORMATO1$), 1))
     FORMATO2$ = Me.MSFFORMULA.TextMatrix(2, COLU1& + 1)
     ENCA1$ = Me.MSF.TextMatrix(0, COLU1&)
     ENCA2$ = Me.MSF.TextMatrix(0, COLU1& + 1)
     CAMPOS$ = ENCA1$ & "/" & FORMATO1$ & ";" & ENCA2$ & "/" & FORMATO2$
     Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
     Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
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
         FRMZELECHO.msf2.Redraw = False
         Call TEXTOLOTES$(ARGUMENBUS$, ";")
         CAMPO1$ = CADENATEX$(1)
         campo2$ = CADENATEX$(2)
         TABLA$ = CADENATEX$(3)
         On Error Resume Next
         CONDI$ = CADENATEX$(4)
         On Error GoTo 0

         COL% = Int(COLU1&)
         SKL$ = ""
         CONDI1$ = TRIM$(VENTANA.Condicion1(COL%))
         CONDI2$ = TRIM$(VENTANA.Condicion2(COL%))
         
         If CONDI$ <> "" Then
           SKL$ = " WHERE " & CONDI$
         End If
         If CONDI1$ <> "" Then
           If TRIM$(SKL$) = "" Then
             SKL$ = " WHERE "
           ElseIf TRIM$(CONDI$) <> "" Then
             SKL$ = " AND "
           End If
           SKL$ = SKL$ & CONDI1$
         End If
         'pasa por la funcion resulcondi2 que considera si el paramentro esta vacio devuelve la condicion vacia
         'o sea que si no tiene filtro en la celda muestra la lista sin filtrar
         If CONDI2$ <> "" Then CONDI2$ = RESULCONDI2$(CONDI2$, VENTABNEW.MSF, REG1&, LETFORM$)
         If CONDI2$ <> "" Then
           If TRIM$(SKL$) = "" Then
             SKL$ = " WHERE "
           ElseIf TRIM$(CONDI1$) <> "" Then
             SKL$ = " AND "
           End If
           SKL$ = SKL$ & CONDI2$
           'HAY QUE VER LAS POSICIONES PARA QEU SI CONDI2 ESTA VACIO QUE NO AGREGUE EL WHERE NI EL AND
         End If
          
         
         
         
         SQLX$ = "SELECT " & CAMPO1$ & " AS CAMP1, " & campo2$ & " AS CAMP2"
         SQLX$ = SQLX$ + " FROM " & TABLA$ & " WITH(NOLOCK)"
'         If TRIM$(CONDI$) <> "" Then SQLX$ = SQLX$ + SKL$
         SQLX$ = SQLX$ + SKL$
         SQLX$ = SQLX$ + " ORDER BY " & campo2$
         Set RS = READSET(SQLX$)
         With RS
           Do While Not .EOF
             J& = J& + 1
             FRMZELECHO.msf2.Rows = J& + 1
             FRMZELECHO.msf2.TextMatrix(J&, 1) = SAFETEXT$(!CAMP1)
             If LETFORM$ = "F" Then
                FRMZELECHO.msf2.TextMatrix(J&, 1) = FORMATNUM$(XVALO(SAFETEXT$(!CAMP1)), FORMATO1$)
             End If
             FRMZELECHO.msf2.TextMatrix(J&, 2) = SAFETEXT$(!CAMP2)
             .MoveNext
            Loop
          End With
     End If
     '
     FRMZELECHO.msf2.Redraw = True
     Screen.MousePointer = 1
     FRMZELECHO.msf2.COL = 1
     FRMZELECHO.msf2.Sort = 3
80   FRMZELECHO.Show 1
     LECTURA1X$ = TRIM$(RESP_ZELE_VECT(1))
     If LECTURA1X$ = "" Then GoTo 99
     TXTMOVIL = LECTURA1X$
     MSF.TextMatrix(REG1&, COLU1&) = LECTURA1X$
     Call TXTMOVIL_KeyPress(13)
99   BOTON.Enabled = True

End Sub
Function RESULCONDI2$(CONDI2$, MSF As MSFlexGrid, REG&, FORMATOX$)
   RESULCONDI2$ = CONDI2$
   POS1% = InStr(TRIM$(CONDI2$), "COL@")
   If POS1% > 0 Then
      COL1% = XVALO(Mid$(CONDI2$, POS1% + 4))
      If FORMATOX$ = "F" Then
         RESULCONDI2$ = Left$(CONDI2$, POS1% - 1) & MSF.TextMatrix(REG&, COL1%)
      Else
         RESULCONDI2$ = Left$(CONDI2$, POS1% - 1)
         If TRIM$(MSF.TextMatrix(REG&, COL1%)) = "" Then
           RESULCONDI2$ = "" ' SI ESTA VACIO EL PARAMETRO SALE PARA QUE NO CONSIDERE EL FILTRO
           Exit Function
         End If
         '" & MSF.TextMatrix(REG&, COL1%) & "'"
         RESULCONDI2$ = RESULCONDI2$ & Chr(39) & TRIM$(MSF.TextMatrix(REG&, COL1%)) & Chr(39)
      End If
   End If
   
End Function

Private Sub BOTREC_Click()
    BOTREC.Enabled = False
    Call T_Espera(1 / 10)
    A = FilaVacia%(Me.MSF, 1, 0) 'SI LA FILA ESTA VACIA LA BORRA ASI NO LA TRANSPORTA AL GRID DE ORIGEN
    If MSF.Rows - 1 > 0 Then
        For JJ% = 1 To MSF.Cols - 1
            If VENTANA.RepiteDato(JJ%) = "S" Then
                 DATO1$ = Me.MSF.TextMatrix(1, JJ%)
                 For i& = 1 To Me.MSF.Rows - 1
                   If FilaVacia%(Me.MSF, 1) < 1 Then
                     If TRIM$(Me.MSF.TextMatrix(i&, JJ%)) = "" Then Me.MSF.TextMatrix(i&, JJ%) = DATO1$
                   End If
                 Next i&
            End If
        Next JJ%
     End If
'    End If

    'POR DEFECTO PermitirFilasVacias ES 0 SI ES ASI LAS FILAS QUE EN LA PRIMERA Y SEGUNDA COLUMNA NO TIENEN DATOS LAS ELIMINA.
    'COMENTADO POR QUE NO TIENE DEMASIADO SENTIDO
'    If VENTANA.PermitirFilasVacias < 1 Then
'      CANTIX = MSF.Rows - 1
'      If CANTIX <= 1 Then GoTo 80
'      For i& = 1 To CANTIX
'        J& = J& + 1
'        If TRIM$(MSF.TextMatrix(J&, 1)) = "" Then
'           If TRIM$(MSF.TextMatrix(J&, 2)) = "" Then
'             Call cmdEliminaItem(Me.MSF, J&)
'             J& = J& - 1
'           End If
'        End If
'      Next i&
'    End If
80  APROBADO% = 1
    BOTON.Visible = False
    TXTMOVIL.Visible = False
    If MSF.Rows = 1 Then MSF.Rows = 2 ' PARA EVITAR QUE DE ERROR LA RUTINA QUE COPIA DE  UN GRID A OTRO.
    Hide
99  BOTREC.Enabled = True

End Sub

Private Sub Command4_Click()
End Sub

Private Sub CMDINSERTAR_Click()
   Call INSERTFILAVACIA(Me.MSF, MSF.Row)
   Call COLOREAR_CELDA_EDITAB_BLANCO(MSF, Me.EDITABLES.Caption)
   TXTMOVIL.Visible = False
   BOTON.Visible = False
End Sub

Private Sub Form_Activate()

  If MSF.Rows > 1 Then Call COLOREAR_CELDA_EDITAB_BLANCO(Me.MSF, Me.EDITABLES.Caption)
  If ACTIVOX% > 0 Then Exit Sub
  ACTIVOX% = 1
  If MSF.Rows > 1 Then
    Call ACTUGRILLA

    'MSF.SetFocus
    UNICAEDITABLE% = 0
    COLU_EDIT$ = TRIM$(EDITABLES.Caption)
    If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
       Call TEXTOLOTES$(COLU_EDIT$, ";")
       On Error Resume Next
       COLU1& = XVALO(CADENATEX$(1))
       REG1& = XVALO(CADENATEX$(4))
       On Error GoTo 0
       If REG1& > MSF.Rows - 1 Then
         REG1& = MSF.Rows - 1 'PARA QUE NO APAREZCA MAS ALLA DEL ULTIMO REGISTRO DEL MSF
         Call POSICIONAR_TXTMOVIL(TXTMOVIL, MSF, REG1&, COLU1&)

       End If

       If UBound(CADENATEX$) = 1 Then UNICAEDITABLE% = 1
       '
       MSF.COL = COLU1&
       'MSF.Row = 1
       MSF.Row = MSF.Rows - 1
       With TXTMOVIL
           If UltimaFilaVacia%(Me.MSF, 0) < 1 Then ' para posionar el text en la fila siguiente vacia cuando se abre
             .TEXT = ""
             MSF.Rows = MSF.Rows + 1
             MSF.Row = MSF.Row + 1
             .ToolTipText = MSF.Row & "," & COLU1&
             Call COLOREAR_CELDA_EDITAB_BLANCO(Me.MSF, Me.EDITABLES.Caption)
             MSF.COL = COLU1& ' LO REPITO POR QUE LA RUTINA COLOREAR..... LO DEJA EN LA ULTIMA POSICION
           End If
           
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
   If VENTANA.Inicializar = 1 Then
     If VENTANA.HabilitaBorrar = 1 Then Me.BOTERASE.Enabled = True
     If VENTANA.HabilitaInsertar = 1 Then Me.CMDINSERTAR.Enabled = True
     If VENTANA.NoMostrarBotonTilde = 1 Then Me.BOTREC.Visible = False

     Call CARGA_ENCABEZADO(MSF, VENTANA.Campox, Me)
'
     MSF.Rows = 2
     FORMATOS$ = GENERA_FORMATOS$(VENTANA.Campox) ' DEVUELVE ESTO = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
     FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(VENTANA.Campox)
     'CAMPOS$ = "           /A16;       /A20;              /A3;      /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
     Call CARGA_ENCABEZADO(VENTABNEW.MSFTOTALES, FORMATOSCONBARRA$, VENTABNEW) 'LES DA FORMATO AL ENCABEZADO DE TOTOLES
     Call CARGA_ENCABEZADO(VENTABNEW.MSFFORMULA, VENTANA.Campox, VENTABNEW)
     Call AGREGAFILA(VENTABNEW.MSFFORMULA, FORMATOS$, 2)   'DEBEN IR EN LA FILA 2
     Me.EDITABLES.Caption = VENTANA.CampEditables
     Me.AGREGAFILASNUEVAS = VENTANA.AgregaRegistro
     Me.Caption = VENTANA.Titulo
    
     '
     'OJO !!! QUE ESTO TIENE QUE IR SIEMPRE DESPUES DE LOS CARGA_ENCABEZADO POR QUE CASO CONTRARIO ROWS VUELVEN A 1
     VENTABNEW.MSFTOTALES.Rows = 2
     VENTABNEW.MSFFORMULA.Rows = 6
     
     If VENTANA.UtilizaFormula > 0 Then
       For i% = 1 To 32
       If VENTANA.formula(i%) <> "" Then
           VENTABNEW.MSFFORMULA.TextMatrix(1, i%) = VENTANA.formula(i%) '"4;*;6" 'DEBEN IR EN FILA 1
         End If
       Next i%
       'CAMPOS ECO LOS PONE EN LA GRILLA OCULTA EN EL REGISTRO 3
       For i% = 1 To 32
         If UCase$(VENTANA.TOTALES(i%)) = "S" Then
            VENTABNEW.MSFTOTALES.TextMatrix(1, i%) = "S" '"Total" 'DEBEN IR EN FILA 1
         End If
       Next i%
       
       For i% = 1 To 32
         If VENTANA.ColEcoCampo(i%) <> "" Then
           VENTABNEW.MSFFORMULA.TextMatrix(3, i%) = VENTANA.ColEcoCampo(i%) '"CODIGO;Descripcion;Master"
         End If
       Next i%
       '
       
       'AGREGA UN TEXTO EN UNA CELDA ENCABEZADO DE MSFTOTALES POR EJEMPLO EN LA CELDA ANTERIOR DONDE SE IMPRIME UN TOTAL
       'LE ESCRIBO TOTAL Y ME QUEDA TOTAL|$125
       For i% = 1 To 32
         If VENTANA.TxtUbicadoGrillaTotal(i%) <> "" Then
            On Error Resume Next
            VENTABNEW.MSFTOTALES.TextMatrix(1, i%) = VENTANA.TxtUbicadoGrillaTotal(i%) '"Total" 'DEBEN IR EN FILA 1
            On Error GoTo 0
         End If
       Next i%
     End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.APROBADO = 0
   ACTIVOX% = 0
End Sub


Private Sub Form_Resize()
   MSF.Width = Me.Width - 1250
   Me.MSFTOTALES.Width = Me.Width - 1250
   Me.MSFTOTALES.Width = Me.Width - 1200
   Me.BOTERASE.Left = Me.Width - 1100
   Me.BOTEXIT.Left = Me.Width - 1100
   Me.BOTREC.Left = Me.Width - 1100
   Me.CMDINSERTAR.Left = Me.Width - 1100
End Sub

Private Sub MSF_Click()
          COLUM = MSF.MouseCol
          If MSF.Row > MSF.Rows - 1 Then Exit Sub
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
            .TEXT = MSF.TEXT
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            .Visible = True
            .ToolTipText = MSF.Row & "," & COLUM
            Call TXTMOVIL_Change
            .SetFocus
          End With
          Call ACTUGRILLA

End Sub

Private Sub MSF_Scroll()
'   Call Form_KeyUp(40, 0)
   If TXTMOVIL.Visible = True Then
     'TXTMOVIL.Left = MSF1.CellLeft
     REG& = XVALO(TXTMOVIL.ToolTipText)
     MSF.TEXT = TXTMOVIL
     TXTMOVIL.Visible = False
   End If
   If BOTON.Visible = True Then BOTON.Visible = False
'   ScrollActivo% = 1
   MSFTOTALES.LeftCol = MSF.LeftCol
End Sub
Private Sub MSFTOTALES_Scroll()
   MSF.LeftCol = MSFTOTALES.LeftCol
End Sub


Private Sub Timer1_Timer()
   If TRIM$(COLU_EDIT$) = "" Then Exit Sub
   If TXTMOVIL.Visible = False Then
     BOTON.Visible = False
   End If
    
   If TXTMOVIL.Top > MSF.CellTop + MSF.Top Then
      TXTMOVIL.Visible = False
      BOTON.Visible = False
      Timer1.Enabled = False
   End If

End Sub

Private Sub TXTMOVIL_Change()

'     MSF.TextMatrix(REG1&, COLU1& + 1) = ""
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     ARGUMENBUS$ = TRIM$(MSFFORMULA.TextMatrix(3, COLU1&))
     If ARGUMENBUS$ <> "" Then
         Call TEXTOLOTES(ARGUMENBUS$, ";")
         CAMPO1$ = CADENATEX$(1)
         campo2$ = CADENATEX$(2)
         TABLA$ = TRIM$(UCase$(CADENATEX$(3)))
         'PUEDE SER QUE TENGA CONDICION EXTRA
         On Error Resume Next
         CONDI0$ = CADENATEX$(4)
         On Error GoTo 0
         'If TRIM$(CONDI$) <> "" Then CONDI$ = CONDI$ + " AND "
         '
         SKL$ = "": COL% = COLU1&
         CONDI1$ = TRIM$(VENTANA.Condicion1(COL%))
         CONDI2$ = TRIM$(VENTANA.Condicion2(COL%))
         'ESTA FUNCION (RESULCONDI2$)' ARMA LA CONDICION CON EL PARAMETRO QUE VA A SER EL VALOR DE LA CELDA SEGUN COLUMNA
         'EJEMPLO SE LE PASA: FAMILIA = COL@1 Y DEVUELVE: FAMILIA = '20' DONDE 20 ES EL VALOR DE LA CELDA
         If CONDI2$ <> "" Then CONDI2$ = RESULCONDI2$(CONDI2$, VENTABNEW.MSF, REG1&, LETFORM$)


         SKL$ = CONDI0$
         If CONDI1$ <> "" Then
            If SKL$ <> "" Then
               SKL$ = SKL$ & " AND " & CONDI1$
            Else
               SKL$ = SKL$ & CONDI1$
            End If
         End If
         If CONDI2$ <> "" Then
            If SKL$ <> "" Then
               SKL$ = SKL$ & " AND " & CONDI2$
            Else
               SKL$ = SKL$ & CONDI2$
            End If
         End If
         
         'SE VALIDA EL FORMATO Y SI ES DE LA TABLA DATASQL, POR QUE EL DATO SE GUARDA COMO TEXTO SIEMPRE Y CON MASCARA 00000001
         'ENTONCES SI ES NUMERICO SE HACE UN CASTEO PARA QUE MACHEE NUMERO CON VALOR NUMERICO DEL CAMPO.
         If SKL$ <> "" Then
            CONDI$ = SKL$ & " AND "
         End If
         FORMATX$ = FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&)
         If UCase$(Left$(FORMATX$, 1)) = "F" And TABLA$ = "DATASQL" Then
            CONDI$ = CONDI$ + " CAST(SUBSTRING(" & CAMPO1$ & ",1,16) AS Integer ) = '" & TXTMOVIL & "'"
         Else
           CONDI$ = CONDI$ + CAMPO1$ & " = '" & TXTMOVIL & "'" 'SI NO COMPARA CON LO QUE CONTIENE LA CAJA DE TEXTO SIN CASTEO
         End If
         '
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
   ElseIf KeyCode = vbKeyF3 Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG1& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     'COPIA LA INFORMACION DEL CASILLERO INMEDIATO SUPERIOR EN TXTMOVIL
     If REG1& > 1 Then
       If TRIM$(MSF.TextMatrix(REG1& - 1, COLU1&)) <> "" Then
         TXTMOVIL.TEXT = TRIM$(MSF.TextMatrix(REG1& - 1, COLU1&))
       End If
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
         
         'ESTO ES PARA QUE SI NO SE ESCRIBE EL PRIMER CAMPO EDITABLE NO AVANCE
         PrimCOLEditable& = XVALO(CADENATEX$(1))
         If COLU1& = PrimCOLEditable& And TRIM$(TXTMOVIL) = "" Then
            TXTMOVIL.Visible = True
            TXTMOVIL.Move MSF.Left + MSF.CellLeft, MSF.Top + MSF.CellTop, MSF.CellWidth - 8, MSF.CellHeight - 8
           Call PINTATEXT(TXTMOVIL)
           Exit Sub
         End If
         '
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
                  'recordar que si un campo esta oculto el mismo no debe estar editable
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
                  MSF.Row = REG& ' lo VUELVO A POSICIONAR POR QUE APARENTEMENTE AL COLOREAR QUEDA EL VALOR DE LA ULTIMA FILA
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
     If REG& > MSF.Rows - 1 Then
       REG& = MSF.Rows - 1 'PARA QUE NO APAREZCA MAS ALLA DEL ULTIMO REGISTRO DEL MSF
       Call POSICIONAR_TXTMOVIL(TXTMOVIL, MSF, REG&, COLU1&)
     End If
     
     
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
            If MSF.Row > 1 Then MSF.Row = MSF.Row - 1
          MSF.COL = COLU1&
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
             MSF.TextMatrix(REG&, COLU1&) = VALORSEGFORMATO(TXTMOVIL, Me.MSFFORMULA, 2, COLU1&)

            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            .SetFocus
          End With
             Call AGREGABOTON(COLU1&, REG1&)

             Call ACTUGRILLA
            On Error GoTo 0
        Case 40 '(flecha 40 abajo)
            On Error Resume Next
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
            MSF.TextMatrix(REG&, COLU1&) = VALORSEGFORMATO(TXTMOVIL, Me.MSFFORMULA, 2, COLU1&)
            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            .SetFocus
          End With
          Call AGREGABOTON(COLU1&, REG1&)
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
    BOTON.Height = TXTMOVIL.Height - 50
    BOTON.Left = TXTMOVIL.Left + (TXTMOVIL.Width - BOTON.Width)
    BOTON.Top = TXTMOVIL.Top
  End If

End Sub
                 
