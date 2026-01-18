VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FRMMODEDI 
   Caption         =   "MODIFICAR"
   ClientHeight    =   1110
   ClientLeft      =   1875
   ClientTop       =   3315
   ClientWidth     =   14550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   14550
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox TXTMOVIL 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Costo segun el proveedor y la via de entrega"
      Top             =   0
      Visible         =   0   'False
      Width           =   1336
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4200
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2520
      OleObjectBlob   =   "FRMMODEDI.frx":0000
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   10200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   730
      Width           =   2415
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   12840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   730
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   675
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   1191
      _Version        =   393216
      BackColorFixed  =   12632256
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.Label EDITABLES 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "FRMMODEDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FORMULARIOX As Form
Public MSF As Object
Public APROBADO%
Dim UNICAEDITABLE%

Private Sub cmdcancelar_Click()
   Unload Me
End Sub

Private Sub Frame1_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   APROBADO% = 1
   Hide
   Command1.Enabled = True
End Sub

Private Sub Form_Activate()
  If MSF1.Rows > 1 Then

    'MSF1.SetFocus
    APROBADO% = 0
    UNICAEDITABLE% = 0
    COLU_EDIT$ = TRIM$(EDITABLES.Caption)
    If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
       Call TEXTOLOTES$(COLU_EDIT$, ";")
       On Error Resume Next
       COLU1& = XVALO(CADENATEX$(1))
       If UBound(CADENATEX$) = 1 Then UNICAEDITABLE% = 1
       '
       MSF1.COL = COLU1&
       MSF1.Row = 1

       With TXTMOVIL
          .ZOrder 0
          .Width = MSF1.CellWidth
          .Left = MSF1.CellLeft + MSF1.Left
          .Top = MSF1.CellTop + MSF1.Top
          .TEXT = MSF1.TEXT
          .Visible = True
          .ToolTipText = MSF1.Row & "," & COLU1&
          .SetFocus
       End With
    End If
  End If
End Sub

Private Sub MSF1_Click()
          COLUM = MSF1.MouseCol
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

50        MSF1.COL = COLUM
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF1.CellWidth
            .Left = MSF1.CellLeft + MSF1.Left
            .Top = MSF1.CellTop + MSF1.Top
            .TEXT = MSF1.TEXT
            .Visible = True
            .ToolTipText = MSF1.Row & "," & COLUM
            .SetFocus
          End With
End Sub
Private Sub TXTMOVIL_GotFocus()
     Call PINTATEXT(TXTMOVIL)
End Sub
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     If TXTMOVIL.Visible = True Then
        MSF1.TEXT = TRIM$(TXTMOVIL.TEXT)
        TXTMOVIL.Visible = False
        If UNICAEDITABLE% = 1 Then Exit Sub
     End If
'

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
                 
                 MSF1.COL = COLU1&
                 MSF1.Row = REG&
                 With TXTMOVIL
                  .ToolTipText = MSF1.Row & "," & COLU1&
                  .ZOrder 0
                  .Width = MSF1.CellWidth
                  .Left = MSF1.CellLeft + MSF1.Left
                  .Top = MSF1.CellTop + MSF1.Top
'                  MSF1.TEXT = .TEXT
                  .TEXT = TRIM$(MSF1.TEXT)
                  .Visible = True
                  Call PINTATEXT(TXTMOVIL)
                  .SetFocus
                 End With
               End If
           Next i&
         Else
            'ACA ES CUANDO LLEGO AL FINAL DE LA FILA
            For i& = UBound(CADENATEX$) To 1 Step -1
               COLU1& = XVALO(CADENATEX$(i&))
               If COLU1& < ColumActual& Or UNICAEDITABLE% = 1 Then
                 On Error GoTo 0
                 'VALIDO SI LLEGO AL FINAL DE LA FILA BAJA 1 Y EMPIEZA DE IZQUIERDA
                 'SI LLEGO AL FINAL DE LOS REGISTROS SUBE AL PRIMER EDITABLE DE IZQUIERDA
                 If MSF1.Rows - 1 > MSF1.Row Then
                   REG& = REG& + 1
                 Else
                   COLU1& = CADENATEX$(1)
                   REG& = 1
                 End If

                 MSF1.COL = COLU1&
                 MSF1.Row = REG&
                 With TXTMOVIL
                  .ToolTipText = MSF1.Row & "," & COLU1&
                  .ZOrder 0
                  .Width = MSF1.CellWidth
                  .Left = MSF1.CellLeft + MSF1.Left
                  .Top = MSF1.CellTop + MSF1.Top
                   .TEXT = TRIM$(MSF1.TEXT)
                  .Visible = True
                  Call PINTATEXT(TXTMOVIL)
                  .SetFocus
                 End With
               End If
             Next i&
           End If
          On Error GoTo 0
       End If
   End If
End Sub

'Private Sub Command1_Click()
'    Command1.Enabled = False
'    '
'    If Option1.Value = False And Option2.Value = False Then
'       Call COMUNI("Debe Seleccionar una Opción a Realizar")
'       GoTo 99
'    End If
'    '
'    If XVALO(Text1) < 1 Then GoTo 98
'    '
'    If XVALO(Text1) = XVALO(MSF1.TextMatrix(1, 0)) Then GoTo 98
'    '
'    If Option1.Value = True Then
'       Call INSERTARYELMINAR(XVALO(Text1))
'    Else
'       Call INSERTYDESPLAZAR(XVALO(Text1))
'    End If
'    '
'98  Unload Me
'99  Command1.Enabled = True
'End Sub

'Private Sub Command1_GotFocus()
'       Command1.BackColor = vbYellow
'
'End Sub

'Private Sub Command1_LostFocus()
'    Command1.BackColor = &H8000000F
'End Sub

Private Sub MSF_Click()
          COLUM = MSF.MouseCol
50        MSF.COL = COLUM
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            .TEXT = MSF.TEXT
            .Visible = True
            .ToolTipText = MSF.Row & "," & COLUM
            .SetFocus
          End With
       

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      Call cmdcancelar_Click
   End If
End Sub

Private Sub Form_Load()
    '
    '
'    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
 End Sub
   
'Private Sub Text1_GotFocus()
'    Text1.BackColor = vbYellow
'    Call PINTATEXT(Text1)
'End Sub

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'      If Option1.Value = True Or Option2.Value = True Then
'         Command1.SetFocus
'      End If
'    End If
'End Sub

'Private Sub Text1_LostFocus()
'    Text1.BackColor = vbWhite
'End Sub

'Private Sub Timer1_Timer()
'    Call PINTATEXT(Text1)
'    Timer1.Enabled = False
'End Sub

'Private Sub Text1_Change()
'    If XVALO(Text1) > MSF.Rows - 1 Then
'       MsgBox "Número de Fila Inválido"
'    End If
'End Sub
'Sub INSERTARYELMINAR(FILA&)
'   If FILA& < 1 Then Exit Sub
'   Call SOBRE_ESCRIBIR_FILA(FILA&)
'   Call borrafila(XVALO(MSF1.TextMatrix(1, 0)), FORMULARIOX, MSF)
'End Sub
'   Sub SOBRE_ESCRIBIR_FILA(FILA&)
'    For i& = 1 To MSF1.Cols - 1
'       DATO$ = MSF1.TextMatrix(1, i&)
'       MSF.TextMatrix(FILA&, i&) = DATO$
'    Next i&
'   End Sub
'Sub INSERTYDESPLAZAR(FILA&)
'    Screen.MousePointer = 11
'    '
'    UR& = MSF.Rows - 1
'    MSF.Rows = MSF.Rows + 1
'    For U& = UR& To FILA& Step -1
'      For i& = 1 To MSF.Cols - 1
'         DATO$ = MSF.TextMatrix(U&, i&)
'         MSF.TextMatrix(U& + 1, i&) = DATO$
'      Next i&
'    Next U&
'    '
'    'SE REALIZA ESTA CONSIDERACION POR QUE CAMBIA CUANDO SE SELECCIONA UNA POSICION MAYOR Y SE INTENTA INCLUIR EN UNA INFERIOR POR EL HECHO
'    'QUE ESTO FUNCIONA PRIMERO CREANDO UN  REGISTRO PARA DEJAR LIBRE LA FILA DONDE SE VA A INSERTAR LUEGO SE SOBRE ESCRIBE ESTA
'    'Y POR ULTIMO SE BORRA EL REGISTRO DE LA POSICION ORIGINAL
'    VALOR% = 0
'    VALOR2% = 1
'    If FILA& > XVALO(MSF1.TextMatrix(1, 0)) Then
'       VALOR% = 1
'       VALOR2% = 0
'    End If
'    '
'    Call SOBRE_ESCRIBIR_FILA(FILA& + VALOR%) ' SOBRE ESCRIBE LA FILA
'    Call borrafila(XVALO(MSF1.TextMatrix(1, 0)) + VALOR2%, FORMULARIOX, MSF)  ' BORRA LA FILA ORIGINAL
'    Screen.MousePointer = 1
'    '
'End Sub

Private Sub TXTMOVIL_Change()

End Sub

Private Sub TXTMOVIL_LostFocus()
     If TXTMOVIL.Visible = True Then
        Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
        On Error Resume Next
        REG& = XVALO(CADENATEX$(1))
        COLU1& = XVALO(CADENATEX$(2))
        On Error GoTo 0
        MSF1.TextMatrix(REG&, COLU1&) = TXTMOVIL
     End If

End Sub


