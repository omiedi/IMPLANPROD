VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VENTAB2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recepción AHP - Valores expresados en Dolares"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   16890
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   16890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton BOTERASE 
      Caption         =   "Del."
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
      Width           =   540
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
      Height          =   525
      Left            =   16320
      Picture         =   "VENTAB2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Aprobar"
      Top             =   4500
      Width           =   540
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
      Left            =   16320
      Picture         =   "VENTAB2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cierra Formulario "
      Top             =   120
      Width           =   540
   End
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
      Picture         =   "VENTAB2.frx":0614
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
      OleObjectBlob   =   "VENTAB2.frx":091E
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "VENTAB2.frx":0B52
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4605
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
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
      Height          =   885
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar / Doble Click Para Seleccionar"
      Top             =   5880
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   1561
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
      Height          =   375
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Encabezado"
      Top             =   4680
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   661
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
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   5640
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
Attribute VB_Name = "VENTAB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APROBADO%
Dim UNICAEDITABLE%
Sub ACTUGRILLA()
   If MSFFORMULA.Rows > 1 Then
      For I& = 1 To MSFFORMULA.Cols - 1
          FORMULAX$ = TRIM$(MSFFORMULA.TextMatrix(1, I&))
          If FORMULAX$ <> "" And Left(FORMULAX$, 1) <> "=" Then
             Call TEXTOLOTES$(FORMULAX$, ";")
             On Error Resume Next
             SIMB$ = CADENATEX$(2)
             COLU1& = CADENATEX$(1)
             COLU2& = CADENATEX$(3)
             For j& = 1 To MSF.Rows - 1
              Select Case SIMB$
                  Case "+"
                    MSF.TextMatrix(j&, I&) = XVALO(MSF.TextMatrix(j&, COLU1&)) + XVALO(MSF.TextMatrix(j&, COLU2&))
                  Case "-"
                    MSF.TextMatrix(j&, I&) = XVALO(MSF.TextMatrix(j&, COLU1&)) - XVALO(MSF.TextMatrix(j&, COLU2&))
                  Case "*"
                    MSF.TextMatrix(j&, I&) = XVALO(MSF.TextMatrix(j&, COLU1&)) * XVALO(MSF.TextMatrix(j&, COLU2&))
                  Case "/"
                    If XVALO(MSF.TextMatrix(j&, COLU2&)) <> 0 Then
                       MSF.TextMatrix(j&, I&) = XVALO(MSF.TextMatrix(j&, COLU1&)) / XVALO(MSF.TextMatrix(j&, COLU2&))
                    Else
                       MSF.TextMatrix(j&, I&) = 0
                    End If
                 End Select
               Next j&
          ElseIf Left(FORMULAX$, 1) = "=" Then
             Call TEXTOLOTES$(FORMULAX$, ";")
             On Error Resume Next
             SIMB$ = CADENATEX$(2)
             COLU1& = CADENATEX$(1)
             COLU2& = CADENATEX$(3)
             Call EJECUTAR_FUNCIONES(FORMULAX$, I&)
          End If
      Next I&
   End If
   If MSFTOTALES.Rows > 1 Then
     For I& = 1 To MSFTOTALES.Cols - 1
      TOTAX$ = TRIM$(MSFTOTALES.TextMatrix(1, I&))
          If TOTAX$ <> "" Then
            MSFTOTALES.TextMatrix(0, I&) = ""
            For j& = 1 To MSF.Rows - 1
               MSFTOTALES.TextMatrix(0, I&) = XVALO(MSFTOTALES.TextMatrix(0, I&)) + XVALO(MSF.TextMatrix(j&, I&))
            Next j&
          End If
      Next I&
      MSFTOTALES.TextMatrix(0, I&) = FORMATNUM$(XVALO(MSFTOTALES.TextMatrix(0, I&)), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
   End If
End Sub
Sub EJECUTAR_FUNCIONES(FORMULAX$, COLUMNARESULTADO&)
      Dim VALOX()
      Call TEXTOLOTES$(FORMULAX$, ";")
      
      FUNXION$ = Mid$(CADENATEX$(1), 2)
      
      Select Case FUNXION$
         Case "PORCENT_ENTRE_2_VALORES"
            COLU1& = XVALO(CADENATEX$(2))
            COLU2& = XVALO(CADENATEX$(3))
            For j& = 1 To MSF.Rows - 1
                MSF.TextMatrix(j&, COLUMNARESULTADO&) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(XVALO(MSF.TextMatrix(j&, COLU1&)), XVALO(MSF.TextMatrix(j&, COLU2&))), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&))
            Next j&
      End Select

End Sub

Private Sub BOTERASE_Click()
   Call cmdEliminaItem(Me.MSF, MSF.Row)
   TXTMOVIL.Visible = False
End Sub

Private Sub BOTEXIT_Click()
   Me.APROBADO = 0
   Unload Me
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
          .SetFocus
       End With
    End If
    
  End If

End Sub

Private Sub Form_Load()
   'PINTO DE BLANCO LAS CELDAS EDITABLES
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.APROBADO = 0
End Sub


Private Sub MSF_Click()
          COLUM = MSF.MouseCol
          COLU_EDIT$ = TRIM$(EDITABLES.Caption)
          If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
            Call TEXTOLOTES$(COLU_EDIT$, ";")
            On Error Resume Next
            For I& = 1 To UBound(CADENATEX$)
               COLU1& = XVALO(CADENATEX$(I&))
               If COLUM = COLU1& Then
                 
                 On Error GoTo 0
                 GoTo 50
               End If
            Next I&
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
End Sub
Private Sub TXTMOVIL_KeyPress(KeyAscii As Integer)
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
         MSF.TextMatrix(REG&, COLU1&) = TRIM$(FORMATNUM$(XVALO(TXTMOVIL), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&)))
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
           For I& = 1 To UBound(CADENATEX$)
               COLU1& = XVALO(CADENATEX$(I&))
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
                   .TEXT = TRIM$(MSF.TEXT)
                  .Visible = True
                  Call PINTATEXT(TXTMOVIL)
                  .SetFocus
                 End With
               End If
           Next I&
         Else
            'ACA ES CUANDO LLEGO AL FINAL DE LA FILA
            For I& = UBound(CADENATEX$) To 1 Step -1
               COLU1& = XVALO(CADENATEX$(I&))
               If COLU1& < ColumActual& Or UNICAEDITABLE% = 1 Then
                 On Error GoTo 0
                 'VALIDO SI LLEGO AL FINAL DE LA FILA BAJA 1 Y EMPIEZA DE IZQUIERDA
                 'SI LLEGO AL FINAL DE LOS REGISTROS SUBE AL PRIMER EDITABLE DE IZQUIERDA
                 If MSF.Rows - 1 > MSF.Row Then
                   REG& = REG& + 1
                 Else
                   COLU1& = CADENATEX$(1)
                   REG& = 1
                 End If

                 MSF.COL = COLU1&
                 MSF.Row = REG&
                 With TXTMOVIL
                  .ToolTipText = MSF.Row & "," & COLU1&
                  .ZOrder 0
                  .Width = MSF.CellWidth
                  .Left = MSF.CellLeft + MSF.Left
                  .Top = MSF.CellTop + MSF.Top
                   .TEXT = TRIM$(MSF.TEXT)
                  .Visible = True
                  Call PINTATEXT(TXTMOVIL)
                  .SetFocus
                 End With
               End If
             Next I&
           End If
          On Error GoTo 0
       End If
   End If
End Sub
Private Sub TXTMOVIL_LostFocus()

     If TXTMOVIL.Visible = True Then
     Call TEXTOLOTES$(TXTMOVIL.ToolTipText, ",")
     On Error Resume Next
     REG& = XVALO(CADENATEX$(1))
     COLU1& = XVALO(CADENATEX$(2))
     On Error GoTo 0
     MSF.TextMatrix(REG&, COLU1&) = TRIM$(FORMATNUM$(XVALO(TXTMOVIL), FormatoSegunCampo$(Me.MSFFORMULA, 2, COLU1&)))
'     TXTMOVIL.Visible = False
     Call ACTUGRILLA
     End If
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
        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            MSF.TEXT = TXTMOVIL.TEXT

            'MSF.SetFocus
            If MSF.Row > 1 Then MSF.Row = MSF.Row - 1
          MSF.COL = COLU1&
          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
            
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
            .SetFocus
          End With
'            Call VERIFICA_CANTI(MSF.Row - 1)
             Call ACTUGRILLA
            On Error GoTo 0
        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
            On Error Resume Next
            'MSF.SetFocus
            MSF.TEXT = TXTMOVIL.TEXT

            If MSF.Row < MSF.Rows Then MSF.Row = MSF.Row + 1
          MSF.COL = COLU1&

          With TXTMOVIL
            .ZOrder 0
            .Width = MSF.CellWidth
            .Left = MSF.CellLeft + MSF.Left
            .Top = MSF.CellTop + MSF.Top
             .TEXT = TRIM$(MSF.TEXT)
            .Visible = True
            Call PINTATEXT(TXTMOVIL)
            .ToolTipText = MSF.Row & "," & COLU1&
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



Private Sub TXTBUSCAR_Change()

End Sub
