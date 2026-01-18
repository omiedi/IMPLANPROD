VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMZELEOEX 
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Bindings        =   "FRMZELEOEX.frx":0000
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6800
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   12648447
      FocusRect       =   2
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
      Width           =   11730
   End
   Begin MSFlexGridLib.MSFlexGrid msf2 
      Bindings        =   "FRMZELEOEX.frx":0014
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   16777215
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   ""
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
Attribute VB_Name = "FRMZELEOEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CODITABLA$
Public RESP_ZELE$


Private Sub Form_Activate()
    If Me.Width > 12000 Then Me.Width = 12000
    Timer1.Enabled = True ' PARA QUE PROVOQUE UNA DEMORA EN PINTAR EL TEXTLISTA
    'POR QUE SE VE QUE POR EL APLICAR SKIN NO FUNCIONA POR QUE PIERDE EL FOCO
    RESP_ZELE$ = ""
    For I& = 0 To UBound(RESP_ZELE_VECT)
       RESP_ZELE_VECT(I&) = ""
    Next I&
End Sub

Private Sub Form_Resize()
   FRMZELEOEX.TXTBUSCAR.Width = FRMZELEOEX.Width - 100
   FRMZELEOEX.MSF1.Width = FRMZELEOEX.Width - 100
   FRMZELEOEX.msf2.Width = FRMZELEOEX.Width - 100
   FRMZELEOEX.MSF1.ColWidth(0) = FRMZELEOEX.Width - 100
   FRMZELEOEX.msf2.ColWidth(0) = FRMZELEOEX.Width - 100
End Sub

Private Sub MSF1_DblClick()
    MSF1.COL = 0
    On Error Resume Next
    RESP_ZELE$ = MSF1.TEXT

    RESP_ZELE$ = MSF1.TEXT
    On Error GoTo 0
    REG& = MSF1.MouseRow
    'If REG& < 1 Then Exit Sub
    For I& = 1 To MSF1.Cols - 1
      RESP_ZELE_VECT(I&) = MSF1.TextMatrix(REG&, I&)
    Next I&
   Hide
    
End Sub

Private Sub MSF2_DblClick()
    MSF1.COL = 0
    RESP_ZELE$ = msf2.TEXT
    REG& = msf2.MouseRow
    'If REG& < 1 Then Exit Sub
    For I& = 0 To msf2.Cols - 1
      RESP_ZELE_VECT(I&) = msf2.TextMatrix(REG&, I&)
    Next I&
    Hide
End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       RESP_ZELE$ = msf2.TEXT
       REG& = MSF1.MouseRow
       'If REG& < 1 Then Exit Sub
       For I& = 1 To msf2.Cols - 1
          RESP_ZELE_VECT(I&) = MSF1.TextMatrix(REG&, I&)
       Next I&
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
       On Error Resume Next
       RESP_ZELE$ = MSF1.TEXT
       REG& = MSF1.MouseRow
       On Error GoTo 0
       'If REG& < 1 Then Exit Sub
       For I& = 1 To MSF1.Cols - 1
          RESP_ZELE_VECT(I&) = MSF1.TextMatrix(REG&, I&)
       Next I&
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


Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF1.Visible = False
      msf2.Visible = True
      Exit Sub
    End If
    '
    MSF1.Rows = 0
    MSF1.Visible = True
    MSF1.ZOrder 0
    MSF1.Cols = msf2.Cols
    MSF1.Redraw = False
    msf2.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN MSF1  LOS QUE MACHEAN
    For I& = 0 To msf2.Rows - 1
      For j& = 0 To msf2.Cols - 1
        If InStr(UCase$(msf2.TextMatrix(I&, 0)), UCase$(TXTBUSCAR)) Then
          REG& = REG& + 1
          MSF1.Rows = MSF1.Rows + 1
          For H& = 0 To msf2.Cols - 1
            MSF1.TextMatrix(REG& - 1, 0) = msf2.TextMatrix(I&, 0)
          Next H&
          Exit For
        End If
      Next j&
    Next I&
    '
    MSF1.Redraw = True
    '
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
