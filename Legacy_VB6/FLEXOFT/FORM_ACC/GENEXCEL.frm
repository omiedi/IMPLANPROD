VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form GENEXCEL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Planilla Excel"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   120
      ScaleHeight     =   165
      ScaleWidth      =   3375
      TabIndex        =   11
      Top             =   1920
      Width           =   3405
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   12
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4440
      Top             =   120
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H004F4F75&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   7035
      ScaleHeight     =   8490
      ScaleWidth      =   1140
      TabIndex        =   5
      Top             =   -315
      Width           =   1170
      Begin VB.PictureBox Picture16 
         Height          =   2600
         Left            =   480
         ScaleHeight     =   2535
         ScaleWidth      =   1020
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   2595
            Left            =   480
            Picture         =   "GENEXCEL.frx":0000
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "GENEXCEL.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5565
      Top             =   105
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "GENEXCEL.frx":24E2
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "GENEXCEL.frx":255C
      Top             =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   6525
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Columna"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5040
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1725
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3435
      TabIndex        =   2
      Top             =   1890
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4590
      TabIndex        =   1
      Top             =   1725
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   240
      Picture         =   "GENEXCEL.frx":2790
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Generando Planilla ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4005
   End
End
Attribute VB_Name = "GENEXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   On Error Resume Next
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   MsgBox "Transferencia Cancelada !!!"
   Unload Me    ' Close: End
   '
End Sub

Private Sub Form_Load()
   '
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Label1.BackColor = &HFCD7B6
      MARCOBARRA.Top = 2200
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If

   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   Call APLICACIONSKINS(Me)
   Timer2.Enabled = True
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Timer1_Timer()
   '
   If Image1.Visible = True Then
        Image1.Visible = False
      Else
        Image1.Visible = True
   End If
   '
End Sub

Sub TRAZAR(i&, FIN&)
    '
    Static FINA&, COEF, TANT, TANTUI, VISI%, VERTI%, TINI, IFORMU&, IBARRA&, IMARCO&
    '
    If FIN& <> FINA& Then
        FINA& = FIN&
        VERTI% = 0
        If FIN& > 0 Then
            COEF = MARCOBARRA.Width / FIN&
            TINI = Timer
        End If
        BARRATRAZA.Left = 0
        BARRATRAZA.Top = 0
        BARRATRAZA.Width = MARCOBARRA.Width
        BARRATRAZA.Height = MARCOBARRA.Height
    End If
    '
    If Abs(Timer - TANTUI) < 0.25 Then
       Exit Sub
    End If
    TANTUI = Timer
    '
    BARRATRAZA.Left = COEF * i&
    BARRATRAZA.Refresh
    MARCOBARRA.Refresh
    Openforms = DoEvents
    '
End Sub
'
Function CONFECHA(TTXX$)
   '
1  On Error GoTo 10
   CONFECHA = CDate(TTXX$)
   On Error GoTo 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   If Val(Left$(TTXX$, 2)) > 28 Then
     TTXX$ = TRIM$(Str$(Val(Left$(TTXX$, 2)) - 1)) + Mid$(TTXX$, 3)
     GoTo 1
   End If
   '
   CONFECHA = CVDAT(0)
   'CONFECHA = CDate("01/Ene/1980")
   '
End Function

Private Sub Timer2_Timer()
   Static YAENTRO%
   '
   If YAENTRO% > 0 Then Exit Sub
   '
   YAENTRO% = 1
   Label2.Caption = "Espere un Momento ..."
   Refresh
   Openforms = DoEvents
   '
   'LUCOM$ = "C:\FLEXOFT\"
   'CMDX$ = RTrim$(LTrim$(Command$))
   'If CMDX$ <> "" Then
   '   LUCOM$ = CMDX$
   'End If
   '
   Screen.MousePointer = 11
   RENG1& = 1: COLU1& = 1
   N1% = FreeFile
   '
   LOTRA = 0
   On Error Resume Next
   LOTRA = FileLen(LUCOM$ + "TREXCEL.TRF")
   On Error GoTo 0
   If LOTRA < 1 Then
      MsgBox "Falta Archivo de Transferencia"
      GoTo 90
   End If
   '
   On Error GoTo 10
   Open LUCOM$ + "TREXCEL.TRF" For Input Shared As #N1%
   On Error GoTo 0
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   FIN& = LOF(N1%): TELEI& = 0
   Line Input #N1%, NOMBREDEARCHIVO$
   TELEI& = TELEI& + Len(NOMBREDEARCHIVO$)
   Label2.Caption = NOMBREDEARCHIVO
   Label2.Refresh
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   Label6.Visible = True
   Openforms = DoEvents
   '
   REPLACOMA% = 0
   If Val("9.99") <> 9.99 Then REPLACOMA% = 1
   '
   While Not EOF(N1%)
      Line Input #N1%, A$
      TELEI& = TELEI& + Len(A$)
      Call TRAZAR(TELEI&, FIN&)
      If A$ = "@NEWLINE@" Then
           RENG1& = RENG1& + 1
           Label3 = Str$(RENG1&)
           Label3.Refresh
           COLU1& = 1
           Label4 = Str$(COLU1&)
           Label4.Refresh
         Else
           If Left$(A$, 3) = "#F#" Then
                FECHA = CONFECHA(Mid$(A$, 4))
                On Error Resume Next
                'If FECHA <> CDate("01/Ene/1980") Then
                If FECHA <> CVDAT(0) Then
                   XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "dd/mm/yy"
                   'XLSHEET.Cells(RENG1&, COLU1&) = CDate(FECHA)
                   XLSHEET.Cells(RENG1&, COLU1&) = FETEXCOR1$(CVINT(FECHA))
                   XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "dd/mm/yy"
                End If
                XLSHEET.Cells(RENG1&, COLU1&).HorizontalAlignment = xlLeft
                On Error GoTo 0
              ElseIf Left$(A$, 3) = "#N#" Then
                TENU$ = TRIM$(Mid$(A$, 4))
                PPXX% = InStr(TENU$, ".")
                If PPXX% > 0 Then
                  If REPLACOMA% = 1 Then
                    Mid$(TENU$, PPXX%, 1) = ","
                  End If
                End If
                If PPXX% < 1 Then PPXX% = 1 + Len(TENU$)
                LUDECI% = Len(TENU$) - PPXX%: If LUDECI% < 0 Then LUDECI% = 0
                EDIVI& = 10 ^ LUDECI%
                valor = (Int(EDIVI& * Val(TENU$) + 0.5)) / EDIVI&
                XLSHEET.Cells(RENG1&, COLU1&) = valor
                If Abs(valor) >= 0.5 / EDIVI& Then
                   If LUDECI% > 0 Then
                        DECI$ = ""
                        While Len(DECI$) < LUDECI%: DECI$ = DECI$ & "0": Wend
                        XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0." & DECI$
                   Else
                        XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#0"
                   End If
                End If
                XLSHEET.Cells(RENG1&, COLU1&).HorizontalAlignment = xlRight
              Else
                TEXTOS$ = A$
                If Left$(A$, 3) = "#T#" Then TEXTOS$ = Mid$(A$, 4)
                'XLSHEET.Cells(RENG1&, COLU1&) = Chr$(39) + TRIM$(TEXTOS$)
                XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "@"
                XLSHEET.Cells(RENG1&, COLU1&) = TRIM$(TEXTOS$)
                XLSHEET.Cells(RENG1&, COLU1&).HorizontalAlignment = xlLeft
           End If
           COLU1& = COLU1& + 1
           Label4 = Str$(COLU1&)
           Label4.Refresh
      End If
   Wend
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   GoTo 90    ' Close: End
   '
   '
10 Resume 11
11 On Error GoTo 0
   MsgBox "Error de Acceso a Archivo de Transferencia"
   Call FINAL("")    ' Close: End
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   Call FINAL("")    ' Close: End
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   GoTo 90    ' Close: End
   '
90 YAENTRO% = 0
   Unload Me    ' Close: End
   '
End Sub
