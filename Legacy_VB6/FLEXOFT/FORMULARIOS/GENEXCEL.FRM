VERSION 5.00
Begin VB.Form GENEXCEL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Planilla Excel"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   2430
      Left            =   6930
      ScaleHeight     =   2370
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   -105
      Width           =   1275
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   125
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   8
         Top             =   1500
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   12000
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
         Height          =   600
         Left            =   125
         Picture         =   "GENEXCEL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "GENEXCEL.frx":014A
         Top             =   1785
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   5565
      Top             =   105
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Columna:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5145
      TabIndex        =   5
      Top             =   1890
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   6300
      TabIndex        =   4
      Top             =   1845
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Registro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2835
      TabIndex        =   3
      Top             =   1890
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Height          =   260
      Left            =   3990
      TabIndex        =   2
      Top             =   1850
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   420
      Picture         =   "GENEXCEL.frx":206C
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   1
      Top             =   840
      Width           =   6525
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Generando Planilla ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4005
   End
End
Attribute VB_Name = "GENEXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
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
   Unload Me   ' Close: End
   '
End Sub

Private Sub Form_Load()
   '
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   Label2.Caption = "Espere un Momento ..."
   Show
   Refresh
   OPENFORMS = DoEvents
   '
   LUCOM$ = "C:\FLEXOFT\"
   CMDX$ = RTrim$(LTrim$(Command$))
   If CMDX$ <> "" Then
      LUCOM$ = CMDX$
   End If
   '
   Screen.MousePointer = 11
   RENG& = 1: COLU& = 1
   N1% = FreeFile
   '
   LOTRA = 0
   On Error Resume Next
   LOTRA = FileLen(LUCOM$ + "TREXCEL.TRF")
   On Error GoTo 0
   If LOTRA < 1 Then
      MsgBox "Falta Archivo de Transferencia"
      Unload Me   ' CLOSE:END
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
   OPENFORMS = DoEvents
   '
   REPLACOMA% = 0
   If Val("9.99") <> 9.99 Then REPLACOMA% = 1
   '
   While Not EOF(N1%)
      Line Input #N1%, A$
      TELEI& = TELEI& + Len(A$)
      Call TRAZAR(TELEI&, FIN&)
      If A$ = "@NEWLINE@" Then
           RENG& = RENG& + 1
           Label3 = Str$(RENG&)
           Label3.Refresh
           COLU& = 1
           Label4 = Str$(COLU&)
           Label4.Refresh
         Else
           If Left$(A$, 3) = "#F#" Then
                FECHA = CONFECHA(Mid$(A$, 4))
                If FECHA <> CDate("01/Ene/1980") Then
                   XLSHEET.Cells(RENG&, COLU&) = FECHA
                End If
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
                VALOR = (Int(EDIVI& * Val(TENU$) + 0.5)) / EDIVI&
                If Abs(VALOR) >= 0.5 / EDIVI& Then
                   XLSHEET.Cells(RENG&, COLU&) = VALOR
                End If
              Else
                TEXTO$ = A$
                If Left$(A$, 3) = "#T#" Then TEXTO$ = Mid$(A$, 4)
                XLSHEET.Cells(RENG&, COLU&) = Chr$(39) + TEXTO$
           End If
           COLU& = COLU& + 1
           Label4 = Str$(COLU&)
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
   Unload Me ' Close: End
   '
   '
10 Resume 11
11 On Error GoTo 0
   MsgBox "Error de Acceso a Archivo de Transferencia"
   Unload Me ' Close: End
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   Unload Me  ' Close: End
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
   Unload Me   ' Close: End
   '
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

Sub TRAZAR(I&, FIN&)
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
    BARRATRAZA.Left = COEF * I&
    BARRATRAZA.Refresh
    MARCOBARRA.Refresh
    OPENFORMS = DoEvents
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
   CONFECHA = CDate("01/Ene/1980")
   '
End Function
