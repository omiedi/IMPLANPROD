VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form REPOEQUI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Equipos de Clientes"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   17040
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   17040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTSERIEMOTOR 
      Appearance      =   0  'Flat
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
      Left            =   11020
      TabIndex        =   8
      Top             =   1080
      Width           =   3380
   End
   Begin VB.TextBox TXTMOTOR 
      Appearance      =   0  'Flat
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
      Left            =   7380
      TabIndex        =   9
      Top             =   1080
      Width           =   3420
   End
   Begin VB.TextBox TXTINTERNO 
      Appearance      =   0  'Flat
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
      Left            =   3735
      TabIndex        =   7
      Top             =   1080
      Width           =   3380
   End
   Begin VB.TextBox TXTSERIE 
      Appearance      =   0  'Flat
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
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3360
   End
   Begin VB.TextBox TXTMODELO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   10920
      TabIndex        =   17
      Top             =   450
      Width           =   3165
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   10950
      Sorted          =   -1  'True
      TabIndex        =   28
      Text            =   "CMBCAMPOS"
      Top             =   450
      Width           =   3400
   End
   Begin VB.TextBox TXTTIPOEQUIPO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7550
      TabIndex        =   34
      Top             =   450
      Width           =   3130
   End
   Begin VB.TextBox TXTMARCA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4150
      TabIndex        =   33
      Top             =   450
      Width           =   3130
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   11040
      Sorted          =   -1  'True
      TabIndex        =   32
      Text            =   "CMBCAMPOS"
      Top             =   1080
      Width           =   3645
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   7440
      Sorted          =   -1  'True
      TabIndex        =   31
      Text            =   "CMBCAMPOS"
      Top             =   1080
      Width           =   3645
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   3750
      Sorted          =   -1  'True
      TabIndex        =   30
      Text            =   "CMBCAMPOS"
      Top             =   1080
      Width           =   3645
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   29
      Text            =   "CMBCAMPOS"
      Top             =   1080
      Width           =   3645
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7560
      Sorted          =   -1  'True
      TabIndex        =   27
      Text            =   "CMBCAMPOS"
      Top             =   450
      Width           =   3400
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4150
      Sorted          =   -1  'True
      TabIndex        =   26
      Text            =   "CMBCAMPOS"
      Top             =   450
      Width           =   3400
   End
   Begin VB.CommandButton Command43 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   14760
      Picture         =   "REPOEQUI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Limpiar Casilleros"
      Top             =   120
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2040
      Picture         =   "REPOEQUI.frx":014A
      TabIndex        =   24
      Top             =   150
      Width           =   300
   End
   Begin VB.TextBox TXTRASOCLI 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   450
      Width           =   3225
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   " R. Social:"
      Top             =   450
      Width           =   900
   End
   Begin VB.TextBox TXTCTACLIEN 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   150
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   " Cta Clte:"
      Top             =   150
      Width           =   900
   End
   Begin VB.CommandButton Command9 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14400
      Picture         =   "REPOEQUI.frx":0454
      TabIndex        =   18
      ToolTipText     =   "Desplega Lista de Selección de Marca, Tipo de Equipo y Modelo"
      Top             =   120
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   4150
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   " Marca"
      Top             =   150
      Width           =   3400
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   7530
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   " Tipo de Equipo"
      Top             =   150
      Width           =   3400
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   " Modelo"
      Top             =   150
      Width           =   3400
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   " Motor"
      Top             =   780
      Width           =   3660
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   11020
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " Serie Motor"
      Top             =   780
      Width           =   3660
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   " Interno"
      Top             =   780
      Width           =   3660
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   " N° Serie"
      Top             =   780
      Width           =   3660
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
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
      Left            =   12960
      TabIndex        =   4
      Top             =   8040
      Width           =   1600
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
      Left            =   3720
      OleObjectBlob   =   "REPOEQUI.frx":075E
      Top             =   4080
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3600
      OleObjectBlob   =   "REPOEQUI.frx":0992
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6405
      Left            =   80
      TabIndex        =   5
      Top             =   1440
      Width           =   16875
      _ExtentX        =   29766
      _ExtentY        =   11298
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
   Begin VB.Label LBLID 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3260
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu mnuacc 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MNUABMEQUIPOS 
         Caption         =   "ABM EQUIPOS "
      End
   End
End
Attribute VB_Name = "REPOEQUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub BOTEXIT_Click()
  Unload Me
End Sub



Private Sub Combo3_Change(Index As Integer)
    Call CARGAGRILLAEQUICLI(Combo3(0).TEXT, Combo3(1).TEXT, Combo3(2).TEXT, Combo3(3).TEXT, Combo3(4).TEXT, Combo3(5).TEXT, Combo3(6).TEXT)
End Sub

Private Sub Combo3_Click(Index As Integer)
    Select Case Index
    Case 0
      Me.TXTMARCA = Combo3(Index).TEXT
    Case 1
      Me.TXTTIPOEQUIPO = Combo3(Index).TEXT
    Case 2
      Me.TXTMODELO = Combo3(Index).TEXT
    Case 3
      Me.TXTSERIE = Combo3(Index).TEXT
    Case 4
      Me.TXTINTERNO = Combo3(Index).TEXT
    Case 5
      Me.TXTMOTOR = Combo3(Index).TEXT
    Case 6
      Me.TXTSERIEMOTOR = Combo3(Index).TEXT
    End Select
    
    Call CARGAGRILLAEQUICLI(TXTMARCA, TXTTIPOEQUIPO, TXTMODELO.TEXT, TXTSERIE, TXTINTERNO, TXTMOTOR, TXTSERIEMOTOR)
End Sub

Private Sub Combo3_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      On Error Resume Next
      If Index < Combo3.UBound Then
         Combo3(Index + 1).SetFocus
      Else
         Combo3(0).SetFocus
      End If
      On Error GoTo 0
   End If
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL(AREXPO$, MSF)
   End If
   '

99 Command1.Enabled = True

End Sub
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
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
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
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
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub






Private Sub Command2_Click()

End Sub

Private Sub COMMAND3_Click()
   Command3.Enabled = False
   Call SELECHO("DEUDOR1")
   NC = XVALO(VALACT1$("deudor1"))
   If NC < 1 Then Exit Sub
   '
   TXTCTACLIEN = NC
   Command3.Enabled = True
End Sub

Private Sub Command43_Click()
   
   Me.TXTMARCA = ""
   Me.TXTMODELO = ""
   Me.TXTTIPOEQUIPO = ""
   Me.LBLID = ""
   Me.TXTMOTOR = ""
   Me.TXTSERIE = ""
   Me.TXTSERIEMOTOR = ""
   Me.TXTINTERNO = ""
   TXTCTACLIEN = ""
   Me.TXTRASOCLI = ""
   Call CARGAGRILLAEQUICLI
   
   On Error Resume Next
   Me.TXTSERIE.SetFocus
   On Error GoTo 0

End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
   Call PoneEnCeroNull("EQUIPOAHP", "STATUS", REGAF&)
   CONDI$ = "ID > 0 AND STATUS >= 0"
   Call CARGALISEL("EQUICLIE", "EQUIPOAHP", "ID/F7;Marca/A32;T_EQUIPO/A32;MODELO/A32", CONDI$)
   Call SELECHO("EQUICLIE/LISTA DE COMBINACIONES")
   NCOMBI& = XVALO(VALACT1$("EQUICLIE"))
   FILASTRING$ = VALACT2$("EQUICLIE")
   MARCAX$ = TRIM$(Mid$(FILASTRING$, 1, 32))
   TIPEQUI$ = TRIM$(Mid$(FILASTRING$, 34, 32))
   MODELO$ = TRIM$(Mid$(FILASTRING$, 67, 32))
   Me.TXTMARCA = MARCAX$
   Me.TXTTIPOEQUIPO = TIPEQUI$
   Me.TXTMODELO = MODELO$
   Me.LBLID = NCOMBI&
   
    Command9.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
   Call CARGA_ENCABEZADO(MSF, "ID/F0;Marca/A32;Tipo de Equipo/A32;Modelo/A32;Número Serie/A20;Interno/A12;Motor/A20;Serie Motor/A24;IDCOMBI/F0;Cta/f6;Razon Social/A24;Mail/A64", Me)
   For i& = 0 To Combo3.UBound
     Combo3(i&).Clear
   Next i&
   Call CARGAGRILLAEQUICLI("CARGARL0SC0MB0S")

End Sub










Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub

Private Sub TXTCTACLI_Change(Index As Integer)
End Sub

Private Sub LBLID_Change()
   SQLX$ = "SELECT * FROM EQUIPOAHP WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ID = " & XVALO(LBLID)
   Set RS = READSET(SQLX$)
   With RS
        Me.TXTMARCA = SAFETEXT$(!MARCA)
        Me.TXTTIPOEQUIPO = SAFETEXT$(!T_EQUIPO)
        Me.TXTMODELO = SAFETEXT$(!MODELO)
   End With
   RS.Close
   Set RS = Nothing
End Sub

Private Sub MNUABMEQUIPOS_Click()
  ABMEQUIPO.Show 1
End Sub





Private Sub mnuExportarExcel_Click()
  Call Command1_Click
End Sub

Private Sub TXTCTACLIEN_Change()
    If XVALO(Me.TXTCTACLIEN) > 0 Then
      Me.TXTRASOCLI = rasocli$(XVALO(Me.TXTCTACLIEN))
        Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)
        Exit Sub
    Else
      Me.TXTRASOCLI = ""
      MSF.Rows = 1
    End If
    '
    Call CARGAGRILLAEQUICLI(Combo3(0).TEXT, Combo3(1).TEXT, Combo3(2).TEXT, Combo3(3).TEXT, Combo3(4).TEXT, Combo3(5).TEXT, Combo3(6).TEXT)
End Sub

Private Sub TXTCTACLIEN_DblClick()
  Me.TXTCTACLIEN = ""
End Sub


Private Sub TXTINTERNO_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTINTERNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTMOTOR)
    End If
End Sub

Private Sub TXTMARCA_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)
End Sub

Private Sub TXTMARCA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call PINTATEXT(Me.TXTTIPOEQUIPO)
   End If
End Sub

Private Sub TXTMODELO_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTMODELO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call PINTATEXT(Me.TXTSERIE)
   End If

End Sub

Private Sub TXTMOTOR_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTMOTOR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTSERIEMOTOR)
    End If
End Sub

Sub CARGAGRILLAEQUICLI(Optional MARCA$, Optional TIPOEQUI$, Optional MODELO$, Optional SERIE$, Optional INTERNO$, Optional MOTOR$, Optional SERIEMOTOR$)
     MSF.Rows = 1
     CARCOMBO% = 0
     'VA FILTRANDO DE ACUERDO A LOS DATOS QUE HAYA O NO EN LAS CAJAS DE TEXTO
     MARCA$ = TRIM$(MARCA$): TIPOEQUI$ = TRIM$(TIPOEQUI$): MODELO$ = TRIM$(MODELO$)
     If MARCA$ = "CARGARL0SC0MB0S" Then MARCA$ = "": CARCOMBO% = 1
     SERIE$ = TRIM$(SERIE$): INTERNO$ = TRIM$(INTERNO$): MOTOR$ = TRIM$(MOTOR$): SERIEMOTOR$ = TRIM$(SERIEMOTOR$)
     Call PoneEnCeroNull("EQUICLI", "STATUS", REGAF&, "NOMESS")
     SQLX$ = "SELECT E.MARCA,E.T_EQUIPO,E.MODELO,C.IDCOMBI,C.NUMEROSERIE,C.INTERNO,C.MOTOR,C.SERIEMOTOR,E.ID,C.IDEQUIPO, C.NUCLIEN "
     SQLX$ = SQLX$ + "  FROM EQUICLI C WITH(NOLOCK) INNER JOIN EQUIPOAHP E ON C.IDCOMBI = E.ID "
     SQLX$ = SQLX$ + " WHERE E.ID >= 0 AND E.STATUS >= 0"
     SQLX$ = SQLX$ + " AND C.STATUS >= 0"
     
     If XVALO(Me.TXTCTACLIEN) > 0 Then SQLX$ = SQLX$ + " AND C.NUCLIEN =  " & XVALO(Me.TXTCTACLIEN)
     If TRIM$(MARCA$) <> "" Then SQLX$ = SQLX$ + " AND E.MARCA = '" & MARCA$ & "'"
     If TRIM$(TIPOEQUI$) <> "" Then SQLX$ = SQLX$ + " AND E.T_EQUIPO = '" & TIPOEQUI$ & "'"
     If TRIM$(MODELO$) <> "" Then SQLX$ = SQLX$ + " AND E.MODELO = '" & MODELO$ & "'"
     '
     If TRIM$(SERIE$) <> "" Then SQLX$ = SQLX$ + " AND C.NUMEROSERIE = '" & SERIE$ & "'"
     If TRIM$(INTERNO$) <> "" Then SQLX$ = SQLX$ + " AND C.INTERNO = '" & INTERNO$ & "'"
     If TRIM$(MOTOR$) <> "" Then SQLX$ = SQLX$ + " AND C.MOTOR = '" & MOTOR$ & "'"
     If TRIM$(SERIEMOTOR$) <> "" Then SQLX$ = SQLX$ + " AND C.SERIEMOTOR = '" & SERIEMOTOR$ & "'"
     '
     SQLX$ = SQLX$ + " ORDER BY E.MARCA,E.T_EQUIPO,E.MODELO,C.NUMEROSERIE,C.INTERNO,C.MOTOR,C.SERIEMOTOR"
     Call VectCliente.cargarvector
     MSF.Redraw = False
     Set RS = READSET(SQLX$)
     With RS
       
        Do While Not .EOF
           j& = j& + 1
           MSF.Rows = j& + 1
           MSF.TextMatrix(j&, 1) = XVALO(!IDCOMBI)
           MSF.TextMatrix(j&, 2) = SAFETEXT$(!MARCA): If CARCOMBO% > 0 Then Combo3(0).AddItem SAFETEXT$(!MARCA)
           MSF.TextMatrix(j&, 3) = SAFETEXT$(!T_EQUIPO): If CARCOMBO% > 0 Then Combo3(1).AddItem SAFETEXT$(!T_EQUIPO)
           MSF.TextMatrix(j&, 4) = SAFETEXT$(!MODELO): If CARCOMBO% > 0 Then Combo3(2).AddItem SAFETEXT$(!MODELO)
           
           MSF.TextMatrix(j&, 5) = SAFETEXT$(!NUMEROSERIE): If CARCOMBO% > 0 Then Combo3(3).AddItem SAFETEXT$(!NUMEROSERIE)
           MSF.TextMatrix(j&, 6) = SAFETEXT$(!INTERNO): If CARCOMBO% > 0 Then Combo3(4).AddItem SAFETEXT$(!INTERNO)
           MSF.TextMatrix(j&, 7) = SAFETEXT$(!MOTOR): If CARCOMBO% > 0 Then Combo3(5).AddItem SAFETEXT$(!MOTOR)
           MSF.TextMatrix(j&, 8) = SAFETEXT$(!SERIEMOTOR): If CARCOMBO% > 0 Then Combo3(6).AddItem SAFETEXT$(!SERIEMOTOR)
           MSF.TextMatrix(j&, 9) = XVALO(!IDEQUIPO)
           NC = XVALO(!NUCLIEN)
           MSF.TextMatrix(j&, 10) = NC
           MSF.TextMatrix(j&, 11) = VectCliente.nombre(NC)
           MSF.TextMatrix(j&, 12) = VectCliente.mailcliente(NC)
           
           
'           Combo3(0).AddItem SAFETEXT$(!MARCA)
'           Combo3(1).AddItem SAFETEXT$(!T_EQUIPO)
'           Combo3(2).AddItem SAFETEXT$(!MODELO)
'
'           Combo3(3).AddItem SAFETEXT$(!NUMEROSERIE)
'           Combo3(4).AddItem SAFETEXT$(!INTERNO)
'           Combo3(5).AddItem SAFETEXT$(!MOTOR)
'           Combo3(6).AddItem SAFETEXT$(!SERIEMOTOR)

          .MoveNext
        Loop
     End With
     RS.Close
     Set RS = Nothing
     'DEPURO LOS DATOS REPETIDOS DE LOS COMBOS

    For r& = 0 To Combo3.UBound
     Call Duplicados1(Combo3(r&))
    Next r&
    '
    Call COLOREAR_GRILLASQL(MSF, &HEFF1F1)

99  MSF.Redraw = True

End Sub



Private Sub TXTSERIE_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTSERIE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTINTERNO)
    End If
End Sub

Private Sub TXTSERIEMOTOR_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTSERIEMOTOR_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     On Error Resume Next
     Command2.SetFocus
     On Error GoTo 0
  End If
End Sub

Private Sub TXTTIPOEQUIPO_Change()
    Call CARGAGRILLAEQUICLI(Me.TXTMARCA, Me.TXTTIPOEQUIPO, Me.TXTMODELO, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub


Private Sub TXTTIPOEQUIPO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call PINTATEXT(Me.TXTMODELO)
   End If

End Sub
