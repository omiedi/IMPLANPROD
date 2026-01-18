VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ABMEQUICLI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Equipos de Clientes"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   14775
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   465
      Left            =   12360
      Picture         =   "ABMEQUICLI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Refrescar"
      Top             =   5430
      Width           =   405
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
      Picture         =   "ABMEQUICLI.frx":014A
      TabIndex        =   27
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12840
      Picture         =   "ABMEQUICLI.frx":0454
      TabIndex        =   26
      Top             =   5430
      Width           =   1860
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
      TabIndex        =   24
      Top             =   420
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
      TabIndex        =   23
      Text            =   " R. Social:"
      Top             =   420
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
      TabIndex        =   22
      Top             =   120
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
      TabIndex        =   21
      Text            =   " Cta Clte:"
      Top             =   120
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
      Picture         =   "ABMEQUICLI.frx":075E
      TabIndex        =   20
      Top             =   120
      Width           =   300
   End
   Begin VB.TextBox TXTMARCA 
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
      Left            =   4150
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   420
      Width           =   3400
   End
   Begin VB.TextBox TXTTIPOEQUIPO 
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
      Left            =   7530
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   420
      Width           =   3400
   End
   Begin VB.TextBox TXTMODELO 
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
      Left            =   10920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   420
      Width           =   3400
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
      Top             =   120
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
      Top             =   120
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
      Top             =   120
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
      Width           =   3660
   End
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
      Width           =   3660
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
      Width           =   3660
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
      Top             =   6240
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
      OleObjectBlob   =   "ABMEQUICLI.frx":0A68
      Top             =   4080
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3600
      OleObjectBlob   =   "ABMEQUICLI.frx":0C9C
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4005
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   7064
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   ""
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
      TabIndex        =   25
      Top             =   120
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
   End
   Begin VB.Menu mnuacc 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MNUABMEQUIPOS 
         Caption         =   "ABM EQUIPOS "
      End
   End
End
Attribute VB_Name = "ABMEQUICLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
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
     If MSF_2.Visible = True Then
      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
     Else
      Call GENEXCEL(AREXPO$, MUESTRADATOS.GRID)
     End If
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
   Command2.Enabled = False
'   Call BORRAREGISTROS("EQUIPOAHP", "ID > 0", REGAF&)
   MARCAX$ = TRIM$(UCase$(TXTMARCA))
   TIPOEQUI$ = TRIM$(UCase$(Me.TXTTIPOEQUIPO))
   MODELOX$ = TRIM$(UCase$(Me.TXTMODELO))
   If MARCAX$ = "" Then
     Call COMUNI("Falta Marca")
     GoTo 99
   End If
   '
   If TIPOEQUI$ = "" Then
     Call COMUNI("Falta Tipo de Equipo")
     GoTo 99
   End If

   If MODELOX$ = "" Then
     Call COMUNI("Falta Modelo")
     GoTo 99
   End If
   '
   
   CONDI$ = "NUMEROSERIE = '" & TRIM$(Me.TXTSERIE) & "'"
   CONDI$ = CONDI$ + " AND INTERNO = '" & TRIM$(Me.TXTINTERNO) & "'"
   CONDI$ = CONDI$ + " AND MOTOR = '" & TRIM$(Me.TXTMOTOR) & "'"
   CONDI$ = CONDI$ + " AND SERIEMOTOR = '" & TRIM$(Me.TXTSERIEMOTOR) & "'"
   CONDI$ = CONDI$ + " AND IDCOMBI = " & XVALO(Me.LBLID)
   CONDI$ = CONDI$ + " AND NUCLIEN = " & XVALO(Me.TXTCTACLIEN)
   
   ID1& = XVALO(VALOBADA("EQUICLI", "IDEQUIPO", CONDI$, "NOMESS"))
   '
   
   If ID1& < 1 Then
50    NUEVO_ID& = 1 + XVALO(VALOBADA("EQUICLI", "MAX(IDEQUIPO)", " IDEQUIPO > 0 ", "NOMESS"))
      SQLX$ = "INSERT INTO EQUICLI (NUMEROSERIE, INTERNO, MOTOR,SERIEMOTOR,IDEQUIPO,NUCLIEN,IDCOMBI,STATUS)"
      SQLX$ = SQLX$ + " VALUES ( "
      SQLX$ = SQLX$ + " '" & TRIM$(Me.TXTSERIE) & "' "
      SQLX$ = SQLX$ + ", '" & TRIM$(Me.TXTINTERNO) & "'"
      SQLX$ = SQLX$ + ", '" & TRIM$(Me.TXTMOTOR) & "'"
      SQLX$ = SQLX$ + ", '" & TRIM$(Me.TXTSERIEMOTOR) & "'"
      SQLX$ = SQLX$ + ", " & NUEVO_ID&
      SQLX$ = SQLX$ + ", " & XVALO(Me.TXTCTACLIEN)
      SQLX$ = SQLX$ + ", " & XVALO(Me.LBLID)
      SQLX$ = SQLX$ + ", 0"
      
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
   Else
60    SQLX$ = "UPDATE EQUICLI SET NUMEROSERIE = '" & TRIM$(Me.TXTSERIE) & "'"
      SQLX$ = SQLX$ + ", INTERNO = '" & TRIM$(Me.TXTINTERNO) & "'"
      SQLX$ = SQLX$ + ", MOTOR = '" & TRIM$(Me.TXTMOTOR) & "'"
      SQLX$ = SQLX$ + ", SERIEMOTOR = '" & TRIM$(Me.TXTSERIEMOTOR) & "'"
      SQLX$ = SQLX$ + ", IDCOMBI = " & XVALO(Me.LBLID)
      SQLX$ = SQLX$ + " WHERE IDEQUIPO = " & ID1&
      FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CARGAGRILLAEQUICLI
   
   Me.TXTMARCA = ""
   Me.TXTMODELO = ""
   Me.TXTTIPOEQUIPO = ""
   Me.LBLID = ""
   Me.TXTMOTOR = ""
   Me.TXTSERIE = ""
   Me.TXTSERIEMOTOR = ""
   Me.TXTINTERNO = ""
   
   On Error Resume Next
   Me.TXTSERIE.SetFocus
   On Error GoTo 0
99 Command2.Enabled = True

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
   Call CARGAGRILLAEQUICLI
   
   On Error Resume Next
   Me.TXTSERIE.SetFocus
   On Error GoTo 0

End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
   Call PoneEnCeroNull("EQUIPOAHP", "STATUS", REGAF&)
   CONDI$ = "ID > 0 AND STATUS >= 0 ORDER BY MODELO ASC"
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
   Call CARGA_ENCABEZADO(MSF, "ID/F0;Marca/A32;Tipo de Equipo/A32;Modelo/A32;Número Serie/A20;Interno/A12;Motor/A20;Serie Motor/A24;IDCOMBI/F0", Me)
   Call CARGAGRILLAEQUICLI

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = J&
    On Error Resume Next

    NPX& = TRIM$(GRID.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If
  End If
End Sub
Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (msf2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            msf2.COL = msf2.MouseCol
            msf2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            msf2.COL = msf2.MouseCol
            msf2.Sort = 1
            MODO = True
        End If
    End If

End Sub

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = msf2.MouseRow
    J& = msf2.MouseCol
    If i& = 0 Or i& > msf2.Rows Then Exit Sub
    '
    On Error Resume Next
    msf2.Row = i&
    msf2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(msf2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If
  End If
End Sub


Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
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
   Set rs = READSET(SQLX$)
   With rs
        Me.TXTMARCA = SAFETEXT$(!MARCA)
        Me.TXTTIPOEQUIPO = SAFETEXT$(!T_EQUIPO)
        Me.TXTMODELO = SAFETEXT$(!MODELO)
   End With
   rs.Close
   Set rs = Nothing
   Call CARGAGRILLAEQUICLI
End Sub

Private Sub MNUABMEQUIPOS_Click()
  ABMEQUIPO.Show 1
End Sub

Private Sub MSF_DblClick()
     REG& = XVALO(MSF.MouseRow)
     SERIEX$ = MSF.TextMatrix(REG&, 5)
     INTERNOX$ = MSF.TextMatrix(REG&, 6)
     MOTORX$ = MSF.TextMatrix(REG&, 7)
     SERIEMOTOR$ = MSF.TextMatrix(REG&, 8)
     IDEQUIPO = MSF.TextMatrix(REG&, 9)
     Me.LBLID = MSF.TextMatrix(REG&, 1)
     '
     Me.TXTSERIE = SERIEX$
     Me.TXTINTERNO = INTERNOX$
     Me.TXTMOTOR = MOTORX$
     Me.TXTSERIEMOTOR = SERIEMOTOR$
     
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        REG& = MSF.MouseRow
        COL& = MSF.MouseCol
        If REG& < 1 Then Exit Sub
        MSF.TopRow = REG&
        MSF.Row = REG&
        MSF.ColSel = 1
        On Error Resume Next
        MSF.SetFocus
        On Error GoTo 0
     '
        IDEQUI& = MSF.TextMatrix(REG&, 9)
        OPX% = COMALTER%("Desea Borrar El Item Seleccionado\\Cancelar\Modificar\Eliminar")
        If OPX% < 2 Then Exit Sub
        If OPX% = 3 Then
           Call BORRAEQUICLI(IDEQUI&)
           Call CARGAGRILLAEQUICLI
           Exit Sub
        End If

        ORIGINALWIDTH = 7185
        REPROGRID.Height = 1000
        'AGREGO LOS CAJAS DE TEXTOS Y LOS LABEL CON SU ROTULO
        Dim FORMX$(), NOCAMPO$(), EDITABLE$(), MEMORIZA$()
        ReDim Preserve MEMORIZA$(1)
        ReDim Preserve NOCAMPO$(1), FORMX$(1)
        '
10      CANCA% = 0
        CAMPOS$ = "ID_COMBI/F0;Marca/A32;Tipo de Equipo/A32;Modelo/A32;Número Serie/A20@;Interno/A12@;Motor/A20@;Serie Motor/A24@;ID_EQUIPO/F0"
        Call TEXTOLOTES$(CAMPOS$, ";")
        
        If UBound(CADENATEX$) > 40 Then
           Call COMUNI("Demasiados Campos (Máximo 40)")
           Exit Sub
        End If
        
        For i& = 1 To UBound(CADENATEX$)
           TEXLEIDO$ = (CADENATEX$(i&))
           PPXX% = InStr(TEXLEIDO$, "/")
           If PPXX% < 1 Then PPXX% = 1 + Len(TEXLEIDO$)
           CANCA% = CANCA% + 1
           If CANCA% >= UBound(NOCAMPO$) Then
             ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%), EDITABLE$(CANCA%)
           End If
           '
           NOCAMPO$(CANCA%) = Left$(TEXLEIDO$, PPXX% - 1)
           FORMX$(CANCA%) = Mid$(TEXLEIDO$, PPXX% + 1)
           EDITABLE(CANCA%) = Right(TEXLEIDO$, 1)
           Load REPROGRID.TXTLABEL(CANCA%)
           Load REPROGRID.TXTCAJA(CANCA%)
           REPROGRID.TXTLABEL(CANCA%) = NOCAMPO$(CANCA%)
           REPROGRID.TXTCAJA(CANCA%).TEXT = ""
           FORMAXTOX$ = UCase$(Left$(FORMX$(CANCA%), 1))
           If FORMAXTOX$ = "A" Or FORMAXTOX$ = "D" Then
              Align1% = 0
           ElseIf FORMAXTOX$ = "F" Or FORMAXTOX$ = "I" Or FORMAXTOX$ = "G" Then
              Align1% = 1
           Else
              Align1% = 2
           End If
           REPROGRID.TXTCAJA(CANCA%).Alignment = Align1%
           If CANCA% > 1 And CANCA% <= 20 Then
              REPROGRID.TXTCAJA(CANCA%).Top = REPROGRID.TXTCAJA(CANCA% - 1).Top + 400
              REPROGRID.TXTLABEL(CANCA%).Top = REPROGRID.TXTLABEL(CANCA% - 1).Top + 400
           ElseIf CANCA% > 20 Then
              If CANCA% = 21 Then
                 REPROGRID.Width = (ORIGINALWIDTH * 2)
              End If
              REPROGRID.TXTCAJA(CANCA%).Top = REPROGRID.TXTCAJA(CANCA% - 20).Top
              REPROGRID.TXTLABEL(CANCA%).Top = REPROGRID.TXTLABEL(CANCA% - 20).Top
              REPROGRID.TXTLABEL(CANCA%).Left = REPROGRID.TXTLABEL(CANCA%).Width + REPROGRID.TXTCAJA(CANCA%).Width + 250
              REPROGRID.TXTCAJA(CANCA%).Left = (REPROGRID.TXTLABEL(CANCA%).Width * 2) + REPROGRID.TXTCAJA(CANCA%).Width + 300
           End If
           '
           If EDITABLE(CANCA%) = "@" Then
              REPROGRID.TXTCAJA(CANCA%).Locked = False
              REPROGRID.TXTCAJA(CANCA%).BackColor = &HFFFFFF
           End If
           '
           REPROGRID.TXTCAJA(CANCA%).Visible = True
           REPROGRID.TXTLABEL(CANCA%).Visible = True
           '
           REPROGRID.Height = REPROGRID.Height + 400
        Next i&
        '
        '
        'AGREGO EL VALOR A LAS CAJAS DE TEXTO
        'SI VIENE REBOTADO DE VALIDACION MUESTRA LOS DATOS DE LAS CAJAS DE TEXTO (LEYENDO EL VECTOR) CASO CONTRARIO TOMA LA DE LA GRILLA.
        If UBound(MEMORIZA$) > 1 Then
            For i& = 1 To REPROGRID.TXTCAJA.Count - 1
              REPROGRID.TXTCAJA(i&) = MEMORIZA(i&)
            Next i&
            Erase MEMORIZA$
        Else
            For i& = 1 To MSF.Cols - 1
              REPROGRID.TXTCAJA(i&) = MSF.TextMatrix(REG&, i&)
            Next i&
        End If
        
        REPROGRID.Show 1
        If REPROGRID.GUARDAR = 1 Then
           Call GUARDAR_INFO_EQUI(OK%)
           If OK% >= 0 Then
             Call ACTUALIZA_FILA(MSF, REG&)
             Unload REPROGRID
           Else
             For i& = 1 To REPROGRID.TXTCAJA.Count - 1
               'GUARDO LOS DATOS EN UN VECTOR PARA MOSTRAR LA INFORMACION CORREGIDA EN EL FRM REPROGRID Y LO VUELVA A MOSTRAR EN EL CASO QUE OK% SEA NEGATIVO
               ReDim Preserve MEMORIZA(i&)
               MEMORIZA(i&) = REPROGRID.TXTCAJA(i&)
               Unload REPROGRID.TXTLABEL(i&)
               Unload REPROGRID.TXTCAJA(i&)
             Next i&
             GoTo 10
           End If
        End If
     End If
End Sub
Sub GUARDAR_INFO_EQUI(OK%)
     OK% = -1
     '
    Dim DATO$()
    ReDim Preserve DATO$(1)
    For i& = 1 To REPROGRID.TXTCAJA.UBound
       If UBound(DATO$) < i& Then ReDim Preserve DATO$(i&)
       DATO$(i&) = REPROGRID.TXTCAJA(i&)
    Next i&

    SQLX$ = "UPDATE EQUICLI SET "
    SQLX$ = SQLX$ + " NUMEROSERIE = '" & TRIM$(DATO$(5)) & "'"
    SQLX$ = SQLX$ + ", INTERNO = '" & TRIM$(DATO$(6)) & "'"
    SQLX$ = SQLX$ + ", MOTOR = '" & TRIM$(DATO$(7)) & "'"
    SQLX$ = SQLX$ + ", SERIEMOTOR = '" & TRIM$(DATO$(8)) & "'"
    SQLX$ = SQLX$ + " WHERE IDCOMBI = " & XVALO(DATO$(1))
    SQLX$ = SQLX$ + " AND NUCLIEN = " & XVALO(TXTCTACLIEN)
    SQLX$ = SQLX$ + " AND IDEQUIPO = " & XVALO(DATO$(9))

    FLEXCONN.Execute (SQLX$)
    '
    OK% = 1
    Call CARGAGRILLAEQUICLI
       
End Sub



Sub BORRAEQUICLI(IDX&)
    CONDI$ = "IDEQUIPO = " & IDX&
    Call BORRAREGISTROS("EQUICLI", CONDI$, REGAF&, "NOMESS")
End Sub


Private Sub TXTCTACLIEN_Change()
    If XVALO(Me.TXTCTACLIEN) > 0 Then
      Me.TXTRASOCLI = rasocli$(XVALO(Me.TXTCTACLIEN))
      Call CARGAGRILLAEQUICLI
    Else
      Me.TXTRASOCLI = ""
      MSF.Rows = 1
    End If
    '
    Call CARGAGRILLAEQUICLI
End Sub

Private Sub TXTINTERNO_Change()
    Call CARGAGRILLAEQUICLI(Me.LBLID, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTINTERNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTMOTOR)
    End If
End Sub

Private Sub TXTMOTOR_Change()
    Call CARGAGRILLAEQUICLI(Me.LBLID, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTMOTOR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTSERIEMOTOR)
    End If
End Sub

Private Sub TXTSERIE_Change()

    Call CARGAGRILLAEQUICLI(Me.LBLID, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub
Sub CARGAGRILLAEQUICLI(Optional IDCOMBI, Optional SERIE$, Optional INTERNO$, Optional MOTOR$, Optional SERIEMOTOR$)
     MSF.Rows = 1
     'VA FILTRANDO DE ACUERDO A LOS DATOS QUE HAYA O NO EN LAS CAJAS DE TEXTO
     SERIE$ = TRIM$(SERIE$): INTERNO$ = TRIM$(INTERNO$): MOTOR$ = TRIM$(MOTOR$): SERIEMOTOR$ = TRIM$(SERIEMOTOR$)
     Call PoneEnCeroNull("EQUICLI", "STATUS", REGAF&, "NOMESS")
     
     SQLX$ = "SELECT E.MARCA,E.T_EQUIPO,E.MODELO,C.IDCOMBI,C.NUMEROSERIE,C.INTERNO,C.MOTOR,C.SERIEMOTOR,E.ID,C.IDEQUIPO "
     SQLX$ = SQLX$ + "  FROM EQUICLI C WITH(NOLOCK) INNER JOIN EQUIPOAHP E ON C.IDCOMBI = E.ID "
     SQLX$ = SQLX$ + " WHERE E.ID > 0 AND E.STATUS >= 0"
     SQLX$ = SQLX$ + " AND C.STATUS >= 0"
     SQLX$ = SQLX$ + " AND C.NUCLIEN =  " & XVALO(Me.TXTCTACLIEN)
     If XVALO(LBLID) > 0 Then SQLX$ = SQLX$ + " AND E.ID = " & XVALO(LBLID)
     If TRIM$(SERIE$) <> "" Then SQLX$ = SQLX$ + " AND C.NUMEROSERIE LIKE '" & SERIE$ & "%'"
     If TRIM$(INTERNO$) <> "" Then SQLX$ = SQLX$ + " AND C.INTERNO LIKE '" & INTERNO$ & "%'"
     If TRIM$(MOTOR$) <> "" Then SQLX$ = SQLX$ + " AND C.MOTOR LIKE '" & MOTOR$ & "%'"
     If TRIM$(SERIEMOTOR$) <> "" Then SQLX$ = SQLX$ + " AND C.SERIEMOTOR LIKE '" & SERIEMOTOR$ & "%'"
     
     SQLX$ = SQLX$ + " ORDER BY E.MODELO,E.MARCA,E.T_EQUIPO,C.NUMEROSERIE,C.INTERNO,C.MOTOR,C.SERIEMOTOR"
     Set rs = READSET(SQLX$)
     With rs
        Do While Not .EOF
           J& = J& + 1
           MSF.Rows = J& + 1
           MSF.TextMatrix(J&, 1) = XVALO(!IDCOMBI)
           MSF.TextMatrix(J&, 2) = SAFETEXT$(!MARCA)
           MSF.TextMatrix(J&, 3) = SAFETEXT$(!T_EQUIPO)
           MSF.TextMatrix(J&, 4) = SAFETEXT$(!MODELO)
           
           MSF.TextMatrix(J&, 5) = SAFETEXT$(!NUMEROSERIE)
           MSF.TextMatrix(J&, 6) = SAFETEXT$(!INTERNO)
           MSF.TextMatrix(J&, 7) = SAFETEXT$(!MOTOR)
           MSF.TextMatrix(J&, 8) = SAFETEXT$(!SERIEMOTOR)
           MSF.TextMatrix(J&, 9) = XVALO(!IDEQUIPO)
           
          .MoveNext
        Loop
     End With
     rs.Close
     Set rs = Nothing
End Sub


Private Sub TXTSERIE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call PINTATEXT(TXTINTERNO)
    End If
End Sub

Private Sub TXTSERIEMOTOR_Change()
    Call CARGAGRILLAEQUICLI(Me.LBLID, Me.TXTSERIE, Me.TXTINTERNO, Me.TXTMOTOR, Me.TXTSERIEMOTOR)

End Sub

Private Sub TXTSERIEMOTOR_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     On Error Resume Next
     Command2.SetFocus
     On Error GoTo 0
  End If
End Sub
