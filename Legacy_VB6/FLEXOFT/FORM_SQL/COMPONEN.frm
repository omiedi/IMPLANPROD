VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form COMPONEN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valorización"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command1 
      Caption         =   "O.K."
      Height          =   810
      Left            =   12600
      Picture         =   "COMPONEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   750
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   5490
      Left            =   12360
      ScaleHeight     =   5430
      ScaleWidth      =   1080
      TabIndex        =   7
      Top             =   0
      Width           =   1140
      Begin VB.CommandButton Command3 
         Caption         =   "Salir"
         Height          =   640
         Left            =   240
         Picture         =   "COMPONEN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   240
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   8
         Top             =   4770
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -195
         Picture         =   "COMPONEN.frx":0454
         Top             =   5050
         Width           =   1515
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   5715
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Escape para Ocultar"
      Top             =   0
      Visible         =   0   'False
      Width           =   6275
      _ExtentX        =   11060
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
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
      Left            =   10680
      TabIndex        =   3
      Top             =   6080
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "COMPONEN.frx":2376
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "COMPONEN.frx":25AA
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "COMPONEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI1X%
Public OKK%

Private Sub Command7_Click()

   'FECHAS DESDE Y HASTA TOMADAS DEL CAPTION
   '***************************************
   POS1% = 7 + InStr(INFO_GRILLA07.Caption, "Desde:")
   FE1$ = Mid$(INFO_GRILLA07.Caption, POS1%, 8)
   DESDEX% = FECHANUM(FE1$)
   '
   pos2% = 7 + InStr(INFO_GRILLA07.Caption, "Hasta:")
   FE2$ = Mid$(INFO_GRILLA07.Caption, pos2%, 8)
   '
   HASTAX% = FECHANUM(FE2$)

   Call ANACOSI09.PRINTINFOGRILLA07(CI1X%, DESDEX%, HASTAX%)
End Sub

Private Sub command1_Click()
  OKK%
  Hide
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OKK% = 0
   Hide
   Command3.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
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
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(ARCHIEX$)
   End If
   
99 Command8.Enabled = True

End Sub
Sub GENERAXCEL(NOMBREDEARCHIVO$)
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
   For I& = 1 To MSF.Rows
       RENG1& = I&
       For COLU1& = 1 To MSF.Cols
          VALORX = MSF.TextMatrix(I& - 1, COLU1 - 1)
          If IsNumeric(VALORX) = True Then
             If InStr(VALORX, ".") > 0 Then
                VALO$ = VALORX
                VALO$ = REPLACAR(VALO$, ".", "")
                VALORX = CDbl(VALO$)
             End If
          End If

          '
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next I&
   '
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$

   XLAPP.Visible = True
   On Error GoTo 0
   '

   XLAPP.QUIT      ' cierra EXCEL que se esta ejecutando minimizado

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
   XLAPP.QUIT ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MSF2.Visible = False
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel > 0 Then
      Cancel = 0
      OKK% = 0
      Hide
   End If
End Sub

Private Sub MSF_DblClick()
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   I& = MSF.Row
   j& = MSF.Col
   '
   If I& < 1 Then Exit Sub
   '
   If j& <> 1 Then Exit Sub
   
   MSF2.Visible = True

   MSF2.Rows = MSF.Cols
   MSF2.Cols = 2
   'VALIDO
   If I& = 0 Or I& > MSF.Rows Then Exit Sub
   If j& > MSF.Cols Then Exit Sub
   '
   NUCOLU% = 1: TEPRUEBA = "123456789012345678902345678901234": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 2: TEPRUEBA = "123456789012345678902345678901234": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "Cliente"
   '
   'CARGA EL ENCABEZADO A MODO VERTICAL
   For H& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(0, H&)
      MSF2.TextMatrix(H&, 0) = TRIM$(DATO1$)
   Next H&
   '
   'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
   For H& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(I&, H&)
      MSF2.TextMatrix(H&, 1) = TRIM$(DATO1$)
   Next H&
   If MSF2.Rows < 24 Then MSF2.Rows = 24
   Call COLOREAR_GRILLASQL(MSF2, &HEFF1F1)
End Sub



