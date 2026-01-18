VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INFO_GRILLA07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Valorización"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   11685
   StartUpPosition =   2  'CenterScreen
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
      Height          =   330
      Left            =   9500
      TabIndex        =   3
      Top             =   4460
      Width           =   2115
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   3885
      Left            =   105
      TabIndex        =   0
      Top             =   525
      Width           =   11500
      _ExtentX        =   20294
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFO_GRILLA07.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFO_GRILLA07.frx":0234
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
Attribute VB_Name = "INFO_GRILLA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command8_Click()
   Command8.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
   ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
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
'   XLSHEET.Cells(1, 1) = "Código"
'   XLSHEET.Cells(1, 2) = "Descripción"
'   XLSHEET.Cells(1, 3) = "Cantidad"

'   For YX& = 1 To ULTREG&("!BOLREDET")
'     RENG1& = YX& + 1
'     YY$ = REGLEIDO$("!BOLREDET", YX&)
'     CI% = CVI(Mid$(YY$, 83, 2))
'     CODIGO1$ = CODEXT$(CI%)
'     DESCARIP1$ = DESCRIT0$(CI%)
'       CAN = CVD(Mid$(YY$, 89, 8))
'     CANTI = FORMATNUM$(CAN, "F9.1")
'     XLSHEET.Cells(RENG1&, 1) = CODIGO1$
'     XLSHEET.Cells(RENG1&, 2) = DESCARIP1$
'     XLSHEET.Cells(RENG1&, 3) = CANTI
'   Next YX&

   For I& = 1 To MSF.Rows
       RENG1& = I&
       For COLU1& = 1 To MSF.Cols
          VALORX = MSF.TextMatrix(I& - 1, COLU1 - 1)
'          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
'             FEXU$ = VALORX
'             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
'          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next I&
   '

   On Error GoTo 30
   'XLSHEET.SaveAs NOMBREDEARCHIVO$

   XLAPP.Visible = True
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



Private Sub Form_Load()
    Call APLICACIONSKINS(Me)

End Sub
