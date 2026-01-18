VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form DIARPAN07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Libro Diario General"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   10080
      TabIndex        =   2
      Top             =   0
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Busqueda por Texto"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox MARCOBARRA 
      Height          =   195
      Left            =   10300
      ScaleHeight     =   135
      ScaleWidth      =   1305
      TabIndex        =   19
      Top             =   600
      Width           =   1365
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   20
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   120
      ScaleHeight     =   6750
      ScaleWidth      =   11655
      TabIndex        =   7
      Top             =   960
      Width           =   11655
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   6800
         Left            =   1035
         TabIndex        =   14
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   6800
         Left            =   2385
         TabIndex        =   13
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   6800
         Left            =   3345
         TabIndex        =   12
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   6800
         Left            =   4815
         TabIndex        =   11
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   6800
         Left            =   8355
         TabIndex        =   10
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   6800
         Left            =   9855
         TabIndex        =   9
         Top             =   45
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   6800
         Left            =   11355
         TabIndex        =   8
         Top             =   45
         Width           =   15
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   11625
         TabIndex        =   16
         Top             =   0
         Width           =   11655
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"DIARPAN12.frx":0000
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   11655
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6330
         Left            =   0
         TabIndex        =   15
         Top             =   420
         Width           =   11655
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -120
         OleObjectBlob   =   "DIARPAN12.frx":00C7
         Top             =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EJERCICIO CONTABLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   -360
         Width           =   4455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID2 
      Height          =   765
      Left            =   1200
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   23
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8433840
      BackColorSel    =   -2147483643
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   3480
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DIARPAN12.frx":02FB
      Height          =   6975
      Left            =   120
      OleObjectBlob   =   "DIARPAN12.frx":030F
      TabIndex        =   1
      Top             =   10000
      Width           =   11655
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "DIARPAN12.frx":0CE2
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE DE LA EMPRESA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   210
      Width           =   5565
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   7920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":0D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":1076
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":1390
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":16AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":19C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":1CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":1FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":2312
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":262C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":2946
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":2C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":2E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":2F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "DIARPAN12.frx":3266
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Libro Diario General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5475
   End
   Begin VB.Menu Files 
      Caption         =   "Archivos"
      Begin VB.Menu Print 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edicion"
      Begin VB.Menu Busca 
         Caption         =   "Nueva Busqueda"
      End
      Begin VB.Menu rebusca 
         Caption         =   "Repetir Ultima        <F3>"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "DIARPAN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARBUS$, ABUSQUE$
'
Private Sub Command19_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command18_Click()
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   
   Command1.Enabled = True
End Sub

Private Sub DBGrid1_DblClick()
   'MsgBox DBGrid1.TEXT
   If DbGrid1.COL = 1 Then
     'ASIDIAR.Hide
     ACTICARGA% = 1
     ASIDIAR.Ltext2 = DbGrid1.TEXT
     On Error Resume Next
     ASIDIAR.Show 1
     If Err Then
       Hide
     End If
     On Error GoTo 0
     'ASIDIAR.Show 1
   End If
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
      Call rebusca_Click
    End If
End Sub

Private Sub Exit_Click()
    Call CIERRATO
End Sub

Sub Form_Load()
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    
    Caption = Caption + " - " + TRIM$(EJACT$)
    '
    PPXX% = InStr(EPAC$, " -"): If PPXX% < 1 Then PPXX% = 1 + Len(EPAC$)
    Label1 = Left$(EPAC$, PPXX% - 1)
    Label2 = DENOEJ$
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Private Sub Busca_Click()
   Call Search
End Sub



Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
     Call rebusca_Click
   End If
End Sub

Private Sub rebusca_Click()
   If ARBUS$ <> "" Then
     VUELTA% = 0
     INI& = LIST1.ListIndex + 2
     '
10   For U& = INI& To LIST1.ListCount
       If InStr(UCase$(LIST1.List(U& - 1)), ARBUS$) > 0 Then
         LIST1.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     '
     If VUELTA% < 1 Then
       VUELTA% = 1
       INI& = 1
       GoTo 10
     End If
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   '
   Select Case UCase$(Button.Key)
      Case "PRINT"
        Call LIDIARPRIN
        'Call List'comente esta linea xq no compilaba, xq no encuentra el sub "list"
      Case "SEARCH"
        Call Search
      Case "EXIT"
        Call CIERRATO
   End Select
   '
End Sub

Sub LIDIARPRIN()
   '
   OPX% = COMALTER%("Opciones de Salida\\Exportar a Excel\Listado Impreso")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   If OPX% = 1 Then
       GoTo 15
   ElseIf OPX% = 2 Then
       GoTo 10
   End If
   Exit Sub
   
10 DES% = FECHANUM(DIARGEN07.Label5)
   HAS% = FECHANUM(DIARGEN07.Label6)
   CORRESP$ = FECHATEX$(DES%) + " al " + FECHATEX$(HAS%)
   '
   For U& = 1 To LIST1.ListCount
     TTXX$ = LIST1.List(U& - 1)
     XX$ = REGBLAN$("LIDIARG")
       FEXI% = FECHANUM(Left$(TTXX$, 8))
     Call REPLA(XX$, MKI$(FEXI%), 1, 2)
       IDASIENTO$ = Mid$(TTXX$, 11, 12)
     Call REPLA(XX$, IDASIENTO$, 3, 12)
       NUAS& = Val(Mid$(TTXX$, 24, 7))
     Call REPLA(XX$, MKS$(NUAS&), 15, 4)
       Cuenta$ = Mid$(TTXX$, 33, 12)
     Call REPLA(XX$, Cuenta$, 19, 16)
       REFE$ = Mid$(TTXX$, 47, 32)
     Call REPLA(XX$, REFE$, 35, 32)
       IDEBE# = XVALO(Mid$(TTXX$, 80, 14))  'SE CAMBIO POR XVALO PORQUE EL IMPORTE ESTA FORMATEADO CON EL FORMATO "G"
     Call REPLA(XX$, MKD$(IDEBE#), 67, 8)
       IHABE# = XVALO(Mid$(TTXX$, 94, 14))  'SE CAMBIO POR XVALO PORQUE EL IMPORTE ESTA FORMATEADO CON EL FORMATO "G"
     Call REPLA(XX$, MKD$(IHABE#), 75, 8)
       DEMOVI$ = Mid$(TTXX$, 129, 40)
     Call REPLA(XX$, DEMOVI$, 87, 40)
     If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Or TRIM$(IDASIENTO$) = "" Then
       ' LO QUE SIGUE ES PARA SUPRIMIR MOVIMIENTOS ANULADOS EN LISTADO
       If (Abs(IDEBE#) + Abs(IHABE#)) >= 0.005 Then
         JJ& = JJ& + 1
         Call GRAREG("LIDIARG", XX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("LIDIARG", JJ&)
   Call CIERRARCH("LIDIARG")
   '
   Call HEADERS("LIDIARG", "")
   Call HEADERS("LIDIARG", "Corresponde a: " + CORRESP$)
   Screen.MousePointer = 1
   Call CONECRUN("*LIDIARG", "Libro Diario General")
   Exit Sub
   
   'presenta cuadro de dialogo
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
   ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   'On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
15 Screen.MousePointer = 11
   'ARCHIEX$ = UCase$(CommonDialog1.filename)
   ARCHIEX$ = UCase$("Diario General.xls")
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
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
'       CantidadColumnasPorHoja& = XVALO(InputBox("Columnas", "Cantidad de Columnas por Hoja", "1"))
'       If CantidadColumnasPorHoja& < 1 Or CantidadColumnasPorHoja& > 6 Then
'            Call COMUNI("Los Valores Admisibles para Columna es de 1 a 6. Por Defecto se Asigna 1 Columna")
'            CantidadColumnasPorHoja& = 1
'       End If
'       CantidadFilasPorHoja& = XVALO(InputBox("Filas", "Cantidad de Filas por Hoja", "1"))
'       If CantidadFilasPorHoja& < 0 Then
'            Call COMUNI("Los Valores Admisibles para Columna es de 1 a 6. Por Defecto se Asigna 1 Columna")
'            CantidadColumnasPorHoja& = 1
'       End If
       '
       FIN& = ULTREG&("ASDIARE")
       RENG1& = 1: IDASANT$ = ""
       For i& = 1 To FIN&
            XX$ = REGLEIDO$("ASDIARE", i&)
            FF% = CVI(Mid$(XX$, 1, 2)): fech$ = FECHATEX$(FF%)
            IDAS$ = TRIM$(Mid$(XX$, 3, 16))
            CONCEPTO$ = SAFETEXT$(VALOBADA("ENCABEASIEN", "REFEASIEN", "NUIDASIEN='" & IDAS$ & "'"))
            If IDAS$ <> IDASANT$ Then
                XLSHEET.Cells(RENG1&, 1) = String(200, "-")
                RENG1& = RENG1& + 1
                RENG1& = RENG1& + 1
                XLSHEET.Cells(RENG1&, 1) = "Concepto del Asiento: " & CONCEPTO$
                XLSHEET.Cells(RENG1&, 4) = "Fecha del Asiento: " & fech$
                RENG1& = RENG1& + 1
                XLSHEET.Cells(RENG1&, 1) = String(200, "-")
                RENG1& = RENG1& + 1
                XLSHEET.Cells(RENG1&, 1) = "Cuenta": XLSHEET.Cells(RENG1&, 2) = "Referencia"
                XLSHEET.Cells(RENG1&, 3) = "Debe": XLSHEET.Cells(RENG1&, 4) = "Haber"
                RENG1& = RENG1& + 1
                XLSHEET.Cells(RENG1&, 1) = String(200, "-")
                RENG1& = RENG1& + 1
                '
                IDASANT$ = IDAS$
            End If
            '
                Cuenta$ = Mid$(XX$, 87, 16)
            XLSHEET.Cells(RENG1&, 1) = Cuenta$
                REFE$ = Mid$(XX$, 103, 64)
            XLSHEET.Cells(RENG1&, 2) = REFE$
                DEBE# = CVD(Mid$(XX$, 167, 8))
            XLSHEET.Cells(RENG1&, 3) = FORMATNUM$(DEBE#, "F12.2")
                HABER# = CVD(Mid$(XX$, 175, 8))
            XLSHEET.Cells(RENG1&, 4) = FORMATNUM$(HABER#, "F12.2")
            RENG1& = RENG1& + 1
            '
       Next i&
       '
       On Error GoTo 30
       XLAPP.Visible = True
       On Error GoTo 0
       '
       XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
    
       Set XLAPP = Nothing     ') libera memoria
       Set XLBOOK = Nothing    ') utilizada por
       Set XLSHEET = Nothing   ') las variables objeto
       '
       Screen.MousePointer = 1
       Call COMUNI("Planilla Excel Creada Exitosamente.\ \IMPORTANTE !!!\La Exportación Incluye Unicamente a los Asientos de Diario.\Los Movimientos Administrativos Aparecen\Solo a partir de los Asientos Resumen\de Ventas, Cobranza, Compras y Pagos.")
       '
   Exit Sub
       '
       '
20         Resume 21
21         On Error GoTo 0
       MsgBox "Utilitario EXCEL no Instalado"
       On Error GoTo 0
       Exit Sub
       '
30         Resume 31
31         On Error GoTo 0
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
   End If
   '
End Sub

Sub Search()
   ABUSQUE$ = UCase$(TRIM$(InputBox$("Busqueda por Texto", , ARBUS$)))
   If ABUSQUE$ <> "" Then
     ARBUS$ = ABUSQUE$
     LIST1.ListIndex = (-1)
     For U& = 1 To LIST1.ListCount
       If InStr(UCase$(LIST1.List(U& - 1)), ABUSQUE$) > 0 Then
         LIST1.ListIndex = U& - 1
         Exit Sub
       End If
     Next U&
     '
     Call MENSERR(24, "Argumento no Encontrado.")
   End If
End Sub

Sub CIERRATO()
   Call CIERRARCH("*.*")
   Unload Me
End Sub
