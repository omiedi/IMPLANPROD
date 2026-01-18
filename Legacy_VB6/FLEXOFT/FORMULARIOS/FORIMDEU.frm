VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIMDEU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de DEUDA CLIENTES"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1470
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   210
      Width           =   2850
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1470
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   630
      Width           =   2850
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   4515
      TabIndex        =   8
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMDEU.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\DEUDA VDR.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   9030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Sheet1$"
      Top             =   105
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   11025
      Picture         =   "FORIMDEU.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIMDEU.frx":0454
      Height          =   6420
      Left            =   0
      OleObjectBlob   =   "FORIMDEU.frx":0464
      TabIndex        =   1
      Top             =   1365
      Width           =   10830
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aro Plastico"
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resorte"
      Height          =   255
      Index           =   3
      Left            =   7665
      TabIndex        =   6
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7665
      TabIndex        =   5
      Top             =   2100
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7665
      TabIndex        =   4
      Top             =   1785
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7665
      TabIndex        =   3
      Top             =   1470
      Width           =   1380
   End
End
Attribute VB_Name = "FORIMDEU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Close
   Hide
End Sub

Private Sub Command2_Click()
Data1.DatabaseName = Text1
Data1.RecordSource = Text2
Data1.Refresh
DBGrid1.Refresh
End Sub


Private Sub Command4_Click()
If COMALTER%("Colocar 'DEUDA VDR.XLS'\en Carpeta 'CLIENTES/VDR/TRAIDOS'\\Continuar\Cancelar") <> 1 Then
Exit Sub
End If
Screen.MousePointer = 11
   URECORD& = Data1.Recordset.Recordcount
   DBGrid1.Height = (URECORD& + 50) * DBGrid1.RowHeight
   '
   LU$ = LUDAT$
   '
   '
   For U& = 0 To URECORD&
     On Error Resume Next
     DBGrid1.Row = U&
     If Err Then GoTo 19
     On Error GoTo 0
     DBGrid1.Col = 0: TT1$ = DBGrid1.TEXT
     DBGrid1.Col = 1: TT2$ = DBGrid1.TEXT
     DBGrid1.Col = 2: TT3$ = DBGrid1.TEXT
     '
     If Left$(UCase$(TT1$), 8) = "CLIENTE:" Then
       TT1$ = TRIM$(Mid$(TT1$, 9))
       NCI% = Val(TT1$)
       GoTo 19
     End If
     '
     If TRIM$(TT2$) = "" Then GoTo 19
     '
     FEMIS = FECHANUM(TT1$): If FEMIS < 1 Then GoTo 19
     '
     VENCIM = FEMIS: While DIENTRE%(FEMIS, VENCIM) < 30: VENCIM = VENCIM + 1: Wend
     DOCULA$ = TT2$
     '
     TIX% = 0
     TIDOCU$ = Left$(DOCULA$, 2)
     If TIDOCU$ = "FC" Then TIX% = 2
     If TIDOCU$ = "ND" Then TIX% = 3
     If TIDOCU$ = "RC" Then TIX% = 4
     If TIDOCU$ = "NC" Then TIX% = 6
     If TIX% < 1 Then
        AAA = BBB
     End If
     '
     LETRA$ = TRIM$(Mid$(DOCULA$, 4, 2))
     VAX% = 1: If LETRA$ = "B" Then VAX% = 2
     If LETRA$ = "E" Then VAX% = 3
     '
     PPXX% = InStr(DOCULA$, "-")
     PREFI$ = TRIM$(Mid$(DOCULA$, PPXX% - 4, 4))
     NUFAX$ = TRIM$(Mid$(DOCULA$, PPXX% + 1, 8))
     Call REPLA(NUFAX$, "0", 1, 1)
     If TIX% <> 4 Then Call REPLA(NUFAX$, 9, 3, 1)
     NF& = Val(NUFAX$)
     '
     IMPO# = Val(REPLACAR(TT3$, ",", "."))
     '
     ABUS$ = "FACTUR"
     If TIX% = 4 Then ABUS$ = "COBRAN"
     '
     RFF$ = RECFILT$(ABUS$, 1, MKI$(NCI%))
     For UK& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, UK&, 4))
       XX$ = REGLEIDO$(ABUS$, REGO&)
       If CVS(Mid$(XX$, 3, 4)) = NF& Then GoTo 19
     Next UK&
     '
     YY$ = REGBLAN$("SVD202")
     Call REPLA(YY$, MKI$(NCI%), 1, 2)
     Call REPLA(YY$, Chr$(TIX%), 3, 1)
     Call REPLA(YY$, Chr$(VAX%), 4, 1)
     Call REPLA(YY$, MKS$(NF&), 5, 4)
     FVI% = VENCIM
     Call REPLA(YY$, MKI$(FVI%), 9, 2)
     Call REPLA(YY$, MKD$(Abs(IMPO#)), 11, 8)
     FEI% = FEMIS
     Call REPLA(YY$, MKI$(FEI%), 19, 2)
     Call REPLA(YY$, MKD$(Abs(IMPO#)), 21, 8)
     Call REPLA(YY$, MKI$(Val(PREFI$)), 29, 2)
     Call REGAPP("SVD202", YY$)
     '
     If TIX% <> 4 Then
         YY$ = REGBLAN$("FACTUR")
         Call REPLA(YY$, MKI$(NCI%), 1, 2)
         Call REPLA(YY$, MKS$(NF&), 3, 4)
         Call REPLA(YY$, MKI$(FEI%), 7, 2)
         Call REPLA(YY$, MKI$(FVI%), 9, 2)
         Call REPLA(YY$, Chr$(TIX%), 13, 1)
         Call REPLA(YY$, Chr$(VAX%), 14, 1)
         IMPONE# = IMPO#: IMPOIVA# = 0
         If VAX% = 1 Then
           IMPONE# = IMPO# / 1.21
           IMPOIVA# = IMPONE# * 0.21
         End If
         Call REPLA(YY$, MKD$(Abs(IMPONE#)), 15, 8)
         Call REPLA(YY$, MKD$(Abs(IMPOIVA#)), 23, 8)
         Call REPLA(YY$, MKD$(Abs(IMPO#)), 39, 8)
         Call REPLA(YY$, MKI$(Val(PREFI$)), 47, 2)
         Call REGAPP("FACTUR", YY$)
       ElseIf TIX% = 4 Then
         YY$ = REGBLAN$("COBRAN")
         Call REPLA(YY$, MKI$(NCI%), 1, 2)
         Call REPLA(YY$, MKS$(NF&), 3, 4)
         Call REPLA(YY$, MKI$(FEI%), 7, 2)
         Call REPLA(YY$, Chr$(TIX%), 9, 1)
         Call REPLA(YY$, Chr$(VAX%), 10, 1)
         Call REPLA(YY$, MKD$(Abs(IMPO#)), 15, 8)
         Call REGAPP("COBRAN", YY$)
     End If
19 Next U&
   '
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("COBRAN")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("*.*")
   '
Screen.MousePointer = 1

   Call COMUNI("La Importación de Composicion de Deuda Clientes\se ha Completado Exitosamente.")
   '
Exit Sub

End Sub

