VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FORIDEUP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de DEUDA PROVEEDORES"
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
      Picture         =   "FORIDEUP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Importación"
      Top             =   945
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "F:\CLIENTES\VDR\TRAIDOS\COMPOSICION PROVEEDORES SJ.xls"
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
      Picture         =   "FORIDEUP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORIDEUP.frx":0454
      Height          =   7470
      Left            =   0
      OleObjectBlob   =   "FORIDEUP.frx":0464
      TabIndex        =   1
      Top             =   315
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
      Left            =   7560
      TabIndex        =   6
      Top             =   1260
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carcaza"
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   5
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acelerante"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   4
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mezcla"
      Height          =   255
      Index           =   0
      Left            =   7560
      TabIndex        =   3
      Top             =   315
      Width           =   1380
   End
End
Attribute VB_Name = "FORIDEUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Close
   Hide
End Sub
Private Sub Command4_Click()
Dim SD#(10), IMC#(10)
If COMALTER%("Colocar 'CODEUPRO.XLS'\en Carpeta 'CLIENTES/VDR/TRAIDOS'\\Continuar\Cancelar") <> 1 Then
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
     DBGrid1.Col = 0: NCI% = Val(DBGrid1.TEXT)
     DBGrid1.Col = 1: FEMIS = FECHANUM(DBGrid1.TEXT)
     DBGrid1.Col = 3: VENCIM = FECHANUM(DBGrid1.TEXT)
     '
     DBGrid1.Col = 2: DOCULA$ = TRIM$(DBGrid1.TEXT)
     If DOCULA$ = "" Then GoTo 19
     '
     TIX% = 0
     TIDOCU$ = Left$(DOCULA$, 2)
     If TIDOCU$ = "FC" Then TIX% = 2
     If TIDOCU$ = "ND" Then TIX% = 3
     If TIDOCU$ = "RC" Or TIDOCU$ = "OP" Then TIX% = 4
     If TIDOCU$ = "NC" Then TIX% = 6
     If TIX% < 1 Then
        AAA = BBB
     End If
     '
     LETRA$ = TRIM$(Mid$(DOCULA$, 4, 2))
     VAX% = 1: If LETRA$ = "B" Then VAX% = 2
     If LETRA$ = "E" Then VAX% = 3
     If LETRA$ = "C" Then VAX% = 3
     '
     PPXX% = InStr(DOCULA$, "-")
     PREFI$ = TRIM$(Mid$(DOCULA$, PPXX% - 4, 4))
     NUFAX$ = TRIM$(Mid$(DOCULA$, PPXX% + 1, 8))
     NF& = Val(NUFAX$)
     '
     DBGrid1.Col = 4: IMPO# = (-1) * Val(DBGrid1.TEXT)
     '
     
     TOFAC#(1) = IMPO#
     TOFAC#(0) = (-1) * IMPO#
     FVII%(1) = VENCIM
     SD#(1) = IMPO#
     FAJUI% = FEMIS
     HOII% = HOY(HOS$)
     If NCI% > 0 Then Call CUECORR(FAJUI%, (-1) * NCI%, TIX%, VAX%, Val(PREFI$), NF&, RASOCLI$((-1) * NCI%), "$", TOFAC#(), FVII%(), SD#(), GRU%(), IMC#(), 1, HOII%)
19 Next U&
   '
   Call CIERRARCH("CUECORR")
   Call CIERRARCH("*.*")
   '
Screen.MousePointer = 1

   Call COMUNI("La Importación de Composicion de Deuda Proveedores\se ha Completado Exitosamente.")
   '
Exit Sub

End Sub

