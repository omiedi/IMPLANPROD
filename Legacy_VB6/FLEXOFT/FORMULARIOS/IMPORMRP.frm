VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form IMPORMRP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Resultado de Pre-Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6845
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Importación Definitiva"
      Top             =   1320
      Visible         =   0   'False
      Width           =   11625
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         IntegralHeight  =   0   'False
         ItemData        =   "IMPORMRP.frx":0000
         Left            =   120
         List            =   "IMPORMRP.frx":0002
         TabIndex        =   21
         Top             =   4680
         Width           =   10400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "List"
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
         Left            =   10680
         Picture         =   "IMPORMRP.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Listado por Impresora"
         Top             =   5880
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10680
         TabIndex        =   19
         Top             =   500
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Height          =   735
         Left            =   10680
         Picture         =   "IMPORMRP.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   3250
         Left            =   10320
         TabIndex        =   12
         Top             =   750
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   150
         Max             =   10000
         TabIndex        =   11
         Top             =   4000
         Width           =   10190
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   140
         ScaleHeight     =   3525
         ScaleWidth      =   10185
         TabIndex        =   9
         Top             =   500
         Width           =   10180
         Begin VB.PictureBox Picture3 
            Height          =   3735
            Left            =   0
            ScaleHeight     =   3675
            ScaleWidth      =   0
            TabIndex        =   13
            Top             =   0
            Width           =   60
         End
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3270
            IntegralHeight  =   0   'False
            ItemData        =   "IMPORMRP.frx":0618
            Left            =   0
            List            =   "IMPORMRP.frx":061A
            TabIndex        =   10
            Top             =   240
            Width           =   30000
         End
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   40
            TabIndex        =   25
            Top             =   0
            Width           =   30000
         End
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Registros no Procesables"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label6 
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
         Left            =   10680
         TabIndex        =   17
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10620
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label TEPRUE 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3720
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Planilla Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPORMRP.frx":061C
         Height          =   6360
         Left            =   150
         OleObjectBlob   =   "IMPORMRP.frx":062C
         TabIndex        =   7
         Top             =   360
         Width           =   11295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Origen de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   10815
      Begin VB.CommandButton Command22 
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
         Height          =   315
         Left            =   7200
         TabIndex        =   2
         Top             =   465
         Width           =   220
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6255
         TabIndex        =   23
         Top             =   480
         Width           =   6315
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   15
            TabIndex        =   24
            Top             =   15
            Width           =   6195
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   7800
         TabIndex        =   3
         Text            =   " "
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
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
      Height          =   855
      Left            =   11040
      Picture         =   "IMPORMRP.frx":0FDE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   340
      Width           =   750
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Refresh
   Data1.Recordset.MoveFirst
   '
End Sub

Private Sub Combo1_Click()
   '
   Data1.RecordSource = Combo1.TEXT
   Data1.Recordset.MoveFirst
   Data1.Refresh
   '
End Sub


Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   '
   APROGMRP.List1.Clear
   APROGMRP.List2.Clear
   For U& = 1 To List4.ListCount
     TTXX$ = List4.List(U& - 1)
     While Len(TTXX$) < 64 + 225: TTXX$ = TTXX$ + " ": Wend
     APROGMRP.List1.AddItem TTXX$
     APROGMRP.List2.AddItem Mid$(TTXX$, 65)
   Next U&
   '
   MXX% = APROGMRP.List1.ListCount - 16
   If MXX% < 0 Then MXX% = 0
   On Error Resume Next
   APROGMRP.VScroll2.Value = 0
   APROGMRP.VScroll2.Max = MXX%
   APROGMRP.List1.TopIndex = 0
   On Error GoTo 0
   '
   Screen.MousePointer = 1
   Call COMUNI("Se ha Completado la Importacion Definitiva")
   Frame5.Visible = False
   DoEvents
   Command2.Enabled = True
   Hide
   '
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   Label1 = UCase$(CommonDialog1.filename)
   Data1.DatabaseName = Label1
   Combo1.Clear
   Combo1.TEXT = ""
   On Error Resume Next
   Data1.Refresh
   '
   CATAB% = Data1.Database.TableDefs.Count
   For TTB% = 1 To CATAB%
     Combo1.AddItem Data1.Database.TableDefs(TTB% - 1).Name
     If TTB% = 1 Then
       Combo1.TEXT = Data1.Database.TableDefs(TTB% - 1).Name
     End If
   Next TTB%
   On Error GoTo 0
   '
99 Screen.MousePointer = 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Printer.Print "Registros no Procesables"
     Printer.Print " "
     For U& = 1 To List1.ListCount
       Printer.Print List1.List(U& - 1)
     Next U&
     Printer.EndDoc
     Screen.MousePointer = 1
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   Frame5.Visible = True
   List1.Clear
   List4.Clear
   On Error Resume Next
   HScroll1.Value = 0
   VScroll2.Value = 0
   On Error GoTo 0
   '
   If EXISTE%(Label1) < 1 Then
     Call MENSERR(24, "Establezca Archivo Origen de Importación")
     Frame5.Visible = False
     GoTo 99
   End If
   '
   TEPRUE = ""
   On Error Resume Next
   URECORD& = Data1.Recordset.RecordCount
   If Err Then
     Call MENSERR(24, "Error de Acceso a Archivo\Origen de Importación")
     GoTo 99
   End If
   DBHANT = DBGrid1.Height
   DBGrid1.Height = (URECORD& + 10) * DBGrid1.RowHeight
   Data1.Recordset.MoveFirst
   '
   Screen.MousePointer = 11
   Call PREIMPORPLAN
   '
20 DBGrid1.Height = DBHANT
   Command4.Enabled = True
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (TEPRUE.Width - HScroll1.Width) / 10
   End If
   '
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Frame5.Visible = False
   DoEvents
End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Label8 = AJUSTI$("Código", 16)
   Label8 = Label8 + AJUSTI$("Descripción", 48)
   Call APROGMRP.PREMESES
   Label8 = Label8 + APROGMRP.Label5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  Label8.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Sub PREIMPORPLAN()
   '
   If DBGrid1.Columns.Count < 2 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(64 + 225)
     DBGrid1.Row = U& - 1
     '
     DBGrid1.Col = 0
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo '" + COELE$ + "' Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
12   Call REPLA(TTXX$, COELE$, 1, 16)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 17, 48)
     '
     UCOL% = 26
     If DBGrid1.Columns.Count < UCOL% Then
       UCOL% = DBGrid1.Columns.Count
     End If
     For COLUX% = 2 To UCOL%
       DBGrid1.Col = COLUX% - 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       VALO$ = FORMATNUM$(VALOR, "F9.1")
       Call REPLA(TTXX$, VALO$, 65 + (COLUX% - 2) * 9, 9)
     Next COLUX%
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(TRIM$(TTXX$) + "*") > Len(TEPRUE) Then
       TEPRUE = TRIM$(TTXX$) + "*"
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   On Error Resume Next
   List4.TopIndex = 0
   On Error GoTo 0
   Screen.MousePointer = 1
   '
20 End Sub

