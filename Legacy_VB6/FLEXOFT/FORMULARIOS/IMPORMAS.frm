VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form IMPORMAS 
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
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
      Height          =   4455
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Importación Definitiva"
      Top             =   3720
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
         Height          =   1140
         IntegralHeight  =   0   'False
         ItemData        =   "IMPORMAS.frx":0000
         Left            =   120
         List            =   "IMPORMAS.frx":0002
         TabIndex        =   34
         Top             =   3120
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
         Left            =   10650
         Picture         =   "IMPORMAS.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Listado por Impresora"
         Top             =   3480
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
         TabIndex        =   32
         Top             =   560
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
         Picture         =   "IMPORMAS.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1965
         Left            =   10320
         TabIndex        =   25
         Top             =   560
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   150
         Max             =   10000
         TabIndex        =   24
         Top             =   2480
         Width           =   10160
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   22
         Top             =   560
         Width           =   10215
         Begin VB.PictureBox Picture3 
            Height          =   1950
            Left            =   0
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   26
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
            Height          =   1700
            IntegralHeight  =   0   'False
            ItemData        =   "IMPORMAS.frx":0618
            Left            =   0
            List            =   "IMPORMAS.frx":061A
            TabIndex        =   23
            Top             =   250
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
            TabIndex        =   38
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
         TabIndex        =   35
         Top             =   2880
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPORMAS.frx":061C
         Height          =   3960
         Left            =   150
         OleObjectBlob   =   "IMPORMAS.frx":062C
         TabIndex        =   20
         Top             =   360
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   15
      Top             =   2400
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   16
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Formato de Importacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   4800
      TabIndex        =   13
      Top             =   240
      Width           =   6100
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   17
         Top             =   350
         Width           =   5655
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Col. Dato / Variable            Tipo     Rango"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   50
            TabIndex        =   18
            Top             =   0
            Width           =   5655
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Archivo a Importar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Categorías de Operarios Productivos"
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
         Index           =   4
         Left            =   360
         TabIndex        =   8
         Top             =   1900
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Tabla de Equipos de Producción"
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
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1610
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Rutas - Secuencias de Operaciones"
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
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1060
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Fórmulas y Estructuras de Producto"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   770
         Width           =   3735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Maestro de Items de Stock"
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
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4560
         Y1              =   1440
         Y2              =   1440
      End
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
      TabIndex        =   2
      Top             =   2640
      Width           =   11655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   36
         Top             =   480
         Width           =   6975
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
            Left            =   20
            TabIndex        =   37
            Top             =   20
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   10
         Text            =   " "
         Top             =   480
         Width           =   2895
      End
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
         Left            =   7920
         TabIndex        =   9
         Top             =   465
         Width           =   220
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
         Left            =   8160
         TabIndex        =   12
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
         TabIndex        =   11
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
      Height          =   750
      Left            =   11040
      Picture         =   "IMPORMAS.frx":0FDE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Inicia Pre-Importación"
      Top             =   1200
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
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2200
      Visible         =   0   'False
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
      Picture         =   "IMPORMAS.frx":12E8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORMAS"
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

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   Screen.MousePointer = 11
   Select Case Val(NARIM)
     Case 0
       Call IMPORMAS
     Case 1
       Call IMPORSTRU
     Case 2
       Call IMPORUTS
     Case 3
       Call IMPORMAQ
     Case 4
       Call IMPORCAT
   End Select
   Screen.MousePointer = 1
   Call COMUNI("Se ha Completado la Importacion Definitiva")
   Frame5.Visible = False
   DoEvents
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo Origen de Datos"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.Filter = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
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
   Select Case Val(NARIM)
     Case 0
       Call PREIMPORMAS
     Case 1
       Call PREIMPORSTRU
     Case 2
       Call PREIMPORUTS
     Case 3
       Call PREIMPORMAQ
     Case 4
       Call PREIMPORCAT
  End Select
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
   Option1(0).Value = True
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

Private Sub Option1_Click(Index As Integer)
   List1.Clear
   List2.Clear
   List4.Clear
   Frame5.Visible = False
   Label1 = ""
   Combo1.Clear
   DBGrid1.ClearFields
   NARIM = TRIM$(Str$(Index))
   Select Case Index
   Case 0
     List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
     List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
     List2.AddItem "  3  Tipo de Material           Num      0 - 255"
     List2.AddItem "  4  Unidad de Medida           String   <= 2 car"
     List2.AddItem "  5  Stock Minimo               Num      0 - 10^8"
     List2.AddItem "  6  Lote Economico             Num      0 - 10^8"
     List2.AddItem "  7  Costo Unitario             Num      0 - 10^8"
     List2.AddItem "  8  Moneda de Costeo           Num      1 - 255"
     List2.AddItem "  9  Fecha de Costeo            Date"
     List2.AddItem " 10  Numero Proveedor Habitual  Num      0 - 32767"
     Label8 = "Código           C-Int  Descripcion                                                         TM  UM        St.Min     Lote Ec  Costo Unit  MC  Fecha       Pr-H"
   Case 1
     List2.AddItem "  1  Codigo de Conjunto         String   <= 15 car"
     List2.AddItem "  2  Codigo del Componente      String   <= 15 car"
     List2.AddItem "  3  Uso / Cantidad Componente  Num      10^-5 - 10^6"
     List2.AddItem "  4  Plazo Disposicion          Num      1 - 255"
     Label8 = "Código Conj.      Código Elem.      Descripcion                         Uso/Cant  PD"
   Case 2
     List2.AddItem "  1  Codigo de Producto         String   <= 15 car"
     List2.AddItem "  2  Descripcion de Operacion   String   <= 48 car"
     List2.AddItem "  3  Codigo de Maquina          String   <= 6 car"
     List2.AddItem "  4  Horas Preparacion (APP)    Num      0 - 999,99"
     List2.AddItem "  5  Horas Fijo por Lote        Num      0 - 999,99"
     List2.AddItem "  6  Minutos Variable / Pza     Num      0 - 99,99999"
     List2.AddItem "  7  % Demora Distribuida       Num      0 - 99,9"
     List2.AddItem "  8  Lote Standard              Num      0 - 10^6"
     List2.AddItem "  9  Operarios                  Num      0,01 - 99,99"
     List2.AddItem " 10  Maquina Adicional 1        String   <= 6 car"
     List2.AddItem " 11  Maquina Adicional 2        String   <= 6 car"
     List2.AddItem " 12  Maquina Adicional 3        String   <= 6 car"
     List2.AddItem " 13  Programa / Hoja Proc.Fab.  String   <= 8 car"
     Label8 = "Código          N-Op  Descripcion de Operacion                          Maquina   Hs.APP    Hs.Fix  Min.Variab  Lote Std    % DD   Opers  MAux-1  MAux-2  MAux-3  No.HPF"
   Case 3
     List2.AddItem "  1  Codigo de Maquina          String   <= 6 car"
     List2.AddItem "  2  Descripcion                String   <= 30 car"
     List2.AddItem "  3  Cuota Horaria $ / hora     Num      0 - 10^6"
     List2.AddItem "  4  Codigo de Seccion          Num      100 - 999"
     List2.AddItem "  5  Maquina Equivalentea       String   <= 6 car"
     List2.AddItem "  6  Coeficiente (%)            Num      0 - 999,99"
     List2.AddItem "  7  Observaciones              String   <= 30 car"
     Label8 = "Código  Descripcion                         $ / Hora   Sección   Alter.        Coef.  Observaciones"
   Case 4
     List2.AddItem "  1  Descripcion Categoria      String   <= 27 car"
     List2.AddItem "  2  Costo M.O.D. / hora        Num      0 - 10^6"
     List2.AddItem "  3  Observaciones              String   <= 32 car"
     Label8 = " N-O  Descripcion                       $ / Hora  Observaciones"
   End Select
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

Sub PREIMPORMAS()
   '
   If DBGrid1.Columns.Count < 10 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     CDXI% = CODINT%(CODEX$)
     If CDXI% > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     VALO$ = AJUSTD$(Str$(CDXI%), 6)
     Call REPLA(TTXX$, VALO$, 17, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 25, 64)
     '
     DBGrid1.Col = 2                     ' TIPO DE MATERIAL
     VALO$ = AJUSTD$(DBGrid1.TEXT, 4)
     If Val(VALO$) < 0 Or Val(VALO$) > 255 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 91, 4)
     '
     DBGrid1.Col = 3                     ' UNIDAD DE MEDIDA
     UME$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, UME$, 97, 4)
     '
     DBGrid1.Col = 4                     ' STOCK MINIMO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     If Val(VALO$) < 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 101, 12)
     '
     DBGrid1.Col = 5                     ' LOTE DE REPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     If Val(VALO$) < 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 113, 12)
     '
     DBGrid1.Col = 6                     ' COSTO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 125, 12)
     '
     DBGrid1.Col = 7                     ' MONEDA DE COSTEO
     VALO$ = AJUSTD$(DBGrid1.TEXT, 4)
      If Val(VALO$) < 0 Or Val(VALO$) > 255 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 137, 4)
     '
     DBGrid1.Col = 8                    ' FECHA DE COSTEO
     VALO$ = ""
     On Error Resume Next
     VALO$ = Format$(DBGrid1.TEXT, "dd/mm/yy")
     On Error GoTo 0
     Call REPLA(TTXX$, VALO$, 143, 8)
     '
     DBGrid1.Col = 9                     ' PROVEEDOR HABITUAL
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 32766 Then VALOR = 0
     VALO$ = FORMATNUM$(VALOR, "F8.0")
     Call REPLA(TTXX$, VALO$, 151, 8)
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
19 Next U&

20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORSTRU()
   '
   If DBGrid1.Columns.Count < 4 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT): If CODEX$ = "" Then CODEX$ = CODEXA$
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Conjunto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     Call LEEEXPLO(CODEX$, 1)
     If CAIT% > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3
     If CAIT% > 0 Then
         ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And CAIT% > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     '
     DBGrid1.Col = 1
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or COELE$ = CODEX$ Or CODINT%(COELE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Elemento Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, COELE$, 19, 16)
     Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 37, 32)
     '
     DBGrid1.Col = 2                     ' USO UNITARIO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 69, 12)
     '
     DBGrid1.Col = 3                     ' PLAZO DISPOSICION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "I4")
     Call REPLA(TTXX$, VALO$, 81, 4)
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
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORUTS()
   '
   If DBGrid1.Columns.Count < 12 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0: CODEXA$ = "": CODEYA$ = ""
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 15 Or CODINT%(CODEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If CODEX$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = CODEX$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(CODEX$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, CODEX$, 1, 16)
     If CODEX$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = CODEX$
     '
     DBGrid1.Col = 1
     DEOPE$ = TRIM$(DBGrid1.TEXT)
     If DEOPE$ = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion de la Operacion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     NOPE% = NOPE% + 10
     VALO$ = FORMATNUM$(NOPE%, "I4")
     Call REPLA(TTXX$, VALO$, 17, 4)
     Call REPLA(TTXX$, DEOPE$, 23, 48)
     '
     DBGrid1.Col = 2                     ' MAQUINA
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina '" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 73, 6)
     '
     DBGrid1.Col = 3                     ' HORAS PREPARACION
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 79, 10)
     '
     DBGrid1.Col = 4                     ' HORAS STANDARD FIJO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 89, 10)
     '
     DBGrid1.Col = 5                     ' MINUTOS STANDARD VARIABLE
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F12.5")
     Call REPLA(TTXX$, VALO$, 99, 12)
     '
     DBGrid1.Col = 6                     ' LOTE STANDARD UNIDADES
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 1 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F10.1")
     Call REPLA(TTXX$, VALO$, 111, 10)
     '
     DBGrid1.Col = 7                     ' % DEM.DISTRIBUIDA
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     If VALOR < 0 Then VALOR = 0
     On Error GoTo 0
     VALO$ = FORMATNUM$(VALOR, "F8.1")
     Call REPLA(TTXX$, VALO$, 121, 8)
     '
     DBGrid1.Col = 8                     ' OPERARIOS
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0.01 Then VALOR = 1
     VALO$ = FORMATNUM$(VALOR, "F8.2")
     Call REPLA(TTXX$, VALO$, 129, 8)
     '
     DBGrid1.Col = 9                     ' MAQUINA AUX 1
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.1'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 139, 6)
     '
     DBGrid1.Col = 10                     ' MAQUINA AUX 2
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.2'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 147, 6)
     '
     DBGrid1.Col = 11                     ' MAQUINA AUX 3
     CMAQUI$ = TRIM$(DBGrid1.TEXT)
     If CMAQUI$ <> "" Then
       If ECOARCH$("PADMAQ", CMAQUI$) = "" Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Maquina Aux.3'" + CMAQUI$ + "' Inexistente o No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CMAQUI$, 155, 6)
     '
     DBGrid1.Col = 12                     ' H.P.FABR.
     HOPROFA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, HOPROFA$, 163, 8)
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
   Screen.MousePointer = 1
   '
20 End Sub

Sub PREIMPORMAQ()
   '
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     CODEX$ = TRIM$(DBGrid1.TEXT)
     If CODEX$ = "" Or Len(CODEX$) < 3 Or Len(CODEX$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, CODEX$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + CODEX$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, CODEX$, 1, 6)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 9, 30)
     '
     DBGrid1.Col = 2                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Or VALOR > 999999.99 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.4")
     Call REPLA(TTXX$, VALO$, 41, 12)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(DBGrid1.TEXT)
     If VTX$ <> "" And Val(VTX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion debe ser numerico 'nnn.nnn'"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(VTX$)
     On Error GoTo 0
     If VALOR > 999.999 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Seccion no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F10.3")
     Call REPLA(TTXX$, VALO$, 53, 10)
     '
     DBGrid1.Col = 4
     CODEQ$ = TRIM$(DBGrid1.TEXT)
     If CODEQ$ <> "" Then
       If Len(CODEQ$) < 3 Or Len(CODEQ$) > 6 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Equivalente No Valido"
         List1.AddItem MERROX$
         GoTo 19
       End If
     End If
     Call REPLA(TTXX$, CODEQ$, 66, 6)
     '
     DBGrid1.Col = 5                     ' Cuota Horaria
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     VAMIN = 0: If CODEQ$ <> "" Then VAMIN = 10
     If CODEQ$ = "" Then VALOR = 0
     If VALOR < VAMIN Or VALOR > 999.9 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Coeficiente Equivalencia fuera de Rango"
       List1.AddItem MERROX$
       GoTo 19
     End If
     VALO$ = FORMATNUM$(VALOR, "F12.1")
     Call REPLA(TTXX$, VALO$, 72, 12)
     '
     DBGrid1.Col = 6
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
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
   
20 End Sub

Sub PREIMPORCAT()
   '
   If DBGrid1.Columns.Count < 3 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   If URECORD& > 255 Then
     Call MENSERR(24, "Imposible Importar - Máximo 255 Registros.")
     GoTo 20
   End If
   '
   If ULTREG&("CATOPER") > 0 Then
     If COMALTER%("Tabla de Categorias Existente\\Conservar\Sobre-Escribir") <> 2 Then
       GoTo 20
     End If
   End If
   '
   JJ& = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     '
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     DECAT$ = TRIM$(DBGrid1.TEXT)
     If DECAT$ <> "" Then
       JJ& = JJ& + 1
       COCAT$ = AJUSTD$(Str$(JJ&), 4)
       Call REPLA(TTXX$, COCAT$, 1, 4)
       Call REPLA(TTXX$, DECAT$, 7, 30)
       '
       DBGrid1.Col = 1
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       If VALOR < 0 Or VALOR > 999999.99 Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Cuota Horaria fuera de Rango"
         List1.AddItem MERROX$
         GoTo 19
       End If
       VALO$ = FORMATNUM$(VALOR, "F12.4")
       Call REPLA(TTXX$, VALO$, 37, 12)
       '
       DBGrid1.Col = 2
       OBSER$ = TRIM$(DBGrid1.TEXT)
       Call REPLA(TTXX$, OBSER$, 51, 32)
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
     End If
19 Next U&
   
20 End Sub

Sub IMPORMAS()
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = Val(Mid$(TTXX$, 17, 6))
     DESIT$ = TRIM$(Mid$(TTXX$, 25, 64))
     TIMAT% = Val(Mid$(TTXX$, 91, 4)) Mod 256
     UNIME$ = TRIM$(Mid$(TTXX$, 97, 4))
     STOMI = Val(Mid$(TTXX$, 101, 12))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = Val(Mid$(TTXX$, 113, 12))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     COSUN = Val(Mid$(TTXX$, 125, 12))
     MONEC% = Val(Mid$(TTXX$, 137, 4)) Mod 256
     FECOU% = FECHANUM(Mid$(TTXX$, 143, 8))
       PRH = Abs(Val(Mid$(TTXX$, 151, 8)))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     REMAS$ = REBLA$
     If CDXI% > 0 Then
       REMAL& = CDXI% + 1
       REMAS$ = REGLEIDO$("MASTER", REMAL&)
     End If
     '
     If DESIT$ <> "" Then Call REPLA(REMAS$, DESIT$, 5, 64)
     If CODIT$ <> "" Then Call REPLA(REMAS$, CODIT$, 77, 16)
     If TIMAT% > 0 Then Call REPLA(REMAS$, Chr$(TIMAT%), 93, 1)
     If Asc(Mid$(REMAS$, 94, 1)) < 1 Then Call REPLA(REMAS$, Chr$(2), 94, 1) ' POR MINIMOS
     If MONEC% > 0 Then Call REPLA(REMAS$, Chr$(MONEC%), 95, 1)
     If Abs(LOMAX) > 0 Then Call REPLA(REMAS$, MKI$(Int(LOMAX)), 103, 2)
     If Abs(STOMI) > 0 Then Call REPLA(REMAS$, MKI$(Int(STOMI)), 105, 2)
     If Abs(LECON) > 0 Then Call REPLA(REMAS$, MKI$(Int(LECON)), 107, 2)
     If UNIME$ <> "" Then Call REPLA(REMAS$, UNIME$, 109, 2)
     If Abs(COSUN) > 0.00005 Then Call REPLA(REMAS$, MKS$(COSUN), 111, 4)
     If FECOU% > 0 Then Call REPLA(REMAS$, MKI$(FECOU%), 115, 2)
     If PRHAB% > 0 Then Call REPLA(REMAS$, MKI$(PRHAB%), 123, 2)
     '
     CIXI% = CVI(Left$(REMAS$, 2))
     If CIXI% < 1 Then
       NMS% = NMS% + 1
       CIXI% = NMS%
       Call REPLA(REMAS$, MKI$(CIXI%), 1, 2)
     End If
     REMAL& = CIXI% + 1
     Call GRAREG("MASTER", REMAS$, REMAL&)
     '
   Next U&
   '
   REMAS$ = REGLEIDO$("MASTER", 1)
   Call REPLA(REMAS$, MKI$(NMS%), 1, 2)
   Call GRAREG("MASTER", REMAS$, 1)
   '
   UREMA& = NMS% + 1
   Call SETULTREG("MASTER", UREMA&)
   Call CIERRARCH("MASTER")
   '
   Call BLOQARCH("MASTER")
   Call AMASTO01.REVIMASTER
   Call REPUNCODIG
   Call CIERRARCH("MASTER")
   '
End Sub
'

Sub IMPORSTRU()
   '
   Dim PUNSTRU$(32767)
   '
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   ATENTIA% = ATENTI%: ATENTI% = 12345 ' PA QUE NO REORGANICE IMPLOSI
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       CAIT% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = Space$(64)
         Call REPLA(TTYY$, Mid$(TTXX$, 19, 16), 1, 16)
         Call REPLA(TTYY$, Mid$(TTXX$, 69, 12), 49, 12)
         Call REPLA(TTYY$, Mid$(TTXX$, 81, 4), 61, 4)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("ESTRULAR")
   ATENTI% = ATENTIA%
   '
End Sub
'
Sub IMPORUTS()
   '
   Dim PUNSTRU$(32767)
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   NMS% = NUMAS%
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COCON$ = TRIM$(Left$(TTXX$, 16))
     CDXI% = CODINT%(COCON$)
     '
     If CDXI% > 0 Then
       If CDXI% <= NMS% Then
         PUNSTRU$(CDXI%) = PUNSTRU$(CDXI%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   REBLA$ = REGBLAN$("SECOPER")
   For CDXI% = 1 To NMS%
     If Len(PUNSTRU$(CDXI%)) > 0 Then
       RFF$ = RECFILT$("SECOPER", 1, MKI$(CDXI%))
       '
       For UKU& = 1 To Len(RFF$) Step 4
         RESE& = CVS(Mid$(RFF$, UKU&, 4))
         Call GRAREG("SECOPER", String$(128, 0), RESE&)
       Next UKU&
       '
       NOPEI% = 0
       For UKU& = 1 To Len(PUNSTRU$(CDXI%)) Step 4
         U& = CVS(Mid$(PUNSTRU$(CDXI%), UKU&, 4))
         TTXX$ = List4.List(U& - 1)
         TTYY$ = REBLA$
         Call REPLA(TTYY$, MKI$(CODINT%(Left$(TTXX$, 16))), 1, 2)
           NOPEI% = NOPEI% + 10
         Call REPLA(TTYY$, MKI$(NOPEI%), 3, 2)
         Call REPLA(TTYY$, Mid$(TTXX$, 23, 48), 5, 48)
         Call REPLA(TTYY$, Mid$(TTXX$, 73, 6), 53, 6)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 79, 10))), 59, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 89, 10))), 63, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 99, 12))), 67, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 111, 10))), 91, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 121, 8))), 75, 4)
         Call REPLA(TTYY$, MKS$(Val(Mid$(TTXX$, 129, 8))), 83, 4)
         Call REPLA(TTYY$, Mid$(TTXX$, 139, 6), 103, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 147, 6), 109, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 155, 6), 115, 6)
         Call REPLA(TTYY$, Mid$(TTXX$, 163), 121, 6)
         Call REGAPP("SECOPER", TTYY$)
       Next UKU&
     End If
   Next CDXI%
   '
   Call CIERRARCH("SECOPER")
   Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAQ()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("PADMAQ")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     COMAQ$ = TRIM$(Left$(TTXX$, 6))
     DEMAQ$ = TRIM$(Mid$(TTXX$, 9, 30))
     CUHOR = Val(Mid$(TTXX$, 41, 12))
     SECCI = Val(Mid$(TTXX$, 53, 10))
     MAQEQ$ = TRIM$(Mid$(TTXX$, 66, 6))
     COEFI% = Val(Mid$(TTXX$, 72, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 87, 30))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, COMAQ$, 1, 6)
     Call REPLA(REMAQ$, DEMAQ$, 7, 30)
     Call REPLA(REMAQ$, MKS$(SECCI), 37, 4)
     Call REPLA(REMAQ$, MAQEQ$, 43, 6)
     Call REPLA(REMAQ$, MKI$(COEFI), 49, 2)
     Call REPLA(REMAQ$, OBSER$, 51, 30)
     Call REPLA(REMAQ$, MKS$(CUHOR), 93, 4)
     Call FILTERSETRECORD("PADMAQ", 1, COMAQ$, REMAQ$) ' sobreescribe SI IGUAL CODIGO
     '
   Next U&
   Call CIERRARCH("PADMAQ")
   '
End Sub

Sub IMPORCAT()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("CATOPER")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     ICAT% = Val(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = Val(Mid$(TTXX$, 37, 12))
     OBSER$ = TRIM$(Mid$(TTXX$, 51, 32))
     '
     REMAQ$ = REBLA$
     Call REPLA(REMAQ$, Chr$(ICAT%), 1, 1)
     Call REPLA(REMAQ$, DECAT$, 2, 27)
     Call REPLA(REMAQ$, MKS$(CUHOR), 29, 4)
     Call REPLA(REMAQ$, OBSER$, 33, 32)
     JJ& = JJ& + 1
     Call GRAREG("CATOPER", REMAQ$, JJ&)
   Next U&
   '
   Call SETULTREG("CATOPER", JJ&)
   Call CIERRARCH("CATOPER")
   '
End Sub

