VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form IMPORMAIL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos de Importacion de Datos"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Lista de Pertenencias"
      ForeColor       =   &H00000000&
      Height          =   1640
      Left            =   120
      TabIndex        =   42
      Top             =   1200
      Width           =   4575
      Begin VB.ListBox List5 
         Height          =   1185
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   300
         Width           =   3885
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Resultado de Pre-Importación"
      Height          =   4245
      Left            =   105
      TabIndex        =   16
      ToolTipText     =   "Importación Definitiva"
      Top             =   3990
      Visible         =   0   'False
      Width           =   11655
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   31
         Top             =   420
         Width           =   10215
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1965
            Left            =   -20
            ScaleHeight     =   1965
            ScaleWidth      =   10215
            TabIndex        =   33
            Top             =   0
            Width           =   10215
            Begin VB.PictureBox Picture3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1950
               Left            =   -20
               ScaleHeight     =   1890
               ScaleWidth      =   0
               TabIndex        =   38
               Top             =   0
               Width           =   60
            End
            Begin VB.PictureBox Picture7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               ScaleHeight     =   210
               ScaleWidth      =   1.1994e5
               TabIndex        =   35
               Top             =   0
               Width           =   1.2e5
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
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
                  Left            =   40
                  TabIndex        =   37
                  Top             =   0
                  Width           =   120
               End
               Begin VB.Label Label82 
                  AutoSize        =   -1  'True
                  BackColor       =   &H008080FF&
                  BackStyle       =   0  'Transparent
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
                  Left            =   105
                  TabIndex        =   36
                  Top             =   0
                  Width           =   120
               End
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
               ItemData        =   "IMPORMAIL.frx":0000
               Left            =   0
               List            =   "IMPORMAIL.frx":0002
               TabIndex        =   34
               Top             =   250
               Width           =   1.3e5
            End
         End
         Begin VB.PictureBox Picture10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   32
            Top             =   0
            Width           =   60
         End
      End
      Begin VB.ListBox List1 
         Height          =   1140
         IntegralHeight  =   0   'False
         ItemData        =   "IMPORMAIL.frx":0004
         Left            =   120
         List            =   "IMPORMAIL.frx":0006
         TabIndex        =   26
         Top             =   2910
         Width           =   10400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "List"
         Height          =   750
         Left            =   10650
         Picture         =   "IMPORMAIL.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Listado por Impresora"
         Top             =   3255
         Width           =   750
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Hide"
         Height          =   300
         Left            =   10680
         TabIndex        =   24
         Top             =   560
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         Height          =   735
         Left            =   10680
         Picture         =   "IMPORMAIL.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1725
         Left            =   10360
         TabIndex        =   18
         Top             =   670
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   17
         Top             =   2340
         Width           =   10250
      End
      Begin VB.Line Line6 
         X1              =   135
         X2              =   135
         Y1              =   420
         Y2              =   2520
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
         TabIndex        =   27
         Top             =   2670
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   10680
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
      Height          =   4245
      Left            =   120
      TabIndex        =   14
      Top             =   3990
      Width           =   11625
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "IMPORMAIL.frx":061C
         Height          =   3750
         Left            =   150
         OleObjectBlob   =   "IMPORMAIL.frx":062C
         TabIndex        =   15
         Top             =   360
         Width           =   11295
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   11040
      ScaleHeight     =   75
      ScaleWidth      =   675
      TabIndex        =   10
      Top             =   2730
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         TabIndex        =   11
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Formato de Importacion"
      Height          =   2730
      Left            =   4800
      TabIndex        =   8
      Top             =   120
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
         Height          =   1980
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   5595
         TabIndex        =   12
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
            TabIndex        =   13
            Top             =   0
            Width           =   5655
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Origen de Datos"
      Height          =   975
      Left            =   105
      TabIndex        =   2
      Top             =   2940
      Width           =   11655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   28
         Top             =   480
         Width           =   6975
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   20
            TabIndex        =   29
            Top             =   20
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   8520
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "O.K."
      Height          =   750
      Left            =   11040
      Picture         =   "IMPORMAIL.frx":0FFF
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   " "
      Top             =   2520
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   750
      Left            =   11025
      Picture         =   "IMPORMAIL.frx":1309
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   4620
      TabIndex        =   30
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Archivo a Importar"
      Height          =   1050
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFE0&
         Caption         =   "Laboratorios Farmaceuticos"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   41
         Top             =   300
         Width           =   3735
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   900
         TabIndex        =   39
         Text            =   " "
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFE0&
         Caption         =   "Origen:"
         Height          =   345
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   705
      End
   End
   Begin VB.Label NARIM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CANCELIMP%
'
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
   CANCELIMP% = 0
   Call IMPORMAS
   Screen.MousePointer = 1
   If CANCELIMP% = 0 Then
     Call COMUNI("Se ha Completado la Importacion Definitiva")
   End If
   CANCELIMP% = 0
   Frame5.Visible = False
   DoEvents
   Command2.Enabled = True
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
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   Call PREIMPORMAS
   Call COMUNI("Se Ha Completado la Pre-Importación.")
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
   Call cargalist5
   '
'   Picture5.Top = 0
'   VScroll1.Min = 0
'   VScroll1.Max = (Picture5.Height - Picture6.Height) / 10
   '
End Sub

Sub cargalist5()
     FIN& = ULTREG&("LISLISTA")
     For U& = 1 To ULTREG&("LISLISTA")
       XX$ = REGLEIDO$("LISLISTA", U&)
       DESCRIPCI$ = TRIM$(Mid$(XX$, 3, 64))
       List5.AddItem DESCRIPCI$
     Next U&
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value ' + 40
  Picture2.Left = -20
  Picture2.Width = HScroll1.Width
  Ab1 = HScroll1.Value
  Ab1 = 10 * Ab1 + 40
  Ab2 = Abs(Picture7.Left)
  If Ab1 > Ab2 + 80 Then
    Picture2.Left = (-1) * (Ab1 - Ab2)
    Picture2.Width = HScroll1.Width - Picture2.Left
  End If
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
     List2.AddItem "  1  Laboratorio                String   <= 48 car"
     List2.AddItem "  2  Contacto                   String   <= 48 car"
     List2.AddItem "  3  Direccion                  String   <= 64 car"
     List2.AddItem "  4  Telefono                   String   <= 48 car"
     List2.AddItem "  5  E-mail                     String   <= 64 car"
     Label8 = "Laboratorio                                     Contacto                                        Direccion                                                       Telefono                                        E-mail"
   End Select
End Sub

Private Sub VScroll1_Change()
   Picture5.Top = (-10) * VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Picture5.Top = (-10) * VScroll1.Value
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
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(512)
        EMPRA$ = "": CONTA$ = "": DIREL$ = "": TELEF$ = "": MAILL$ = ""
        On Error Resume Next
        EMPRA$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        CONTA$ = .Recordset.Fields(1)
        DIREL$ = .Recordset.Fields(2)   ' TRIM$(DBGrid1.TEXT)
        TELEF$ = .Recordset.Fields(3)   ' TRIM$(DBGrid1.TEXT)
        MAILL$ = .Recordset.Fields(4)
12      Call REPLA(TTXX$, EMPRA$, 1, 48)
        Call REPLA(TTXX$, CONTA$, 49, 48)
        Call REPLA(TTXX$, DIREL$, 97, 64)
        Call REPLA(TTXX$, TELEF$, 161, 48)
        Call REPLA(TTXX$, MAILL$, 209, 64)
        On Error GoTo 0
        '
        If TRIM$(TTXX$) <> "" Then
           List4.AddItem TTXX$
           Label6 = TRIM$(Str$(List4.ListCount))
           If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
             TEPRUE = RTrim(TTXX$ + "*")
           End If
           On Error Resume Next
           List4.TopIndex = List4.ListCount - 7
           On Error GoTo 0
           DoEvents
        End If
        '
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub
'
Sub IMPORMAS()
   '
   HOII% = HOY(HOS$)
   PERLIST$ = ""
   For KKI% = 1 To List5.ListCount
      If List5.Selected(KKI% - 1) = True Then PERLIST$ = PERLIST$ + Chr$(KKI%)
   Next KKI%
   PERLIST$ = PERLIST$ + String$(16, 0)
   '
   FIN& = List4.ListCount
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     A$ = List4.List(U& - 1)
     APENO$ = TRIM$(Mid$(A$, 49, 48))
     DIREL$ = TRIM$(Mid$(A$, 209, 64))
     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
     EMPRA$ = TRIM$(Left$(A$, 48))
     ORIGE$ = Text18
     OBSER$ = "Domic: " + TRIM$(Mid$(A$, 97, 64)) + vbCrLf$ + "Telef: " + TRIM$(Mid$(A$, 161, 48))
     GoSub ALTANUE
   Next U&
   Call CIERRARCH("LIDIRMAI")
  '
Exit Sub
'
ALTANUE:
   DIREL$ = LCase$(TRIM$(DIREL$))
   PPXX% = InStr(DIREL$, " ")
   RESTO$ = ""
   If PPXX% > 0 Then
      DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
      RESTO$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
   End If
   PPXX% = InStr(DIREL$, "/")
   If PPXX% > 0 Then DIREL$ = TRIM$(Left$(DIREL$, PPXX%))
   If Left$(DIREL$, 1) = Chr$(39) Then
     DIREL$ = TRIM$(Mid$(DIREL$, 2))
   End If
   '
   If TRIM$(DIREL$) <> "" Then
     If InStr(DIREL$, "@") > 0 Then
       '
       POARRO% = InStr(DIREL$, "@")
       If InStr(POARRO% + 1, DIREL$, "@") > 0 Then
          List1.AddItem DIREL$
          GoTo 89
          'PARA EVITAR CARGAR DIRECCIONES MALAS
       End If
       '
       For KKU& = 1 To List1.ListCount
         If List1.List(KKU& - 1) = DIREL$ Then
            List1.AddItem DIREL$
            GoTo 89
         End If
       Next KKU&
       '
       If InStr(UCase$(APENO$), "MAILER-DAEMON") > 0 Then GoTo 89
       If InStr(UCase$(APENO$), "UNDISCLOSED-RECIPIEN") > 0 Then GoTo 89
       '
       RENUE$ = REGBLAN$("LIDIRMAI")
       '
       PPXX% = InStr(APENO$, "@")
       If PPXX% > 0 Then APENO$ = TRIM$(Left$(APENO$, PPXX% - 1))
       If APENO$ <> "" Then
         APENO$ = LCase$(APENO$)
         APENO$ = REPLACAR$(APENO$, "_", " ")
         '
         Call REPLA(APENO$, UCase$(Left$(APENO$, 1)), 1, 1)
         For KKI% = 2 To Len(APENO$)
           If Asc(Mid$(APENO$, KKI% - 1, 1)) < 48 Then
             Call REPLA(APENO$, UCase$(Mid$(APENO$, KKI%, 1)), KKI%, 1)
           End If
         Next KKI%
       End If
       '
       JJ& = 1 + ULTREG("LIDIRMAI")
       Call REPLA(RENUE$, MKS$(JJ&), 1, 4)
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       '
       Call REPLA(RENUE$, TRIM$(MUESTRA$), 5, 60)
       Call REPLA(RENUE$, APENO$, 65, 40)
       Call REPLA(RENUE$, DIREL$, 105, 60)
       Call REPLA(RENUE$, APELA$, 165, 8)
       Call REPLA(RENUE$, EMPRA$, 173, 64)
       Call REPLA(RENUE$, ORIGE$, 237, 16)
       Call REPLA(RENUE$, MKI$(HOII%), 253, 2)
       Call REPLA(RENUE$, MKI$(0), 255, 2)
          MBAJ$ = ""
       Call REPLA(RENUE$, MBAJ$, 257, 8)
       Call REPLA(RENUE$, OBSER$, 265, 230)
       Call REPLA(RENUE$, PERLIST$, 495, 16)
       Call REGAPP("LIDIRMAI", RENUE$)
       '
     End If
   End If
   '
89 Return
'


End Sub


