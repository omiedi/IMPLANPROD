VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form IMPORAS07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
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
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
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
      Height          =   4245
      Left            =   105
      TabIndex        =   16
      ToolTipText     =   "Importación Definitiva"
      Top             =   3990
      Visible         =   0   'False
      Width           =   11655
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
         ItemData        =   "IMPORAS07.frx":0000
         Left            =   120
         List            =   "IMPORAS07.frx":0002
         TabIndex        =   29
         Top             =   2910
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
         Picture         =   "IMPORAS07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Listado por Impresora"
         Top             =   3255
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
         TabIndex        =   27
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
         Picture         =   "IMPORAS07.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Inicia Importación Definitiva"
         Top             =   1800
         Width           =   750
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   1835
         Left            =   10360
         TabIndex        =   20
         Top             =   700
         Width           =   225
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   120
         Max             =   10000
         TabIndex        =   19
         Top             =   2310
         Width           =   10250
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1965
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   10215
         TabIndex        =   17
         Top             =   420
         Width           =   10215
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
            ItemData        =   "IMPORAS07.frx":0618
            Left            =   0
            List            =   "IMPORAS07.frx":061A
            TabIndex        =   18
            Top             =   250
            Width           =   40000
         End
         Begin VB.PictureBox Picture7 
            Height          =   270
            Left            =   0
            ScaleHeight     =   210
            ScaleWidth      =   89940
            TabIndex        =   40
            Top             =   0
            Width           =   90000
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
               TabIndex        =   42
               Top             =   0
               Width           =   120
            End
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
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   120
            End
         End
         Begin VB.PictureBox Picture3 
            Height          =   1950
            Left            =   -20
            ScaleHeight     =   1890
            ScaleWidth      =   0
            TabIndex        =   21
            Top             =   0
            Width           =   60
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
         TabIndex        =   30
         Top             =   2670
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         Bindings        =   "IMPORAS07.frx":061C
         Height          =   3750
         Left            =   150
         OleObjectBlob   =   "IMPORAS07.frx":062C
         TabIndex        =   15
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
      TabIndex        =   10
      Top             =   2730
      Width           =   740
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   11
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
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
      Height          =   2610
      Left            =   4800
      TabIndex        =   8
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
         Height          =   1860
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
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
      Height          =   2610
      Left            =   105
      TabIndex        =   3
      Top             =   240
      Width           =   4575
      Begin VB.PictureBox Picture6 
         Height          =   2100
         Left            =   210
         ScaleHeight     =   2040
         ScaleWidth      =   4050
         TabIndex        =   34
         Top             =   350
         Width           =   4110
         Begin VB.VScrollBar VScroll1 
            Height          =   2050
            Left            =   3830
            TabIndex        =   39
            Top             =   0
            Width           =   225
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFF0&
            BorderStyle     =   0  'None
            Height          =   4290
            Left            =   0
            ScaleHeight     =   4290
            ScaleWidth      =   4740
            TabIndex        =   35
            Top             =   0
            Width           =   4740
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Tablas de Marca y Modelo"
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
               Left            =   255
               TabIndex        =   43
               Top             =   945
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
               Caption         =   "Lista Productos Terminados"
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
               Left            =   255
               TabIndex        =   38
               Top             =   540
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
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
               Left            =   255
               TabIndex        =   37
               Top             =   330
               Width           =   3735
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFFFF0&
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
               Left            =   255
               TabIndex        =   36
               Top             =   105
               Width           =   3735
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   3750
               Y2              =   3750
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   2800
               Y2              =   2800
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   4560
               Y1              =   2350
               Y2              =   2350
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   -105
               X2              =   4455
               Y1              =   1500
               Y2              =   1500
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   -105
               X2              =   4455
               Y1              =   870
               Y2              =   870
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
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
      Left            =   105
      TabIndex        =   2
      Top             =   2940
      Width           =   11655
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   6915
         TabIndex        =   31
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
            TabIndex        =   32
            Top             =   20
            Width           =   6975
         End
      End
      Begin VB.ComboBox Combo1 
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
      Picture         =   "IMPORAS07.frx":0FDE
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
      Picture         =   "IMPORAS07.frx":12E8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   360
      Width           =   750
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   4620
      TabIndex        =   33
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label NARIM 
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "IMPORAS07"
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
   Select Case XVALO(NARIM)
     Case 0
       Call IMPORMAS
     Case 1
       Call IMPORSTRU
     Case 2
       Call IMPORMAS
     Case 3
       Call IMPORMARMOD
     Case 4
       'Call IMPORCAT
     Case 5
       Call IMPORCLI
     Case 6
       Call IMPORPRO
     Case 7
       Call IMPORPLA
     '\\\OT-05-0207-WAR-25/04/05///
     Case 8
       Call IMPORNCLI
     '\\\OT-05-0207-WAR-FIN///
     Case 9
       Call IMPORCUECLI
     Case 10
       Call IMPORCUEPRO
     Case 11
       Call IMPORCARCHE
     Case 12
       Call IMPORMOVSTK
   End Select
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
   On Error GoTo 0
   Screen.MousePointer = 11
   Label6 = ""
   Select Case XVALO(NARIM)
     Case 0
       Call PREIMPORMAS
     Case 1
       Call PREIMPORSTRU
     Case 2
       Call PREIMPORMASNOR
     Case 3
       Call PREIMPORMARMOD
     Case 4
       Call PREIMPORCAT
     Case 5
       Call PREIMPORCLI
     Case 6
       Call PREIMPORPRO
     Case 7
       Call PREIMPORPLA
     '\\\OT-05-0207-WAR-21/04/05///
     Case 8
       Call PREIMPORNORCLI
     '\\\OT-05-0207-WAR-FIN///
     '\\\OT-05-0272-WAR-09/05/05///
     Case 9
       Call PREIMPORCUECLI
     Case 10
       Call PREIMPORCUEPRO
     '\\\OT-05-0272-WAR-FIN///
     '\\\OT-05-0429-WAR-28/06/05///
     Case 11
       Call PREIMPORCARCHE
     '\\\OT-05-0429-WAR-FIN///
     Case 12
       '\\\OT-5-469-FG-05/07/05///
       Call PREIMPORMOVSTK
       '\\\OT-5-469-FG-FIN///
  End Select
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

Sub PREIMPORMARMOD()
   
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   For U& = 3 To URECORD&
     Call TRACE(20, U&, URECORD&)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     COELE$ = TRIM$(DBGrid1.TEXT)
     If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or CODINT%(COELE$) < 1 Then GoTo 19
     '
     For colui% = 6 To 10 Step 2
       '
       DBGrid1.Col = colui%
       MARCA$ = TRIM$(DBGrid1.TEXT)
       DBGrid1.Col = colui% + 1
       MODELO$ = DBGrid1.TEXT
       '
       If MARCA$ <> "" Then
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, COELE$, 1, 16)
         Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 17, 52)
         '
         Call REPLA(TTXX$, MARCA$, 71, 24)
         Call REPLA(TTXX$, MODELO$, 96, 24)
         '
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
     Next colui%
19 Next U&
   '
   On Error GoTo 0
   Screen.MousePointer = 1
   '

End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   Option1(0).Value = True
   '
   Picture5.Top = 0
   VScroll1.Min = 0
   VScroll1.Max = (Picture5.Height - Picture6.Height) / 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub
'
Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value
  'Label8.Left = (-1) * 10 * HScroll1.Value + 40
  Picture7.Left = (-1) * 10 * HScroll1.Value + 40
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
     '\\\OT-05-0499-WAR-13/07/05///
     List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
     List2.AddItem "  2  Codigo Alternativo         String   <= 24 car"
     List2.AddItem "  3  Descripcion del Articulo   String   <= 64 car"
     List2.AddItem "  4  Desc. Complementaria       String   <= 64 car"
     List2.AddItem "  5  Atributo A                 String   <= 32 car"
     List2.AddItem "  6  Atributo B                 String   <= 32 car"
     '\\\OT-05-0499-WAR-FIN///
     List2.AddItem "  7  Tipo de Material           Num      0 - 255"
     List2.AddItem "  8  Unidad de Medida           String   <= 2 car"
     List2.AddItem "  9  Stock Minimo               Num      0 - 10^8"
     List2.AddItem " 10  Lote Economico             Num      0 - 10^8"
     List2.AddItem " 11  Costo Unitario             Num      0 - 10^8"
     List2.AddItem " 12  Moneda de Costeo           Num      1 - 255"
     List2.AddItem " 13  Fecha de Costeo            Date"
     List2.AddItem " 14  Numero Proveedor Habitual  Num      0 - 32767"
     List2.AddItem " 15  Stock Actual               Num      +/- 999999"
     ENCABE$ = "Código           C-Int  Cod.Alternativo          Descripcion                                                      Descripción Complementaria                                       Atributo A                       Atributo B                         TM UM          St.Min     Lote Ec   Costo Unit   MC   Fecha      Pr-H     Stock"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
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
     List2.AddItem "  5  Maquina Equivalente        String   <= 6 car"
     List2.AddItem "  6  Coeficiente (%)            Num      0 - 999,99"
     List2.AddItem "  7  Observaciones              String   <= 30 car"
     Label8 = "Código  Descripcion                         $ / Hora   Sección   Alter.        Coef.  Observaciones"
   Case 4
     List2.AddItem "  1  Descripcion Categoria      String   <= 27 car"
     List2.AddItem "  2  Costo M.O.D. / hora        Num      0 - 10^6"
     List2.AddItem "  3  Observaciones              String   <= 32 car"
     Label8 = " N-O  Descripcion                       $ / Hora  Observaciones"
   Case 5
     List2.AddItem "  1  Numero Cliente             Num      1 - 9999"
     List2.AddItem "  2  Razón Social del Cliente   String   <= 62 car"
     List2.AddItem "  3  Domicilio                  String   <= 64 car"
     List2.AddItem "  4  Código Postal              String   <= 16 car"
     List2.AddItem "  5  Localidad                  String   <= 48 car"
     List2.AddItem "  6  Teléfono                   String   <= 24 car"
     List2.AddItem "  7  Fax                        String   <= 24 car"
     List2.AddItem "  8  Persona de Contacto        String   <= 32 car"
     List2.AddItem "  9  CUIT                       String   =  8 car"
     List2.AddItem " 10  Grupo de Cliente           Num      0 - 255"
     List2.AddItem " 11  Codigo de Cuenta           String   <= 12 car"
     List2.AddItem " 12  Número de Vendedor         Num      0 - 255"
     List2.AddItem " 13  Codigo de Provincia        String   <= 1 car"
     List2.AddItem " 14  Categoría I.V.A.           Num      0 - 7"
     List2.AddItem " 15  Condicion de Pago          Num      1 - 255"
     List2.AddItem " 16  Límite de Credito          Num      0 - 10^9"
      Label8 = " Cta. Razón Social                                    Domicilio                                       " & _
     "C.Postal        Localidad                   Teléfono                Fax                     Contacto                CUIT             Gp.  Cod.Cta       V.  Pr  PI  CP     Lim.Cre"
   Case 6
     List2.AddItem "  1  Numero Proveedor           Num      1 - 9999"
     List2.AddItem "  2  Razón Social del Proveedor String   <= 62 car"
     List2.AddItem "  3  Domicilio                  String   <= 64 car"
     List2.AddItem "  4  Código Postal              String   <= 16 car"
     List2.AddItem "  5  Localidad                  String   <= 48 car"
     List2.AddItem "  6  Teléfono                   String   <= 24 car"
     List2.AddItem "  7  Fax                        String   <= 24 car"
     List2.AddItem "  8  Persona de Contacto        String   <= 32 car"
     List2.AddItem "  9  CUIT                       String   =  8 car"
     List2.AddItem " 10  Tipo de Cuenta             Num      0 - 1"
     List2.AddItem " 11  Codigo de Cuenta           String   <= 16 car"
     '\\\OT-05-0202-WAR-19/04/05///
     List2.AddItem " 12  Codigo de Provincia        String   <=1 car"
     List2.AddItem " 13  Categoria de IVA           Num      0 - 7"
     List2.AddItem " 14  Cuenta Madre               String   <= 12 car"
     List2.AddItem " 15  Nombre de Fantasía         String   <= 12 car"
     ENCABE$ = "Cta.  Razón Social                                    Domicilio                                       C.Postal    Localidad                       " + _
     "Teléfono                Fax                     Contacto                CUIT            T.   Cod.Cta.    Pr. IVA  Cta. Madre    N.Fantasia"
     Label8 = Left$(ENCABE$, 256)
     Label82 = Mid$(ENCABE$, 257)
     Label82.Left = Label8.Left + Label8.Width
     '\\\OT-05-0202-WAR-FIN///
   Case 7
     List2.AddItem "  1  Nivel de Agrupación        Num      0 - 5"
     List2.AddItem "  2  Codigo de Cuenta           String   <= 12 car"
     List2.AddItem "  3  Denominación de Cuenta     String   <= 64 car"
     List2.AddItem "  4  Tipo de Cuenta             String   D/H/I/G"
     List2.AddItem "  5  Imputable                  String   Si / No"
     Label8 = "  Nvl Codigo        Descripcion                                       T.  ID"
     '\\\OT-05-0207-WAR-21/04/05///
   Case 8
     List2.AddItem "  1  Cod. Cliente               String   0 - 12"
     List2.AddItem "  2  Codigo de Articulo         String   <= 16 car"
     List2.AddItem "  3  Nro. Original              String   <= 28 car"
     Label8 = "Cliente      Articulo          Descripcion             Nro.Original"
     '\\\OT-05-0207-WAR-FIN///
     '\\\OT-05-0272-WAR-09/05/05///
   Case 9
     List2.AddItem "  1  Cod. Cliente               String   0 - 12"
     List2.AddItem "  2  Referencia                 String   <= 62 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Nro. Comprobante Largo     String   <= 24 car"
     List2.AddItem "  6  Nro. Comprobante Corto     String   <= 12 car"
     List2.AddItem "  7  Importe Original           Num      0 - 10^8"
     List2.AddItem "  8  Importe Pendiente          Num      0 - 10^8"
     Label8 = "Cliente             Referencia                      F.Emisión F.Vencim.   Comp.Largo          Comp.Corto       Imp.Orig.    Imp.Pend. "
   Case 10
     List2.AddItem "  1  Cod. Prov.                 String   <= 12 car"
     List2.AddItem "  2  Referencia                 String   <= 62 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Nro. Comprobante Largo     String   <= 24 car"
     List2.AddItem "  6  Nro. Comprobante Corto     String   <= 12 car"
     List2.AddItem "  7  Importe Original           Num      0 - 10^8"
     List2.AddItem "  8  Importe Pendiente          Num      0 - 10^8"
     Label8 = "Cliente             Referencia                      F.Emisión F.Vencim.   Comp.Largo          Comp.Corto       Imp.Orig.    Imp.Pend. "
     '\\\OT-05-0272-WAR-FIN///
   Case 11
    '\\\OT-05-0429-WAR-28/06/05///
     List2.AddItem "  1  Banco Emisor               String   <= 18 car"
     List2.AddItem "  2  Nro. Cheque                String   <= 12 car"
     List2.AddItem "  3  Fecha de Emisión           Date"
     List2.AddItem "  4  Fecha de Vencimiento       Date"
     List2.AddItem "  5  Importe Original           Num      0 - 10^9"
     List2.AddItem "  6  Documento Recep. Largo     String   <= 18 car"
     List2.AddItem "  7  Documento Recep. Corto     String   <= 12 car"
     List2.AddItem "  8  Cliente / Origen           String   <= 30 car"
     Label8 = "Banco               Nro. Cheque   F.Emis.   F.Venc.         Importe   Recibo Largo        Recibo Corto  Cliente / Origen"
    '\\\OT-05-0429-WAR-FIN///
   Case 12
     List2.AddItem "  1  Codigo de Articulo         String   <= 15 car"
     List2.AddItem "  2  Descripcion del Articulo   String   <= 64 car"
     List2.AddItem "  3  Fecha de Movimiento        Date"
     List2.AddItem "  4  Depósito Número            Num      1 - 255"
     List2.AddItem "  5  Cantidad (+/-)             Num      0 - 10^8"
     List2.AddItem "  6  Codigo de Movimiento       String   <= 2 car"
     List2.AddItem "  7  Documento Primario         String   <= 12 car"
     List2.AddItem "  8  Documento Secundario       String   <= 12 car"
     List2.AddItem "  9  Valor Unitario             Num      0 - 10^8"
     List2.AddItem " 10  Moneda de Valuacion        String   <= 2 car"
     List2.AddItem " 11  Código de Cliente          String   <= 12 car"
     List2.AddItem " 12  Código de Proveedor        String   <= 12 car"
     List2.AddItem " 13  Detalle / Referencia       String   <= 32 car"
     Label8 = "Código           C-Int  Descripción                                                       Fecha      Dep        Cant. CM  D.Prim.       D.Sec.              V.Unit.  M.  Cli.  Pro. Detalle / Referencia"
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
   If DBGrid1.Columns.Count < 10 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     .Recordset.MoveNext
     .Recordset.MoveNext
     .Recordset.MoveNext
     '
     U& = 0
     Do While Not .Recordset.EOF
        U& = U& + 1
        Call TRACE(20, U&, URECORD&)
        TTXX$ = Space$(254)
        codex$ = ""
        On Error Resume Next
        codex$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        On Error GoTo 0
        If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 15 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
        CDXI% = CODINT%(codex$)
        If CDXI% > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + codex$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      Call REPLA(TTXX$, codex$, 1, 16)
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        COALT$ = ""
        On Error Resume Next
        COALT$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
'        If Len(COALT$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Código Alternativo"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, COALT$, 25, 24)
        '\\\OT-05-0499-WAR-FIN///
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(2))
        On Error GoTo 0
        If Len(DEDEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DEDEX$, 50, 64)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        DECOM$ = ""
        On Error Resume Next
        DECOM$ = TRIM$(.Recordset.Fields(3))
        On Error GoTo 0
        If Len(DECOM$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion Complementaria"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DECOM$, 115, 64)
        '
        ATRA$ = ""
        On Error Resume Next
        ATRA$ = TRIM$(.Recordset.Fields(4))
        On Error GoTo 0
        Call REPLA(TTXX$, ATRA$, 180, 32)
        '
        ATRB$ = ""
        On Error Resume Next
        ATRB$ = TRIM$(.Recordset.Fields(5))
        On Error GoTo 0
'        If Len(ATRB$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo B"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, ATRB$, 213, 32)
        '
        VALO$ = "1"   ' FUERZA MATERIA PRIMA
        Call REPLA(TTXX$, VALO$, 246, 4)
        '
        UME$ = "U."
        Call REPLA(TTXX$, UME$, 251, 4)
        '
        '
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORMASNOR()
   '
   If DBGrid1.Columns.Count < 7 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
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
        TTXX$ = Space$(254)
        codex$ = ""
        On Error Resume Next
        codex$ = .Recordset.Fields(0)   ' TRIM$(DBGrid1.TEXT)
        On Error GoTo 0
        If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 15 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
        CDXI% = CODINT%(codex$)
        If CDXI% > 0 Then
          If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
          If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
          ATX% = COMALTER%("Código / Campo Clave '" + codex$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
          If ATX% <= 2 Then GoTo 19 ' SALTEAR
        End If
        '
12      Call REPLA(TTXX$, codex$, 1, 16)
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        COALT$ = ""
        On Error Resume Next
        COALT$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
'        If Len(COALT$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Código Alternativo"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, COALT$, 25, 24)
        '\\\OT-05-0499-WAR-FIN///
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(2))
        On Error GoTo 0
        If Len(DEDEX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DEDEX$, 50, 64)
        '
        '\\\OT-05-0499-WAR-13/07/05///
        DECOM$ = ""
        On Error Resume Next
        DECOM$ = TRIM$(.Recordset.Fields(3))
        On Error GoTo 0
'        If Len(DECOM$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion Complementaria"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, DECOM$, 115, 64)
        '
        ATRA$ = ""
        On Error Resume Next
        ATRA$ = TRIM$(.Recordset.Fields(4))
        On Error GoTo 0
'        If Len(ATRA$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo A"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, ATRA$, 180, 32)
        '
        ATRB$ = ""
        On Error Resume Next
        ATRB$ = TRIM$(.Recordset.Fields(5))
        On Error GoTo 0
'        If Len(ATRB$) < 1 Then
'          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Atributo B"
'          List1.AddItem MERROX$
'          GoTo 19
'        End If
        Call REPLA(TTXX$, ATRB$, 213, 32)
        '
        '\\\OT-05-0499-WAR-FIN///
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(6), 4)
        On Error GoTo 0
        If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Tipo de Material no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 246, 4)
        '
        UME$ = "U."
'        On Error Resume Next
'        UME$ = TRIM$(.Recordset.Fields(7))
'        On Error GoTo 0
        Call REPLA(TTXX$, UME$, 251, 4)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(8))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Stock Minimo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 256, 12)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(9))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.1")
        If XVALO(VALO$) < 0 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Lote de Reposicion no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 269, 12)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(10))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALOR, "F12.4")
        Call REPLA(TTXX$, VALO$, 282, 12)
        '
        On Error Resume Next
        VALO$ = AJUSTD$(.Recordset.Fields(11), 4)
        On Error GoTo 0
         If XVALO(VALO$) < 0 Or XVALO(VALO$) > 255 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Moneda de Costeo no Procesable"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, VALO$, 295, 4)
        '
        VALO$ = ""
        On Error Resume Next
        VALO$ = Format$(.Recordset.Fields(12), "dd/mm/yy")
        On Error GoTo 0
        Call REPLA(TTXX$, VALO$, 300, 8)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(13))
        On Error GoTo 0
        If VALOR < 0 Or VALOR > 32766 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F8.0")
        Call REPLA(TTXX$, VALO$, 309, 8)
        '
        VALOR = 0
        On Error Resume Next
        VALOR = CDbl(.Recordset.Fields(14))
        On Error GoTo 0
        If VALOR < (-9999999) Or VALOR > 9999999 Then VALOR = 0
        VALO$ = FORMATNUM$(VALOR, "F10.1")
        Call REPLA(TTXX$, VALO$, 318, 10)
        '
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$ + "*")
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
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
   For colui% = 13 To DBGrid1.Columns.Count - 1 Step 2
     DBGrid1.Row = 0
     On Error Resume Next
     DBGrid1.Col = colui% + 1
     If Err Then GoTo 29
     On Error GoTo 0
     codex$ = DBGrid1.TEXT
     '
     If codex$ = "" Then GoTo 29
     CODICON% = CODINT(codex$)
     If CODICON% < 1 Then
       Call MENSERR(24, "Codigo '" + codex$ + "' no Registrado")
       GoTo 29
     End If
     '
     ATX% = 0: CODEXA$ = ""
     For U& = 3 To URECORD&
       Call TRACE(20, U&, URECORD&)
       TTXX$ = Space$(254)
       Call REPLA(TTXX$, codex$, 1, 16)
       '
       DBGrid1.Row = U& - 1
       DBGrid1.Col = colui%
       VALOR = 0
       On Error Resume Next
       VALOR = CDbl(DBGrid1.TEXT)
       On Error GoTo 0
       '
       If VALOR > 0 Then
         VALO$ = FORMATNUM$(VALOR, "F12.4")
         Call REPLA(TTXX$, VALO$, 69, 12)
         '
         DBGrid1.Col = colui% + 1
         REFO$ = TRIM$(DBGrid1.TEXT)
         TTXX$ = Left$(TTXX$, 86) + REFO$
         '
         DBGrid1.Col = 0
         COELE$ = TRIM$(DBGrid1.TEXT)
         If COELE$ = "" Or Len(COELE$) < 3 Or Len(COELE$) > 15 Or COELE$ = codex$ Or CODINT%(COELE$) < 1 Then
            Call MENSERR(24, "Codigo Elemento '" + COELE$ + "' no Registrado")
            GoTo 19
         End If
         Call REPLA(TTXX$, COELE$, 19, 16)
         Call REPLA(TTXX$, DESCRIT0$(CODINT%(COELE$)), 37, 32)
         '
         List4.AddItem TTXX$
         Label6 = TRIM$(Str$(List4.ListCount))
         If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
           TEPRUE = RTrim(TTXX$ + "*")
         End If
         On Error Resume Next
         List4.TopIndex = List4.ListCount - 7
         On Error GoTo 0
         DoEvents
         '
       End If
19   Next U&
29 Next colui%
   '
   On Error GoTo 0
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
     codex$ = TRIM$(DBGrid1.TEXT)
     If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 15 Or CODINT%(codex$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo de Producto Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
     If codex$ = CODEXA$ Then GoTo 10 ' EL MISMO CODIGO
     '
     CODEXA$ = codex$
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CODINT%(codex$)))
     If Len(RFF$) > 0 And ATX% = 2 Then GoTo 19
     '
     'ATX% = 3   ' SOBREESCRIBIR
     If Len(RFF$) > 0 Then
         ATX% = COMALTER%("Secuencia Operaciones de Código '" + codex$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       Else
         GoTo 12
     End If
     '
10   If ATX% <= 2 And Len(RFF$) > 0 Then GoTo 19 ' SALTEAR
     '
12   Call REPLA(TTXX$, codex$, 1, 16)
     If codex$ <> CODEYA$ Then NOPE% = 0
     CODEYA$ = codex$
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
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
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
     codex$ = TRIM$(DBGrid1.TEXT)
     If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 6 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Maquina Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     RFF$ = RECFILT$("PADMAQ", 1, codex$)
     If Len(RFF$) > 0 Then
       If ATX% = 2 Then GoTo 19  ' SALTEAR SIEMPRE
       If ATX% = 4 Then GoTo 12  ' SOBREESCRIBE SIEMPRE
       ATX% = COMALTER%("Código / Campo Clave '" + codex$ + "' Preexistente en Base de Datos\\Conservar\Conservar Siempre\Reemplazar\Reemplazar Siempre")
       If ATX% <= 2 Then GoTo 19 ' SALTEAR
     End If
     '
12   Call REPLA(TTXX$, codex$, 1, 6)
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
     If VTX$ <> "" And XVALO(VTX$) < 1 Then
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
     DBGrid1.Col = 4                    ' Maquina equivalente
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
     DBGrid1.Col = 5                     ' Coeficiente de equivalencia
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
     DBGrid1.Col = 6                    ' Observaciones
     OBSER$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, OBSER$, 87, 30)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
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
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
       '
     End If
19 Next U&
   
20 End Sub

Sub PREIMPORCLI()
   '
   If DBGrid1.Columns.Count < 11 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(284)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0
     NUCLIE = XVALO(TRIM$(DBGrid1.TEXT))
     If NUCLIE < 1 Or NUCLIE > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cliente no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     codex$ = AJUSTD$(Str$(Int(NUCLIE + 0.5)), 5)
12   Call REPLA(TTXX$, codex$, 1, 5)
     '
     DBGrid1.Col = 1
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social de Cliente"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 64)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 48)
     Call REPLA(TTXX$, DOMI$, 55, 48)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 12)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 32)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DBGrid1.Col = 9                     ' GRUPO DE CLIENTE
     GRUPOCLI$ = AJUSTI$(DBGrid1.TEXT, 4)
     Call REPLA(TTXX$, GRUPOCLI$, 236, 4)
     '
     DBGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODICUEN$, 241, 12)
     '
     DBGrid1.Col = 11                    ' VENDEDOR
     VENDE$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, VENDE$, 253, 4)
     '
     DBGrid1.Col = 12                    ' PROVINCIA
     PROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, PROVI$, 259, 2)
     '
     DBGrid1.Col = 13                    ' CATEGORIA IVA
     CIVA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, CIVA$, 261, 4)
     '
     DBGrid1.Col = 14                    ' CONDICION DE PAGO
     COPA$ = FORMATNUM$(XVALO(TRIM$(DBGrid1.TEXT)), "I4")
     Call REPLA(TTXX$, COPA$, 265, 4)
     '
     DBGrid1.Col = 15                    ' LIMITE DE CREDITO
     VALOR = 0
     On Error Resume Next
     VALOR = CDbl(DBGrid1.TEXT)
     On Error GoTo 0
     If VALOR < 0 Then VALOR = 0
     VALO$ = FORMATNUM$(VALOR, "F12.2")
     Call REPLA(TTXX$, VALO$, 269, 12)
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
End Sub

Sub PREIMPORPRO()
   '
   If DBGrid1.Columns.Count < 15 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   URECORD& = Data1.Recordset.RecordCount
   'MODIF ECHA POR UNA IMPORTACIÓN ECHA EN TELEBIT 16/06/05
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   Call APERBASDAT("PROVE")
   ATX% = 0
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(290)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0
     NUPROV& = XVALO(TRIM$(DBGrid1.TEXT))   ' NUMERO DE CUENTA
     If NUPROV& < 1 Or NUPROV& > 32000 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Proveedor no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     With Proved12
       .FindFirst "Nume_Cli = " & NUPROV& & " "
       If .NoMatch = False Then
         MERROX$ = "Cuenta Número '" + TRIM$(Str$(NUPROV&)) + "' Existente en B.Datos."
         List1.AddItem MERROX$
         GoTo 19
       End If
     End With
     '
     codex$ = AJUSTD$(Str$(Int(NUPROV + 0.5)), 5)
12   Call REPLA(TTXX$, codex$, 1, 5)
     '
     DBGrid1.Col = 1                     ' RAZON SOCIAL
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta R.Social del Proveedor"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 7, 48)
     '
     DBGrid1.Col = 2                     ' DOMICILIO
     DOMI$ = AJUSTI$(DBGrid1.TEXT, 64)
     Call REPLA(TTXX$, DOMI$, 55, 64)
     '
     DBGrid1.Col = 3                     ' CODIGO POSTAL
     CODPOS$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODPOS$, 103, 16)
     '
     DBGrid1.Col = 4                     ' LOCALIDAD
     LOCA$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, LOCA$, 115, 48)
     '
     DBGrid1.Col = 5                     ' TELEFONO
     TELEF$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, TELEF$, 147, 24)
     '
     DBGrid1.Col = 6                     ' FAX
     FAX$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FAX$, 171, 24)
     '
     DBGrid1.Col = 7                     ' PERSONA DE CONTACTO
     CONTACTO$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CONTACTO$, 195, 24)
     '
     DBGrid1.Col = 8                     ' CUIT
     CUIT$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CUIT$, 219, 16)
     '
     DBGrid1.Col = 9                     ' TIPO DE CUENTA
     TIPOCTA$ = TRIM$(DBGrid1.TEXT)
     If TIPOCTA$ = "" Then TIPOCTA$ = "0"
     If TIPOCTA$ = "0" Or TIPOCTA$ = "1" Then
       Call REPLA(TTXX$, TIPOCTA$, 235, 4)
     Else
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Tipo de Cuenta"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     DBGrid1.Col = 10                    ' CODIGO DE CUENTA
     CODICUEN$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODICUEN$, 240, 12)
     '
     '\\\OT-05-0202-WAR-18/04/05///
     DBGrid1.Col = 11                    ' CODIGO DE PROVINCIA
     CODIPROVI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CODIPROVI$, 253, 2)
     '
     DBGrid1.Col = 12                    ' CODIGO DE POSICION IVA
     CODIVA% = XVALO(DBGrid1.TEXT)
     Call REPLA(TTXX$, FORMATNUM$(CODIVA%, "I4"), 255, 4)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 13                    ' CUENTA MADRE
     CTAMADRE$ = TRIM$(DBGrid1.TEXT)
     If Len(CTAMADRE$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Error en Longitud de Cuenta Madre"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, CTAMADRE, 261, 12)
     '\\\OT-05-0202-WAR-FIN///
     '
     DBGrid1.Col = 14                    ' numbre de fantasia
     FANTASI$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, FANTASI$, 275, 16)
     '
    '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&
   '
20 DBGrid1.RowHeight = hrow
   Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORPLA()
   '
   If DBGrid1.Columns.Count < 5 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   ATX% = 0
   URECORD& = Data1.Recordset.RecordCount
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(254)
     DBGrid1.Row = U& - 1
     DBGrid1.Col = 0
     '
     NIVEI% = XVALO(DBGrid1.TEXT)
     If NIVEI% < 0 Or NIVEI% > 5 Then NIVEI% = 0
     Call REPLA(TTXX$, AJUSTD$(Str$(NIVEI%), 4), 1, 4)
     '
     DBGrid1.Col = 1
     codex$ = TRIM$(DBGrid1.TEXT)
     If codex$ = "" Or Len(codex$) < 1 Or Len(codex$) > 12 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, codex$, 7, 12)
     '
     For KKU% = 1 To List4.ListCount
       If TRIM$(Mid$(List4.List(KKU% - 1), 7, 12)) = codex$ Then
         MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Cuenta Duplicado o Pre-Existente"
         List1.AddItem MERROX$
         GoTo 19
       End If
     Next KKU%
     '
     DBGrid1.Col = 2
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Descripcion"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, DEDEX$, 21, 48)
     '
     DBGrid1.Col = 3                    ' Seccion
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ <> "D" And VTX$ <> "H" And VTX$ <> "I" Then VTX$ = "G"
     Call REPLA(TTXX$, VTX$, 71, 1)
     '
     DBGrid1.Col = 4
     VTX$ = TRIM$(UCase$(DBGrid1.TEXT))
     If VTX$ = "0" Or VTX$ = "NO" Then VTX$ = "NO"
     If VTX$ <> "NO" Then VTX$ = "SI"
     Call REPLA(TTXX$, VTX$, 75, 2)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
     '
19 Next U&
   '
   DBGrid1.RowHeight = hrow   ' RESTAURA ALTURA ANTERIOR
   '
20 End Sub
 '\\\OT-05-0207-WAR-21/04/05///
Sub PREIMPORNORCLI()
   '
   If DBGrid1.Columns.Count < 3 Then
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
     DBGrid1.Col = 0            'CLIENTE CODIGO STRING
     codex$ = TRIM$(DBGrid1.TEXT)
     If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 15 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     If U& = 250 Then
       XX = "AA"
     End If
     '
12   Call REPLA(TTXX$, codex$, 1, 12)
     '
     DBGrid1.Col = 1            ' ARTICULO
     DEDEX$ = TRIM$(DBGrid1.TEXT) '+ " " + CODEX$ + " "
     If Len(DEDEX$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Articulo"
       List1.AddItem MERROX$
       GoTo 19
     End If
     CDXI% = CODINT%(DEDEX$)
     If CDXI% < 0 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
     Call REPLA(TTXX$, DEDEX$, 14, 16)
     Call REPLA(TTXX$, DESCRIT0(CDXI%), 32, 22)
     '
     DBGrid1.Col = 2                     ' NUMERO ORIGINAL
     VALO$ = TRIM$(DBGrid1.TEXT)
     If Len(VALO$) < 0 Or Len(VALO$) > 28 Or TRIM(VALO$) = "" Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero Original no Procesable"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, VALO$, 56, 28)
     '
     List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&

20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0207-WAR-FIN///
Sub IMPORMAS()
   '
   FIN& = List4.ListCount
   Call APERBASDAT("MASTER")
   '
   NMS% = NUMAS%
   'REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     '\\\OT-05-0499-WAR-13/07/05///
       'ACA TOMAR NUEVOS CAMPOS
       CODALT$ = TRIM$(Mid$(TTXX$, 25, 24))
       DESCOMP$ = TRIM$(Mid$(TTXX$, 115, 64))
       ATRIA$ = TRIM$(Mid$(TTXX$, 180, 32))
       ATRIB$ = TRIM$(Mid$(TTXX$, 213, 32))
     '\\\OT-05-0499-WAR-FIN///
     DESIT$ = TRIM$(Mid$(TTXX$, 50, 64))
     TIMAT% = XVALO(Mid$(TTXX$, 246, 4)) Mod 256
     UNIME$ = TRIM$(Mid$(TTXX$, 251, 4))
     STOMI = XVALO(Mid$(TTXX$, 256, 12))
       If STOMI > 30000 Then STOMI = (-1) * STOMI / 100
       If STOMI < (-30000) Then STOMI = (-30000)
     LECON = XVALO(Mid$(TTXX$, 269, 12))
       If LECON > 30000 Then LECON = (-1) * LECON / 100
       If LECON < (-30000) Then LECON = (-30000)
     LOMAX = 1
       If LOMAX < Abs(LECON) Then LOMAX = LECON
     COSUN = XVALO(Mid$(TTXX$, 282, 12))
     MONEC% = XVALO(Mid$(TTXX$, 295, 4)) Mod 256
     FECOU% = FECHANUM(Mid$(TTXX$, 300, 8))
       PRH = Abs(XVALO(Mid$(TTXX$, 309, 8)))
       While PRH > 30000: PRH = PRH - 10000: Wend
     PRHAB% = PRH
     '
     STOAC = XVALO(Mid$(TTXX$, 318, 10))
     '
     On Error GoTo 0
     If CODINT%(CODIT$) = 0 Then
       With Master
         .AddNew   ' YA CONVERTIDO
         ON ERROR RESUME NEXT
         !CodiEmpr = Codiemp%
         ON ERROR GOTO 0
         '\\\OT-05-0499-WAR-13/07/05///
         !NPLANO = CODALT$
         !DESCRIDOS = DESCOMP$
         !TALLE = ATRIA$
         !COLOR = ATRIB$
         '\\\OT-05-0499-WAR-FIN///
         !CODIGO = CODIT$
         !DESCRIPCION = DESIT$
         !Status = 1
           NMS% = NMS% + 1
         !CODINT = NMS%
         !UMEDIDA = AJUSTI$(UNIME$, 4)
         !TIMATE = TIMAT%
         !CRITREP = 0
         !Monecos = MONEC%
         !FECOSTO = CVDAT(FECOU%)
         !COSUNI = COSUN
         !STOMIN = STOMI
         !LOREPOS = LECON
         !STOMAX = LOMAX
         !PROVHABI = PRHAB%
         '
         !Alias = " "
         !PORNACIO = 0
         !PESUNI = 0
         '!COLOR = " "
         '!TALLE = " "
         !DESCOMPO = 0
         !GRUPIVA = 0
         !Grucoven = 0
         !CLinea = " "
         !Cfamilia = " "
         '!NPLANO = " "
         !Revision = " "
         !FEREVIS = CVDAT(0)
         !Observa = " "
         !ESPECIFICACION = " "
        .Update
       End With
       '
       If Abs(STOAC) >= 0.05 Then
         HOII% = HOY(HOS$)
         NORI% = NORI% + 1
         Call MOVISTOK(HOII%, NMS%, "", "AJ", "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "AJ-" + HOS$, "Stock Importado", NORI%, CDbl(COSUN), MONEC%, 0, 0, STOAC)
       End If
       '
     End If
     '
   Next U&
   '
   Master.Close
   UPDATMASTER% = 1
   Call GENECOMAS
   Screen.MousePointer = 1
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
         TTYY$ = Space$(128)
         Call REPLA(TTYY$, Mid$(TTXX$, 19, 16), 1, 16)
         Call REPLA(TTYY$, Mid$(TTXX$, 69, 12), 49, 12)
         Call REPLA(TTYY$, Mid$(TTXX$, 81, 4), 61, 4)
         TTYY$ = TTYY$ + Mid$(TTXX$, 87)
         CAIT% = CAIT% + 1
         EXPLOLIN$(CAIT%) = TTYY$
       Next UKU&
       Call EXPLOGR(CODEXT$(CDXI%), 1)
     End If
   Next CDXI%
   '
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
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 79, 10))), 59, 4)
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 89, 10))), 63, 4)
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 99, 12))), 67, 4)
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 111, 10))), 91, 4)
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 121, 8))), 75, 4)
         Call REPLA(TTYY$, MKS$(XVALO(Mid$(TTXX$, 129, 8))), 83, 4)
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
Sub IMPORMARMOD()
   '
   FIN& = List4.ListCount
   REBLA$ = REGBLAN$("MARCAMOD")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODE$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = CODINT%(CODE$)
     If CIXI% > 0 Then
       MARCA$ = Mid$(TTXX$, 71, 24)
       MODELO$ = Mid$(TTXX$, 96, 24)
       TTYY$ = MKI$(CIXI%) + MARCA$ + MODELO$
       RFF$ = RECFILT$("MARCAMOD", 1, MKI$(CIXI%))
       For KU& = 1 To Len(RFF$) Step 4
         REMA& = CVS(Mid$(RFF$, KU&, 4))
         XX$ = REGLEIDO$("MARCAMOD", REMA&)
         If Left$(XX$, 50) = TTYY$ Then GoTo 19
       Next KU&
       Call REGAPP("MARCAMOD", TTYY$ + Space$(128))
       Call CIERRARCH("MARCAMOD")
     End If
19 Next U&
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
     ICAT% = XVALO(Left$(TTXX$, 4)) Mod 256
     DECAT$ = TRIM$(Mid$(TTXX$, 7, 30))
     CUHOR = XVALO(Mid$(TTXX$, 37, 12))
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

Sub IMPORCLI()
   '
   PUNUCLI$ = String$(32767, 0)
   XX0$ = REGBLAN$("DEUDOR1") + REGBLAN$("DEUDOR2")
   '
   ' Genera DEUDOR1 y DEUDOR2
   Call ACLI_GES07.GENECOCLI
   '
   ' Establece en el "vector" PUNUCLI$, los numeros de clientes ya utilizados
   For U& = 1 To ULTREG&("DEUDOR1")
     XX$ = REGLEIDO$("DEUDOR1", U&)
     NCI% = CVI(Left$(XX$, 2))
     If NCI% > 0 Then
       If NCI% <= 32767 Then
         Mid$(PUNUCLI$, NCI%, 1) = Chr$(1)
       End If
     End If
   Next U&
   '
   Call APERBASDAT("CLIEN")
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI% = XVALO(Mid$(TTXX$, 1, 5))
     If NCI% < 1 Then GoTo 19
     ' Verifica si el numero ya ha sido utilizado
     If Asc(Mid$(PUNUCLI$, NCI%, 1)) > 0 Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 7, 48)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 16))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     GruCli% = XVALO(Mid$(TTXX$, 235, 4))
     CODICUEN$ = Mid$(TTXX$, 241, 12)
     '\\\OT-05-0192-WAR-18/04/05///
     PIVA% = XVALO(Mid$(TTXX$, 261, 4))
     CONDI% = XVALO(Mid$(TTXX$, 253, 4))
     PROVI$ = TRIM$(Mid$(TTXX$, 259, 2))
     LIPRE% = 0
     COVEN% = XVALO(Mid$(TTXX$, 265, 4))
     LICRE = XVALO(Mid$(TTXX$, 269, 12))
     CUEMA% = 0
     'GoTo 19 'PARA TESTEAR
     '\\\OT-05-0192-WAR-FIN///
     '
     On Error GoTo 0
     With Deudor12
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !Stat_Cli = 1                        'STATUS ACTIVO
       !Nume_Cli = NCI%                     'CODIGO ENTERO
       !Codi_Cli = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = 0                        'TIPO CTA
       !GrCo_Cli = GruCli%                  'TIPO CLI  VA A GRUPO CLI
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = Left$(RASO$, 10)         'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = PROVI$                  'COD PROVINCIA
       '!Pais_Cli = PAISOK$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       !Cpag_Cli = COVEN%                   'CONDICION DE PAGO
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       '!Mail_Cli = MAILPRO$                'MAIL
       !Piva_Cli = PIVA%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Lcre_Cli = LICRE                    'LIMITE CREDITO
       !Vend_Cli = CONDI%                   'CODIGO VENDEDOR
       !Ctac_Cli = CONTACTX$
       .Update
     End With
     '
     Call ACLI_GES07.GENECOCLI
     '
     'CONTAD% = CONTAD% + 1
     '
     If NCI% > 0 Then
        Mid$(PUNUCLI$, NCI%, 1) = Chr$(1)
     End If
     '
19 Next U&
   '
End Sub

Sub IMPORPRO()
   '
   XX0$ = REGBLAN$("PROVED1") + REGBLAN$("PROVED2")
   '
   Call APERBASDAT("PROVE")
   '
   FIN& = List4.ListCount
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     '
     NCI% = XVALO(Mid$(TTXX$, 1, 5))
     If NCI% < 1 Then GoTo 19
     CODICUEN$ = Mid$(TTXX$, 240, 12)
     If TRIM$(CODICUEN$) = "" Then GoTo 19
     '
     RASO$ = Mid$(TTXX$, 7, 48)
     DOMI$ = Mid$(TTXX$, 55, 48)
     CPOSTX$ = Mid$(TTXX$, 103, 12)
     CPOST% = 0: If XVALO(CPOSTX$) < 32767 Then CPOST% = XVALO(CPOSTX$)
     LOCA$ = Mid$(TTXX$, 115, 32)
     TELE$ = Mid$(TTXX$, 147, 24)
     FAXX$ = Mid$(TTXX$, 171, 24)
     CONTACTX$ = Mid$(TTXX$, 195, 24)
     CUIT$ = TRIM$(Mid$(TTXX$, 219, 16))
       If CUIT$ <> "" Then VCC% = VALICUIT%(CUIT$)
     TIPOCTA% = XVALO(Mid$(TTXX$, 235, 4))
     CODIPROVI1$ = Mid$(TTXX$, 253, 2)
     CODIVA1% = XVALO(Mid$(TTXX$, 255, 4))
     CTAMADRE$ = TRIM$(Mid$(TTXX$, 261, 12))
     FANTASI$ = TRIM$(Mid$(TTXX$, 275, 16))
     'PIVA% = 1
     IBRU$ = ""
     CONDI% = 0
     LIPRE% = 0
     COVEN% = 0
     LICRE = 0
     CUEMA% = 0
     '
     With Proved12
       .FindFirst "Codi_Cli = '" & CODICUEN$ & "'"
       If .NoMatch = False Then
           OPX% = COMALTER%("Cuenta '" + TRIM$(CODICUEN$) + "' Existente en B.Datos\\Actualizar\Conservar\Cancelar Proceso")
           If OPX% < 1 Or OPX% > 2 Then GoTo 99
           If OPX% = 2 Then GoTo 19
           .Edit
         Else
           .AddNew   ' YA CONVERTIDO
           ON ERROR RESUME NEXT
           !CodiEmpr = Codiemp%
           ON ERROR GOTO 0
           !Nume_Cli = NCI%                     'CODIGO ENTERO
       End If
       !Stat_Cli = 1                        'STATUS ACTIVO
       !Codi_Cli = CODICUEN$                'CODIGO TEXT
       !Tcpr_Cli = TIPOCTA%                      'TIPO CTA
       !Raso_Cli = RASO$                    'RAZON SOCIAL
       !Fant_Cli = FANTASI$                 'FANTASIA
       !Domi_Cli = DOMI$                    'DIRECCION
       !Loca_Cli = LOCA$                    'LOCALIDAD
       !Prov_Cli = CODIPROVI1$              'COD PROVINCIA
       '!Pais_Cli = PAISOK$                 'PAIS VALIDADO
       !Cpos_Cli = CPOSTX$                  'CODIGO POSTAL
       !Tele_Cli = TELE$                    'TELEFONO
       '!Tele2_Cli = TELEF2$                '2º TELEFONO
       !Faxi_Cli = FAXX$                    'FAX
       !Cuit_Cli = CUIT$                    'CUIT SIN VALIDAR
       '!Cpag_Cli = 0                       'POR DEFAULT (XQ' NO SE JUSTIFICAN 60/900)
       '!Nibr_Cli = INGBRU$                 'IMP BRUTOS
       '!Mail_Cli = MAILPRO$                'MAIL
       !Piva_Cli = CODIVA1%                    'POS IVA
       '!Posivat_Cli = PIVATXT$             'POSICION DEL IVA EN TEXTO
       '!C_Ret = CRET$                      'RETENCION
       !Ctac_Cli = CONTACTX$
       CMINT% = CUECOINT%(CTAMADRE$)
       !CuMa_Cli = CMINT%                   'CUENTA MADRE
       .Update
     End With
     '
19 Next U&
   Call ACRE_GES07.GENECOPROV    ' PARA ACTUALIZAR LAS LISTAS DE SELECCIóN
   '
99 On Error Resume Next
   BadaproVE.Close
   On Error GoTo 0
   '
End Sub

Sub IMPORPLA()
   '
   Call BLOQARCH("CUENTAS")
   Call BLOQARCH("DENOCUE")
   Call BLOQARCH("INVECUE")
   Call BLOQARCH("PLAGRAL")
   Call BLOQARCH("PLANCUEN")
   Call BLOQARCH("CUECON")
   Call BLOQARCH("ECUECON")
   Call BLOQARCH("ECUECOT")
   '
   JJ& = 0: J4& = 0: J5& = 0: J6& = 0
   URECU& = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("PLANCUEN")
   '
   FIN& = List4.ListCount
   List3.Clear
   '
   Call APERBASDAT("PLANCUEN")
20 On Error Resume Next
   Plancuen.Close
   On Error GoTo 0
   '
   On Error Resume Next
   Set Plancuen = Contable.OpenRecordset("PLANCUEN", dbOpenTable, dbDenyWrite)
   If Err Then
     If COMALTER%("No se Puede Abrir 'PLANCUEN'\en Modo Exclusivo.\\Reintentar\Cancelar") <> 2 Then GoTo 20
     GoTo 90
   End If
   '
   On Error Resume Next
   Plancuen.MoveLast
   NUREPLAN% = Plancuen.RecordCount
   Plancuen.MoveFirst
   On Error GoTo 0
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     ZZ$ = List4.List(U& - 1)
     AGRECUE% = 0
     COCUE$ = TRIM$(Mid$(ZZ$, 7, 12))
     DECUE$ = TRIM$(Mid$(ZZ$, 21, 48))
     '
     CUEIN% = 0
     For KU& = 1 To URECU&
       Z1$ = REGLEIDO$("CUENTAS", KU&)
       If TRIM$(Left$(Z1$, 12)) = COCUE$ Then
         CUEIN% = KU&
         Exit For
       End If
     Next KU&
       '
       If CUEIN% < 1 Or CUEIN% > URECU& Then
         URECU& = URECU& + 1
         CUEIN% = URECU&
         AGRECUE% = 1
       End If
       '
     RECUENTA& = CUEIN%
     TICU$ = Mid$(ZZ$, 71, 1)
     IBLE% = 1: If UCase$(Mid$(ZZ$, 75, 1)) = "N" Then IBLE% = 0
     STAT% = 1
     NIVEI% = XVALO(Left$(ZZ$, 4))
     '
     ' GRABAR EN CUENTAS .DAT
     If AGRECUE% < 1 Then
         Z1$ = REGLEIDO$("CUENTAS", RECUENTA&)
       Else
         Z1$ = REGBLAN$("CUENTAS")
     End If
     Call REPLA(Z1$, COCUE$, 1, 12)
     Call REPLA(Z1$, TICU$, 13, 1)
     Call GRAREG("CUENTAS", Z1$, RECUENTA&)
     '
     ' GRABAR EN DENOCUE.DAT
     For KU& = 1 + ULTREG&("DENOCUE") To RECUENTA&
       Call GRAREG("DENOCUE", Space$(127) + ".", KU&)
     Next KU&
     Z2$ = REGLEIDO$("DENOCUE", RECUENTA&)
     Call REPLA(Z2$, DECUE$, 1, 100)
     Call REPLA(Z2$, DECUE$, 101, 28)
     Call GRAREG("DENOCUE", Z2$, RECUENTA&)
     '
     ' GRABAR EN PLANCUEN.DAT
     Z3$ = REBLA$
     Call REPLA(Z3$, COCUE$, 1, 12)
     Call REPLA(Z3$, DECUE$, 13, 46)
     Call REPLA(Z3$, MKI$(CUEIN%), 59, 2)
     Call REPLA(Z3$, Chr$(STAT%), 61, 1)
     Call REPLA(Z3$, Chr$(NIVEI%), 62, 1)
     Call REPLA(Z3$, TICU$, 63, 1)
     Call REPLA(Z3$, Chr$(IBLE%), 64, 1)
     Call REPLA(Z3$, DECUE$, 129, 128)
     JJ& = JJ& + 1
     Call GRAREG("PLANCUEN", Z3$, JJ&)
     '
     ' GRABAR EN PLAGRAL.DAT
     Z4$ = MKI$(CUEIN%) + MKI$(NIVEI%)
     Call GRAREG("PLAGRAL", Z4$, JJ&)
     '
     ' GRABAR EN LIST3 PARA INVECUE
     TTXX$ = Space$(18)
     Call REPLA(TTXX$, COCUE$, 1, 12)
     Call REPLA(TTXX$, TRIM$(Str$(CUEIN%)), 13, 6)
     For KI% = 1 To List3.ListCount
       If List3.List(KI% - 1) > TTXX$ Then
         List3.AddItem TTXX$, KI% - 1
         GoTo 25
       End If
     Next KI%
     List3.AddItem TTXX$
     '
25   Z4$ = REGBLAN$("CUECON")
     Z5$ = REGBLAN$("ECUECON")
     '
     Call REPLA(Z5$, COCUE$, 1, 12)
     Call REPLA(Z5$, DECUE$, 13, 28)
     Call REPLA(Z5$, MKI$(CUEIN%), 41, 2)
     J5& = J5& + 1
     Call GRAREG("ECUECOT", Z5$, J5&)
     '
     If STAT = 1 Then
       If IBLE% = 1 Then
         '
         J6& = J6& + 1
         Call GRAREG("ECUECON", Z5$, J6&)
         '
         If TICU$ = "I" Or TICU$ = "G" Then
           Call REPLA(Z4$, MKI$(CUEIN%), 1, 2)
           Call REPLA(Z4$, DECUE$, 3, 30)
           Call REPLA(Z4$, COCUE$, 33, 12)
           Call REPLA(Z4$, COCUE$, 45, 12)
           J4& = J4& + 1
           Call GRAREG("CUECON", Z4$, J4&)
         End If
         '
       End If
     End If
     '
     On Error GoTo 0
     With Plancuen
       .AddNew   ' YA CONVERTIDO
       ON ERROR RESUME NEXT
       !CodiEmpr = Codiemp%
       ON ERROR GOTO 0
       !CodExCue = COCUE$
       !DENOCUE = DECUE$
       !CODINCUE = CUEIN%
       !STATCUE = STAT%
       !NIVAGRU = NIVEI%
       Select Case TICU$
            Case "D": !TipoCue = "Debe"
            Case "H": !TipoCue = "Haber"
            Case "G": !TipoCue = "Gasto"
            Case "I": !TipoCue = "Ingre"
       End Select
       !POSIMPU = IBLE%
       !Observa = " "
       !Salini = 0
       !Sudebe = 0
       !Suhabe = 0
       !Result = 0
       !Salfin = 0
       '!IndiAjus = INAJUS%
       '!MoneExtr = MONED%
       .Update
     End With
     '
   Next U&
   '
90 Call SETULTREG("PLANCUEN", JJ&)
   Call SETULTREG("PLAGRAL", JJ&)
   Call SETULTREG("CUECON", J4&)
   Call SETULTREG("ECUECOT", J5&)
   Call SETULTREG("ECUECON", J6&)
   Call CIERRARCH("CUENTAS")
   Call CIERRARCH("DENOCUE")
   Plancuen.Close
   '
   REBLA$ = REGBLAN$("INVECUE")
   J7& = 0
   For KI% = 1 To List3.ListCount
     TTXX$ = List3.List(KI% - 1)
     COCUE$ = Left$(TTXX$, 12)
     ICUE% = XVALO(Mid$(TTXX$, 13))
     Z7$ = REBLA$
     Call REPLA(Z7$, COCUE$, 1, 12)
     Call REPLA(Z7$, MKI$(ICUE%), 13, 2)
     J7& = J7& + 1
     Call GRAREG("INVECUE", Z7$, J7&)
   Next KI%
   Call SETULTREG("INVECUE", J7&)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Close
   '
   ABCT$ = LOADFILE$(LUDAT$ + "CUENTAS.DAT"): Call SAVEFILE(LUDAT$ + "CUENTAS.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "DENOCUE.DAT"): Call SAVEFILE(LUDAT$ + "DENOCUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "INVECUE.DAT"): Call SAVEFILE(LUDAT$ + "INVECUE.OLD", ABCT$)
   ABCT$ = LOADFILE$(LUDAT$ + "PLAGRAL.DAT"): Call SAVEFILE(LUDAT$ + "PLAGRAL.OLD", ABCT$)
   '
   Screen.MousePointer = 1
   '
End Sub
  '\\\OT-05-0207-WAR-25/04/05///
Sub IMPORNCLI()
    '
    Screen.MousePointer = 11
    '
    REBLA$ = REGBLAN$("REPLACOC")
    J& = 0
    For KI% = 1 To List4.ListCount
      TTXX$ = List4.List(KI% - 1)
      NCI1% = NUMECLI(Mid$(TTXX$, 1, 12)) 'Nº cliente
      NP1% = CODINT(Mid$(TTXX$, 14, 16)) 'Nº ARTICULO
      NO1$ = TRIM$(Mid$(TTXX$, 56, 28))  'Nº ORIGINAL
      Z$ = REBLA$
      Call REPLA(Z$, MKI(NCI1%), 1, 2)
      Call REPLA(Z$, MKI$(NP1%), 3, 2)
      Call REPLA(Z$, NO1$, 5, 28)
      J& = J& + 1
      Call GRAREG("REPLACOC", Z$, J&)
    Next KI%
    Call SETULTREG("REPLACOC", J&)
    '
    Screen.MousePointer = 1
    '
End Sub
'\\\OT-05-0207-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub PREIMPORCUECLI()
   '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
       U& = U& + 1
       Call TRACE(20, U&, URECORD&)
       '
       On Error Resume Next
       CodiClie$ = ""
       CodiClie$ = .Recordset.Fields(0)
       REFEMOVI$ = ""
       REFEMOVI$ = .Recordset.Fields(1)
       FECHEMIS$ = ""
       FECHEMIS$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
       FECHVENC$ = ""
       FECHVENC$ = FECHATEX$(CVINT(.Recordset.Fields(3)))
         If FECHANUM(FECHVENC$) < FECHANUM(FECHEMIS$) Then FECHVENC$ = FECHEMIS$
       CODICOMLAR$ = ""
       CODIMLAR$ = .Recordset.Fields(4)
       CODICOMCOR$ = ""
       CODIMCOR$ = .Recordset.Fields(5)
       If CODIMCOR$ = "" Then
         CODIMCOR$ = Left$(CODIMLAR$, 5) + Right$(CODIMLAR$, 4)
       End If
       IMORIGI# = 0: IMPENDE# = 0
       On Error Resume Next
       IMORIGI# = CDbl(.Recordset.Fields(6))
       IMPENDE# = CDbl(.Recordset.Fields(7))
       On Error GoTo 0
       '
2      NUCLIE% = NUMECLI%(CodiClie$)
       If NUCLIE% < 1 Then
         If COMALTER%("Cliente " + TRIM$(REFEMOVI$) + " No Registrado.\Desea Darlo de Alta ?\\Si\No") = 1 Then
           Call APERBASDAT("CLIEN")
           NROPRO& = 0
           Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli >= " & PRONUCLI& & " ORDER BY Nume_Cli DESC;", dbOpenSnapshot)
           On Error Resume Next
           '
3          NROPRO& = NROPRO& + 1
           Tranclie.FindFirst "Nume_Cli = " & NROPRO&
           If Tranclie.NoMatch = False Then GoTo 3
           '
           NCLIE% = NROPRO&
           Deudor12.Close
           Call APERBASDAT("CLIEN")
           With Deudor12
               .AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
               !CodiEmpr = Codiemp%
               ON ERROR GOTO 0
               !Nume_Cli = TRIM$(Str$(NCLIE%))
               !Raso_Cli = REFEMOVI$
               !Codi_Cli = CodiClie$
               !Cuit_Cli = "00-00000000-0"
               !Piva_Cli = 1
               !Stat_Cli = 1
               '
               On Error Resume Next
               .Update
               If Err Then
                  Call MENSERR(24, "Alta no Procesada Correctamente")
               End If
           End With
           GoTo 2
         End If
         '
         If CodiClie$ <> CODICLIA$ Then
           CODICLIA$ = CodiClie$
           MERROX$ = "Falta Cliente '" + TRIM$(CodiClie$) + " - " + REFEMOVI$ + "' en B.Datos de Clientes"
           List1.AddItem MERROX$
         End If
         GoTo 19
       End If
       '
       ' VALIDAR QUE NO IMPORTE DOS VECES EL MISMO COMPROBANTE
       '
       TTXX$ = Space$(254)
       Call REPLA(TTXX$, CodiClie$, 1, 12)
       Call REPLA(TTXX$, FORMATNUM$(NUCLIE%, "F6,0"), 13, 6)
       Call REPLA(TTXX$, REFEMOVI$, 21, 30)
       Call REPLA(TTXX$, FECHEMIS$, 53, 8)
       Call REPLA(TTXX$, FECHVENC$, 63, 8)
       Call REPLA(TTXX$, CODIMLAR$, 75, 20)
       Call REPLA(TTXX$, CODIMCOR$, 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMORIGI#, "F12.2"), 107, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMPENDE#, "F12.2"), 119, 12)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub PREIMPORCUEPRO()
   '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   Call APERBASDAT("CABAN")
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   '
   INI& = 1: FINI& = URECORD&
   If URECORD& > 30000 Then
     TTXX$ = "Imposible Importar en Secuencia Unica.\Elija Rango de Importacion:\\1 - 30000"
     FINI& = 30000
5    INI& = FINI& + 1
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
     TTXX$ = TTXX$ + "\" + TRIM$(Str$(INI&)) + " - " + TRIM$(Str$(FINI&))
     If FINI& < URECORD& Then GoTo 5
     '
     OPX& = COMALTER%(TTXX$)
     If OPX& < 1 Then GoTo 20
     INI& = 1 + 30000 * (OPX& - 1)
     FINI& = INI& + 29999
     If FINI& > URECORD& Then FINI& = URECORD&
   End If
   '
   Screen.MousePointer = 11
   With Data1
     .Recordset.MoveFirst
     U& = 0
     Do While Not .Recordset.EOF
       U& = U& + 1
       Call TRACE(20, U&, URECORD&)
       If U& < INI& Or (U& > FINI& And FINI& < URECORD&) Then GoTo 19
       '
       On Error Resume Next
       CodiClie$ = ""
       CodiClie$ = .Recordset.Fields(0)
       REFEMOVI$ = ""
       REFEMOVI$ = .Recordset.Fields(1)
       FECHEMIS$ = ""
       FECHEMIS$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
       FECHVENC$ = ""
       FECHVENC$ = FECHATEX$(CVINT(.Recordset.Fields(3)))
         If FECHANUM(FECHVENC$) < FECHANUM(FECHEMIS$) Then FECHVENC$ = FECHEMIS$
       CODIMLAR$ = ""
       CODIMLAR$ = .Recordset.Fields(4)
       CODIMCOR$ = ""
       CODIMCOR$ = .Recordset.Fields(5)
       If CODIMCOR$ = "" Then
         CODIMCOR$ = Left$(CODIMLAR$, 5) + Right$(CODIMLAR$, 4)
       End If
       IMORIGI# = 0: IMPENDE# = 0
       On Error Resume Next
       IMORIGI# = CDbl(.Recordset.Fields(6))
       IMPENDE# = CDbl(.Recordset.Fields(7))
       On Error GoTo 0
       '
2      NUCLIE% = NUMEPRO%(CodiClie$)
       If NUCLIE% < 1 Then
         If COMALTER%("Proveedor " + TRIM$(REFEMOVI$) + " No Registrado.\Desea Darlo de Alta ?\\Si\No") = 1 Then
           Call APERBASDAT("PROVE")
           NROPRO& = 0
           Set Tranclie = Badaprov.OpenRecordset("SELECT Nume_Cli FROM Proved12 WHERE Nume_Cli >= " & PRONUCLI& & " ORDER BY Nume_Cli DESC;", dbOpenSnapshot)
           On Error Resume Next
           '
3          NROPRO& = NROPRO& + 1
           Tranclie.FindFirst "Nume_Cli = " & NROPRO&
           If Tranclie.NoMatch = False Then GoTo 3
           '
           NCLIE% = NROPRO&
           Proved12.Close
           Call APERBASDAT("PROVE")
           With Proved12
               .AddNew   ' YA CONVERTIDO
               ON ERROR RESUME NEXT
               !CodiEmpr = Codiemp%
               ON ERROR GOTO 0
               !Nume_Cli = TRIM$(Str$(NCLIE%))
               !Raso_Cli = REFEMOVI$
               !Codi_Cli = CodiClie$
               !Cuit_Cli = "00-00000000-0"
               !Piva_Cli = 1
               !Stat_Cli = 1
               !Tcpr_Cli = 0
               '
               On Error Resume Next
               .Update
               If Err Then
                  Call MENSERR(24, "Alta no Procesada Correctamente")
               End If
           End With
           GoTo 2
         End If
         '
         If CodiClie$ <> CODICLIA$ Then
           CODICLIA$ = CodiClie$
           MERROX$ = "Falta Proveedor '" + TRIM$(CodiClie$) + " - " + REFEMOVI$ + "' en B.Datos Proveedores"
           List1.AddItem MERROX$
         End If
         GoTo 19
       End If
       '
       TTXX$ = Space$(254)
       Call REPLA(TTXX$, CodiClie$, 1, 12)
       Call REPLA(TTXX$, FORMATNUM$(NUCLIE%, "F6,0"), 13, 6)
       Call REPLA(TTXX$, REFEMOVI$, 21, 30)
       Call REPLA(TTXX$, FECHEMIS$, 53, 8)
       Call REPLA(TTXX$, FECHVENC$, 63, 8)
       Call REPLA(TTXX$, CODIMLAR$, 75, 20)
       Call REPLA(TTXX$, CODIMCOR$, 95, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMORIGI#, "F12.2"), 107, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMPENDE#, "F12.2"), 119, 12)
       '
       List4.AddItem TTXX$
       Label6 = TRIM$(Str$(List4.ListCount))
       If Len(RTrim(TTXX$ + "*")) > Len(TEPRUE) Then
         TEPRUE = RTrim(TTXX$ + "*")
       End If
       On Error Resume Next
       List4.TopIndex = List4.ListCount - 7
       On Error GoTo 0
       DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0429-WAR-28/06/05///
Sub PREIMPORCARCHE()
  '
   If DBGrid1.Columns.Count < 8 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
   '
   Screen.MousePointer = 11
   DoEvents
   Call APERBASDAT("CABAN")
   Screen.MousePointer = 11
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "WHERE TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND OpciMovim = 'CH' "
   Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   '
   URECORD& = Data1.Recordset.RecordCount
   ATX% = 0
   hrow = DBGrid1.RowHeight   ' PARA QUE NO CORTE ANTES DEL FINAL
   DBGrid1.RowHeight = 10
   '
   For U& = 1 To URECORD&
     Call TRACE(20, U&, URECORD&)
     TTXX$ = Space$(290)
     On Error Resume Next
     DBGrid1.Row = U& - 1
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Completar Pre-Importación.\Demasiados Registros en Tabla.")
       GoTo 20
     End If
     On Error GoTo 0
     '
     DBGrid1.Col = 0                     ' BANCO
     BANC$ = TRIM$(DBGrid1.TEXT)
     If Len(BANC$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Banco"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, BANC$, 1, 18)
     '
     DBGrid1.Col = 1
     NUCHE$ = TRIM$(DBGrid1.TEXT)   ' NUMERO DE CHEQUE
     If Len(NUCHE$) < 1 Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Numero de Cheque no Valido"
       List1.AddItem MERROX$
       GoTo 19
     End If
     Call REPLA(TTXX$, NUCHE$, 21, 12)
     '
     DBGrid1.Col = 2                     'FECHA EMISION
     FECHEMIS$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHEMIS$, 35, 8)
     '
     DBGrid1.Col = 3                     'FECHA VTO.
     FECHEVEN$ = FECHATEX$(CVINT(TRIM(DBGrid1.TEXT)))
     Call REPLA(TTXX$, FECHEVEN$, 45, 8)
     '
     DBGrid1.Col = 4                     'IMPORTE ORIGINAL
     IMPO1# = CDbl(TRIM$(DBGrid1.TEXT))
     IMPO$ = FORMATNUM(IMPO1#, "F13.2")
     IMPO2$ = TRIM$(IMPO$)
     Call REPLA(TTXX$, IMPO$, 55, 13)
     '
     DBGrid1.Col = 5                     ' DOCUMENTO ORIGEN LARGO
     DOCU$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOCU$, 71, 18)
     '
     DBGrid1.Col = 6                     ' DOCUMENTO ORIGEN CORTO
     DOCUC$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, DOCUC$, 91, 12)
     '
     DBGrid1.Col = 7                     ' CLIENTE
     CLI1$ = TRIM$(DBGrid1.TEXT)
     Call REPLA(TTXX$, CLI1$, 105, 30)
       '
       IDEVA$ = Left$(TTXX$, 32)
       IDEVA$ = IDEVA$ + Mid$(TTXX$, 45, 8)
     SQLX$ = "DescriMov = '" & IDEVA$ & "' "
     SQLX$ = SQLX$ & "AND ABS(IDebeComp - " & IMPO2$ & ") < 0.005 "
     CuecoTemp.FindFirst SQLX$
     If CuecoTemp.NoMatch = False Then
       MERROX$ = "Registro" + Str$(U& - 1) + " - Cheque " + IDEVA$ + " - $ " + TRIM$(IMPO$) + " Existente en Cartera"
       List1.AddItem MERROX$
       GoTo 19
     End If
     '
17   List4.AddItem TTXX$
     Label6 = TRIM$(Str$(List4.ListCount))
     If Len(RTrim$(TTXX$ + "*")) > Len(TEPRUE) Then
       TEPRUE = RTrim$(TTXX$ + "*")
     End If
     On Error Resume Next
     List4.TopIndex = List4.ListCount - 7
     On Error GoTo 0
     DoEvents
19 Next U&
   '
20 DBGrid1.RowHeight = hrow
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0429-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub IMPORCUECLI()
   '
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 95, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 75, 20))
     CodiEmp% = 0
     CodiSuc% = 0
     NumeClie% = XVALO(Mid$(TTXX$, 13, 6))
     NumeProv% = 0
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
       NC% = XVALO(Mid$(TTXX$, 13, 6))
     Referen$ = RASOCLI$(NC%)
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     CuentaCo% = CUECOCLI%(NC%)
     TiMovi$ = "FVEN"
       '
       OPMOV$ = UCase$(TRIM$(Mid$(TTXX$, 95, 2)))
       Select Case (OPMOV$)
        Case "FC"
           TICOMX$ = "FC": OPMOVIX$ = "FAC"
        Case "RB"
           TICOMX$ = "RB": OPMOVIX$ = "CCORR"
        Case "ND"
           TICOMX$ = "ND": OPMOVIX$ = "DEB"
        Case "NC"
       End Select
       '
     OpMovi$ = OPMOV$
     VaMovi$ = Mid$(CoComLar$, 4, 1)
     Descrip$ = CoComLar$
     NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
     NumeCom = XVALO(Mid$(CoComLar$, 11))
     StatInit = 0
     ImpoDeb = XVALO(Mid$(TTXX$, 107, 12))
     Observa$ = RASOCLI$(NC%)
     Call GRACABAN
     '
     IBruTot# = XVALO(Mid$(TTXX$, 107, 12))
     IIvaTot# = 0
     NuorVen% = 0
     FeVenc = CVDAT(FECHANUM(Mid$(TTXX$, 63, 8)))
     ImVenc# = XVALO(Mid$(TTXX$, 119, 12))
     Call GRADEUPEN
     '
   Next U&
   CajaBanc.Close
   SaDeuPen.Close
   '
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///
'\\\OT-05-0272-WAR-09/05/05///
Sub IMPORCUEPRO()
 '
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   NOCLOSE% = 1
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 95, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 75, 20))
     CodiEmp% = 0
     CodiSuc% = 0
     NumeClie% = 0
     NumeProv% = XVALO(Mid$(TTXX$, 13, 6))
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
       NC% = XVALO(Mid$(TTXX$, 13, 6))
     Referen$ = RASOCLI$(NC% * (-1))
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 53, 8)))
     CuentaCo% = CUECOCLI%(NC% * (-1))
     TiMovi$ = "FCOM"
       '
       OPMOV$ = UCase$(TRIM$(Mid$(TTXX$, 95, 2)))
       Select Case (OPMOV$)
        Case "FC"
           TiMovi$ = "FCOM"
           OPMOVIX$ = "FC"
        Case "OP"
           TiMovi$ = "OPAG"
           OPMOVIX$ = "OP"
        Case "ND"
           OPMOVIX$ = "ND"
        Case "NC"
           OPMOVIX$ = "NC"
       End Select
       '
     OpMovi$ = OPMOVIX$
     VaMovi$ = Mid$(CoComLar$, 4, 1)
     Descrip$ = CoComLar$
     NumeTal = XVALO(Mid$(CoComLar$, 6, 4))
     NumeCom = XVALO(Mid$(CoComLar$, 11))
     StatInit = 0
     ImpoDeb = XVALO(Mid$(TTXX$, 107, 12))
     Observa$ = RASOCLI$(NC% * (-1))
     Call GRACABAN
     '
     IBruTot# = Abs(XVALO(Mid$(TTXX$, 107, 12)))
     IIvaTot# = 0
     NuorVen% = 0
     FeVenc = CVDAT(FECHANUM(Mid$(TTXX$, 63, 8)))
     ImVenc# = (XVALO(Mid$(TTXX$, 119, 12))) * (-1)
     Call GRACREPEN
     '
   Next U&
   CajaBanc.Close
   Sacrepen.Close
   '
   Screen.MousePointer = 1
   '
End Sub
'\\\OT-05-0272-WAR-FIN///

Sub IMPORCARCHE()
   '
   CUCACHE% = CUEICARCH%
   If CUCACHE% < 1 Then
     Call MENSERR(24, "Importación Imposible.\Falta Definir Cuanta Contable\Asociada a la Cartera de Valores")
     CANCELIMP% = 1
     Exit Sub
   End If
   '
   Call SELECHO("ECUECON/Seleccione Cuenta de Contrapartida")
   CONTRAPA$ = TRIM$(VALACT1$("ECUECON"))
   CUCONTRA% = CUEINT%(CONTRAPA$)
   If CONTRAPA$ = "" Or CUCONTRA% < 1 Then
     Call MENSERR(24, "Importación Cancelada.\Debe Definir una Cuenta de Contra Partida\Para el Asiento Implícito")
     CANCELIMP% = 1
     Exit Sub
   End If
   
   FIN& = List4.ListCount
   Call APERBASDAT("CABAN")
   '
   TOTAVAL# = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Label6 = TRIM$(Str$(U&)): DoEvents
     TTXX$ = List4.List(U& - 1)
     '
30   CoComCor$ = TRIM$(Mid$(TTXX$, 91, 12))
     CoComLar$ = TRIM$(Mid$(TTXX$, 71, 18))
     CodiEmp% = 0
     CodiSuc% = SUCUI%
     NumeClie% = 0
     NumeProv% = 0
     FechEmi = CVDAT(FECHANUM(Mid$(TTXX$, 35, 8)))
     Referen$ = Mid$(TTXX$, 105, 30)
     FeContab = CVDAT(FECHANUM(Mid$(TTXX$, 35, 8)))
     FeAcredi = CVDAT(FECHANUM(Mid$(TTXX$, 45, 8)))
     CuentaCo% = CUCACHE%
     TiMovi$ = "CHRECI"
     OpMovi$ = "CH"
     VaMovi$ = ""
       IDEVA$ = Left$(TTXX$, 32)
       IDEVA$ = IDEVA$ + Mid$(TTXX$, 45, 8)
     Descrip$ = IDEVA$
     NumeTal = 0
     NumeCom = XVALO(Mid$(TTXX$, 94, 9))
     StatInit = 1
     ImpoDeb# = XVALO(Mid$(TTXX$, 55, 13))
       TOTAVAL# = TOTAVAL# + ImpoDeb#
     Observa$ = ""
     '
     Call GRACABAN
     '
   Next U&
   '
   FEDEP% = HOY(HOS$)
   FEX = FEDEP%
   FETX$ = FECHATEX$(FEX)
   NRECIB$ = "IP-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-00"
   CoComCor$ = NRECIB$
   CoComLar$ = NRECIB$
   CodiEmp% = 0
   CodiSuc% = SUCUI%
   NumeClie% = 0
   NumeProv% = 0
   FechEmi = CVDAT(FEDEP%)
   Referen$ = "Importacion Cartera Cheques " + NRECIB$
   FeContab = CVDAT(FEDEP%)
   FeAcredi = CVDAT(FEDEP%)
   CuentaCo% = CUCONTRA%
   TiMovi$ = "IMPOR"
   OpMovi$ = "CARCHE"
   VaMovi$ = ""
   Descrip$ = "Importacion Cartera Cheques " + NRECIB$
   NumeTal = 0
   NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
   StatInit = 0
   ImpoDeb# = (-1) * TOTAVAL#
   Call GRACABAN
   '
   Screen.MousePointer = 1
   '
End Sub

Sub PREIMPORMOVSTK()
   '
   '\\\OT-5-469-FG-05/07/05///
   If DBGrid1.Columns.Count < 10 Then
     Call MENSERR(24, "Formato de Archivo Incorrecto.\Faltan Columnas de Datos.")
     GoTo 20
   End If
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
        TTXX$ = Space$(254)
        codex$ = ""
        On Error Resume Next
        codex$ = TRIM$(.Recordset.Fields(0))   ' TRIM$(DBGrid1.TEXT)
        On Error GoTo 0
        If codex$ = "" Or Len(codex$) < 3 Or Len(codex$) > 15 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo Inexistente o No Valido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        '
12      Call REPLA(TTXX$, codex$, 1, 16)
        '
13      CDXI% = CODINT%(codex$)
        If CDXI% < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Codigo '" + codex$ + "' Inexistente o No Valido"
          If COMALTER%(MERROX$ + "\Desea Dar de Alta Este Código ?\\Si\No") = 1 Then
            VDEF$ = String$(128, 0)
            Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
            Call REPLA(VDEF$, "U.", 109, 2)
            Call REPLA(VDEF$, Chr$(1), 93, 1)
            Call REPLA(VDEF$, codex$, 77, 16)
            DEDEX$ = ""
            On Error Resume Next
            DEDEX$ = TRIM$(.Recordset.Fields(1))
            On Error GoTo 0
            Call REPLA(VDEF$, DEDEX$, 5, 64)
            Call AMAS_GES07.ALTACODIGO(VDEF$, OKX%)
            If OKX% = 1 Then GoTo 13
          End If
          '
          List1.AddItem MERROX$
          GoTo 19
        End If
        VALO$ = AJUSTD$(Str$(CDXI%), 6)
        Call REPLA(TTXX$, VALO$, 17, 6)
        '
        DEDEX$ = ""
        On Error Resume Next
        DEDEX$ = TRIM$(.Recordset.Fields(1))
        On Error GoTo 0
        If DEDEX$ = "" Then DEDEX$ = DESCRIT0$(CDXI%)
        Call REPLA(TTXX$, DEDEX$, 25, 64)
        '
        On Error Resume Next
        FECHMOVX$ = FECHATEX$(CVINT(.Recordset.Fields(2)))
        On Error GoTo 0
        If FECHANUM(FECHMOVX$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Fecha de Movimiento No Válida"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, FECHMOVX$, 91, 8)
        '
        DEPX = 1 ' dEPóSITO
        On Error Resume Next
        DEPX = CDbl(.Recordset.Fields(3))
        On Error GoTo 0
        If DEPX < 1 Or DEPX > 255 Then DEPX = 1
        VALO$ = FORMATNUM$(DEPX, "F4.0")
        Call REPLA(TTXX$, VALO$, 101, 4)
        
        CANTI = 0 ' CANTIDAD
        On Error Resume Next
        CANTI = CDbl(.Recordset.Fields(4))
        On Error GoTo 0
        CANTIX$ = FORMATNUM$(CANTI, "F12.1")
        Call REPLA(TTXX$, CANTIX$, 105, 12)
        CANTIX$ = TRIM$(CANTIX$)
        '
        CODMOVI$ = ""
        On Error Resume Next
        CODMOVI$ = TRIM$(.Recordset.Fields(5))
        On Error GoTo 0
        If Len(CODMOVI$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Código de Movimiento No Válido"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, CODMOVI$, 119, 2)
        '
        DOCPRI$ = ""
        On Error Resume Next
        DOCPRI$ = TRIM$(.Recordset.Fields(6))
        On Error GoTo 0
        If Len(DOCPRI$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Documento Primario"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DOCPRI$, 123, 12)
        '
        DOCSEC$ = ""
        On Error Resume Next
        DOCSEC$ = TRIM$(.Recordset.Fields(7))
        On Error GoTo 0
        Call REPLA(TTXX$, DOCSEC$, 137, 12)
        
        VALUNI = 0
        On Error Resume Next
        VALUNI = CDbl(.Recordset.Fields(8))
        On Error GoTo 0
        VALO$ = FORMATNUM$(VALUNI, "F14.4")
        Call REPLA(TTXX$, VALO$, 149, 14)
        '
        MONX = 1 ' MONEDA
        On Error Resume Next
        MONX = CDbl(.Recordset.Fields(9))
        On Error GoTo 0
        If MONX < 1 Or MONX > 4 Then MONX = 1
        VALO$ = FORMATNUM$(MONX, "F4.0")
        Call REPLA(TTXX$, VALO$, 163, 4)
        '
        CLIX$ = ""
        NPROV% = 0
        On Error Resume Next
        CLIX$ = TRIM$(CDbl(.Recordset.Fields(10)))
        On Error GoTo 0
        If CLIX$ = "" Then
            NCLIE% = 0
          Else
            NCLIE% = NUMECLI%(CLIX$)
            If NCLIE% < 1 Then
              NCLIE% = XVALO(CLIX$)
            End If
        End If
        '
        If NCLIE% < 1 Then
          CLIX$ = ""
          NPROV% = 0
          On Error Resume Next
          CLIX$ = TRIM$(CDbl(.Recordset.Fields(11)))
          On Error GoTo 0
          If CLIX$ = "" Then
              NPROV% = 0
            Else
              NPROV% = NUMEPRO%(CLIX$)
              If NPROV% < 1 Then
                NPROV% = XVALO(CLIX$)
              End If
          End If
        End If
        '
        VALO$ = FORMATNUM$(NCLIE%, "F6.0")
        Call REPLA(TTXX$, VALO$, 167, 6)
        '
        VALO$ = FORMATNUM$(NPROV%, "F6.0")
        Call REPLA(TTXX$, VALO$, 173, 6)
        '
        DETALLE$ = ""
        On Error Resume Next
        DETALLE$ = TRIM$(.Recordset.Fields(12))
        On Error GoTo 0
        If Len(DETALLE$) < 1 Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Falta Detalle / Referencia del Movimiento"
          List1.AddItem MERROX$
          GoTo 19
        End If
        Call REPLA(TTXX$, DETALLE$, 181, 32)
        '
        FEXI% = FECHANUM(FECHMOVX$)
        FEMOVD = Int(CDbl(CVDAT(FEXI%)))
        If Abs(XVALO(CANTIX$)) < 0.01 Then CANTIX$ = "0"
        Call APERBASDAT("STOCKS")
        SQLX$ = "Cod_Inte = " & CDXI% & " "
        SQLX$ = SQLX$ + "AND DepoMovi = " & DEPX & " "
        SQLX$ = SQLX$ + "AND DoPriCor = '" & DOCPRI$ & "' "
        SQLX$ = SQLX$ + "AND ABS((CantIngre - CantSalid) - " & CANTIX$ & ") < 0.05 "
        SQLX$ = SQLX$ + "AND INT(CDBL(FechMov)) = " & FEMOVD
        '
        Movisto.FindFirst SQLX$
        If Movisto.NoMatch = False Then
          MERROX$ = "Registro" + Str$(U& - 1) + " - Movimiento Pre-Existente en Base de Datos."
          List1.AddItem MERROX$
          GoTo 19
        End If
        
        List4.AddItem TTXX$
        Label6 = TRIM$(Str$(List4.ListCount))
        If Len(RTrim(TTXX$) + "*") > Len(TEPRUE) Then
          TEPRUE = RTrim(TTXX$) + "*"
        End If
        On Error Resume Next
        List4.TopIndex = List4.ListCount - 7
        On Error GoTo 0
        DoEvents
19   .Recordset.MoveNext
     Loop
   End With
   '
20 Screen.MousePointer = 1
   '
   '\\\OT-5-469-FG-FIN///
End Sub

Sub IMPORMOVSTK()
   '
   Screen.MousePointer = 11
   FIN& = List4.ListCount
   Call APERBASDAT("STOCKS")
   NOCLOSE% = 1
   '
   NMS% = NUMAS%
   'REBLA$ = REGBLAN$("MASTER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TTXX$ = List4.List(U& - 1)
     CODIT$ = TRIM$(Left$(TTXX$, 16))
     CIXI% = XVALO(Mid$(TTXX$, 17, 6))
     FEMO% = FECHANUM(Mid$(TTXX$, 91, 8))
     LOTE$ = ""
     CMOV$ = TRIM$(Mid$(TTXX$, 119, 2))
     DOPRILA$ = TRIM$(Mid$(TTXX$, 123, 12)): DOPRICO$ = DOPRILA$
     DOSELA$ = TRIM$(Mid$(TTXX$, 137, 12)): DOSECO$ = DOSELA$
     REFE$ = TRIM$(Mid$(TTXX$, 181, 32))
     NUORIT% = 0
     VUNI# = XVALO(Mid$(TTXX$, 149, 14))
     MONEII% = XVALO(Mid$(TTXX$, 163, 4))
     NC% = XVALO(Mid$(TTXX$, 167, 6)) - XVALO(Mid$(TTXX$, 173, 6))
     DEPO% = XVALO(Mid$(TTXX$, 101, 4))
     CANTI = XVALO(Mid$(TTXX$, 105, 12))
     '
     Call MOVISTOK(FEMO%, CIXI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI)
     '
   Next U&
   '
   Movisto.Close
   Screen.MousePointer = 1
   '
End Sub
'

